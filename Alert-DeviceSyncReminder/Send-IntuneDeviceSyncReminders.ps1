# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
    Identifies Intune devices that haven't synced in a specified period and sends email notifications to their primary users.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all managed devices from Intune, identifies those that haven't synced within the specified
    threshold period, and sends email notifications to the primary users of those devices with instructions
    to sync their devices.
    
.PARAMETER DaysSinceLastSync
    The number of days to use as a threshold for determining "stale" devices that need to sync.
    Default is 7 days.
    
.PARAMETER EmailSender
    The email address that will appear as the sender of the notification emails.
    
.PARAMETER ExcludedDeviceCategories
    Optional. An array of device categories to exclude from the notification process.
    
.PARAMETER MaxEmailsPerRun
    Optional. Maximum number of emails to send in a single runbook execution. Default is 100.
    
.PARAMETER BatchSize
    Optional. Number of devices to process in each batch. Default is 50.
    
.PARAMETER BatchDelaySeconds
    Optional. Number of seconds to wait between processing batches. Default is 10.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually send emails.
    
.PARAMETER TestEmailAddress
    Optional. If specified, all emails will be sent to this address instead of the actual device users.
    Use this for testing purposes.
    
.PARAMETER LogoUrl
    Optional. URL to the company logo to use in the email template. Default is a placeholder.

.NOTES
    File Name: Send-IntuneDeviceSyncReminders.ps1
    Author: Ryan Schultz
    Version: 1.1
    Created: 2025-04-07
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$DaysSinceLastSync = 7,
    
    [Parameter(Mandatory = $true)]
    [string]$EmailSender,
    
    [Parameter(Mandatory = $false)]
    [string[]]$ExcludedDeviceCategories = @(),
    
    [Parameter(Mandatory = $false)]
    [int]$MaxEmailsPerRun = 100,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchDelaySeconds = 10,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [string]$TestEmailAddress = "",
    
    [Parameter(Mandatory = $false)]
    [string]$LogoUrl = "https://img.logoipsum.com/363.svg",

    [Parameter(Mandatory = $false)]
    [string]$SupportEmail = "",
    
    [Parameter(Mandatory = $false)]
    [string]$SupportPhone = "",

    [Parameter(Mandatory = $false)]
    [string]$ITDepartmentEmail = ""
)

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { 
            Write-Error $Message
            Write-Verbose $LogMessage -Verbose
        }
        "WARNING" { 
            Write-Warning $Message 
            Write-Verbose $LogMessage -Verbose
        }
        "WHATIF" { 
            Write-Verbose "[WHATIF] $Message" -Verbose
        }
        default { 
            Write-Verbose $LogMessage -Verbose
        }
    }
}

# Connect to Azure using the managed identity
function Get-MsGraphToken {
    try {
        Write-Log "Authenticating with Managed Identity..."
        Connect-AzAccount -Identity | Out-Null

        $tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

        if ($tokenObj.Token -is [System.Security.SecureString]) {
            Write-Log "Token is SecureString, converting to plain text..."
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            )
        } else {
            Write-Log "Token is plain string, no conversion needed."
            $token = $tokenObj.Token
        }

        if (-not [string]::IsNullOrEmpty($token)) {
            Write-Log "Token acquired successfully."
            return $token
        } else {
            throw "Token was empty."
        }
    }
    catch {
        Write-Error "Failed to acquire Microsoft Graph token using Managed Identity: $_"
        throw
    }
}

function Invoke-MsGraphRequestWithRetry {
    param (
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $params = @{
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        Method      = $Method
        ContentType = $ContentType
    }
    
    if ($null -ne $Body -and $Method -ne "GET") {
        $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
    }
    
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            
            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $retryCount -lt $MaxRetries) {
                $retryAfter = $backoffSeconds
                if ($_.Exception.Response -ne $null -and $_.Exception.Response.Headers -ne $null) {
                    $retryAfterHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq "Retry-After" }
                    if ($retryAfterHeader) {
                        $retryAfter = [int]$retryAfterHeader.Value[0]
                    }
                }
                
                if ($statusCode -eq 429) {
                    Write-Log "Request throttled by Graph API (429). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                else {
                    Write-Log "Server error (5xx). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                
                Start-Sleep -Seconds $retryAfter
                
                $retryCount++
                $backoffSeconds = $backoffSeconds * 2
            }
            else {
                Write-Log "Graph API request failed with status code $statusCode`: $_" -Type "ERROR"
                throw $_
            }
        }
    }
}

function Get-IntuneDevices {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving Intune devices..."
        $filter = "operatingSystem eq 'Windows' or operatingSystem eq 'iOS' or operatingSystem eq 'Android' or operatingSystem eq 'Linux'"
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$filter&`$select=id,deviceName,managedDeviceOwnerType,deviceType,operatingSystem,osVersion,complianceState,lastSyncDateTime,emailAddress,userPrincipalName,serialNumber,model,manufacturer,enrolledDateTime,userDisplayName,deviceCategoryDisplayName"
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $devices += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of devices..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $devices += $response.value
        }
        
        $osCounts = @{}
        $devices | ForEach-Object {
            $os = $_.operatingSystem
            if (-not $osCounts.ContainsKey($os)) {
                $osCounts[$os] = 0
            }
            $osCounts[$os]++
        }
        
        $osCountsString = $osCounts.GetEnumerator() | ForEach-Object {
            "$($_.Value) $($_.Key)"
        } | Join-String -Separator ", "
        
        Write-Log "Retrieved $($devices.Count) devices from Intune ($osCountsString)"
        
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve Intune devices: $_" -Type "ERROR"
        throw "Failed to retrieve Intune devices: $_"
    }
}

function Get-DevicePrimaryUser {
    param (
        [string]$Token,
        [string]$DeviceId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving primary user for device $DeviceId..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/users?`$select=id,displayName,mail,userPrincipalName,givenName"
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        if ($response.value.Count -gt 0) {
            return $response.value[0]
        }
        else {
            Write-Log "No primary user found for device $DeviceId" -Type "WARNING"
            return $null
        }
    }
    catch {
        Write-Log "Failed to retrieve primary user for device $DeviceId`: $_" -Type "ERROR"
        return $null
    }
}

function Send-GraphMailMessage {
    param (
        [string]$Token,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$HtmlBody,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would send email from $From to $To with subject: $Subject" -Type "WHATIF"
            return $true
        }
        
        Write-Log "Sending email to $To using Microsoft Graph API..."
        
        $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        
        $messageBody = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "HTML"
                    content = $HtmlBody
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $To
                        }
                    }
                )
            }
            saveToSentItems = $true
        }
        
        Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $messageBody -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Email sent to $To successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send email to $To`: $_" -Type "ERROR"
        return $false
    }
}

function Send-SyncReminderEmail {
    param (
        [string]$Token,
        [string]$To,
        [string]$Username,
        [string]$DeviceName,
        [string]$LastSyncTime,
        [string]$From,
        [string]$LogoUrl,
        [string]$SupportEmail = "it@example.com",
        [string]$SupportPhone = "(555) 123-4567",
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        $subject = "Action Required: Your device $DeviceName needs to sync with Intune"
        
        # HTML Email Body
        $body = @"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Device Sync Required</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 0;">
    <div style="max-width: 800px; margin: 0 auto; border: 1px solid #ddd;">
        <!-- Header Section -->
        <div style="background-color: #0078D4; color: white; padding: 20px; text-align: center;">
            <img src="$LogoUrl" alt="Company Logo" width="150" style="display: inline-block;">
            <h2 style="margin-top: 15px; margin-bottom: 0;">Action Required: Device Sync Overdue</h2>
        </div>
        
        <!-- Content Section -->
        <div style="padding: 20px;">
            <!-- Intro -->
            <div style="margin-bottom: 20px;">
                <p>Hello $Username,</p>
                <p>Your device, <strong>$DeviceName</strong>, has not synced with Intune since <strong>$LastSyncTime</strong>. To ensure your device remains compliant and secure, please perform the following steps:</p>
            </div>
            
            <!-- Sync Section -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 10px;">Sync Your Device</h3>
                
                <!-- Windows Instructions -->
                <!-- WINDOWS_INSTRUCTIONS_START -->
                <h4 style="margin-top: 20px; margin-bottom: 10px;">For Windows Devices:</h4>
                <ol style="margin-top: 0; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Ensure your device is powered on and connected to the internet.</li>
                    <li style="margin-bottom: 8px;">Click on the Start menu and search for "Company Portal".</li>
                    <li style="margin-bottom: 8px;">Open the <strong>Company Portal</strong> app.</li>
                    <li style="margin-bottom: 8px;">Select the Settings icon in the bottom left corner.</li>
                    <li style="margin-bottom: 8px;">Click the <strong>Sync</strong> button to initiate a sync.</li>
                </ol>
                <!-- WINDOWS_INSTRUCTIONS_END -->
                
                <!-- iOS Instructions -->
                <!-- IOS_INSTRUCTIONS_START -->
                <h4 style="margin-top: 20px; margin-bottom: 10px;">For iOS/iPadOS Devices:</h4>
                <ol style="margin-top: 0; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Ensure your device is powered on and connected to the internet (Wi-Fi or cellular).</li>
                    <li style="margin-bottom: 8px;">Open the <strong>Company Portal</strong> app on your device.</li>
                    <li style="margin-bottom: 8px;">Tap on <strong>Devices</strong> at the bottom of the screen.</li>
                    <li style="margin-bottom: 8px;">Select your device from the list.</li>
                    <li style="margin-bottom: 8px;">Tap on <strong>Check Status</strong> or <strong>Sync</strong> to initiate a sync.</li>
                </ol>
                <!-- IOS_INSTRUCTIONS_END -->
                
                <!-- Android Instructions -->
                <!-- ANDROID_INSTRUCTIONS_START -->
                <!-- <h4 style="margin-top: 20px; margin-bottom: 10px;">For Android Devices:</h4>
                <ol style="margin-top: 0; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Ensure your device is powered on and connected to the internet.</li>
                    <li style="margin-bottom: 8px;">Open the <strong>Company Portal</strong> app on your device.</li>
                    <li style="margin-bottom: 8px;">Tap the menu icon (three lines) in the top left corner.</li>
                    <li style="margin-bottom: 8px;">Tap <strong>Devices</strong>, then select your device.</li>
                    <li style="margin-bottom: 8px;">Tap <strong>Check Status</strong> or <strong>Sync Device</strong>.</li>
                </ol> -->
                <!-- ANDROID_INSTRUCTIONS_END -->
            </div>
            
            <!-- UPDATES_SECTION_START -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 10px;">Check for Updates</h3>
                <p>While syncing, also check for any pending updates:</p>
                
                <!-- Windows Updates -->
                <!-- WINDOWS_UPDATES_START -->
                <h4 style="margin-top: 20px; margin-bottom: 10px;">For Windows Devices:</h4>
                <ol style="margin-top: 0; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Open <strong>Settings</strong> > <strong>Windows Update</strong>.</li>
                    <li style="margin-bottom: 8px;">Click <strong>Check for updates</strong> and install any available updates.</li>
                    <li style="margin-bottom: 8px;">Restart your device if prompted.</li>
                </ol>
                <!-- WINDOWS_UPDATES_END -->
                
                <!-- iOS Updates -->
                <!-- IOS_UPDATES_START -->
                <h4 style="margin-top: 20px; margin-bottom: 10px;">For iOS/iPadOS Devices:</h4>
                <ol style="margin-top: 0; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Go to <strong>Settings</strong> > <strong>General</strong> > <strong>Software Update</strong>.</li>
                    <li style="margin-bottom: 8px;">If updates are available, tap <strong>Download and Install</strong>.</li>
                </ol>
                <!-- IOS_UPDATES_END -->
            </div>
            <!-- UPDATES_SECTION_END -->
            
            <!-- Why This Matters Section -->
            <div style="margin: 25px 0; padding: 15px; background-color: #f8f8f8; border-left: 4px solid #007bff;">
                <p style="margin: 0;"><strong>Why this matters:</strong> Regular syncing ensures your device receives the latest security policies and configurations. Keeping your device updated helps protect your data and our organization's network.</p>
            </div>
            
            <!-- Support Info -->
            <div style="margin-top: 20px;">
                <p>If you encounter any issues or need assistance, please contact the IT Help Desk at <a href="mailto:$SupportEmail" style="color: #007bff; text-decoration: underline;">$SupportEmail</a> or call $SupportPhone.</p>
            </div>
        </div>
        
        <!-- Footer Section -->
        <div style="padding: 15px; background-color: #f2f2f2; border-top: 1px solid #ddd; text-align: center; font-size: 14px; color: #666;">
            <p style="margin: 0;"><strong>Do not reply to this message.</strong> This email was sent from an unmonitored mailbox.</p>
        </div>
    </div>
</body>
</html>
"@
        
        $body = $body.Replace('$LogoUrl', $LogoUrl)
        $body = $body.Replace('$Username', $Username)
        $body = $body.Replace('$DeviceName', $DeviceName)
        $body = $body.Replace('$LastSyncTime', $LastSyncTime)
        $body = $body.Replace('$SupportEmail', $SupportEmail)
        $body = $body.Replace('$SupportPhone', $SupportPhone)
        
        return Send-GraphMailMessage -Token $Token -From $From -To $To -Subject $subject -HtmlBody $body -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    }
    catch {
        Write-Log "Failed to prepare email to $To`: $_" -Type "ERROR"
        return $false
    }
}

function Send-ITNotificationEmail {
    param (
        [string]$Token,
        [string]$To,
        [array]$DevicesWithoutUsers,
        [string]$From,
        [string]$LogoUrl,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        $subject = "Intune Device Sync Report: Devices Without Primary Users"
        
        $deviceTableRows = $DevicesWithoutUsers | ForEach-Object {
            "<tr>
                <td style='padding: 8px; border: 1px solid #ddd;'>$($_.deviceName)</td>
                <td style='padding: 8px; border: 1px solid #ddd;'>$($_.operatingSystem)</td>
                <td style='padding: 8px; border: 1px solid #ddd;'>$($_.model)</td>
                <td style='padding: 8px; border: 1px solid #ddd;'>$($_.serialNumber)</td>
                <td style='padding: 8px; border: 1px solid #ddd;'>$([datetime]$_.lastSyncDateTime)</td>
            </tr>"
        }
        
        $body = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Intune Devices Without Primary Users</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; background-color: #f9f9f9; margin: 0; padding: 0; color: #333;">
  <table cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width: 800px; margin: 0 auto; background-color: #ffffff;">
    <tr>
      <td style="padding: 20px;">
        <!-- Logo -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td align="center" style="padding-bottom: 20px;">
              <img src="https://via.placeholder.com/150x50?text=Your+Logo" alt="Company Logo" width="150" style="display: block;" />
            </td>
          </tr>
        </table>
        
        <!-- Header -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td align="center" style="padding-bottom: 20px; border-bottom: 1px solid #eee;">
              <h2 style="color: #333; margin: 0;">Intune Devices Without Primary Users</h2>
            </td>
          </tr>
        </table>
        
        <!-- Summary Box -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 20px 0; background-color: #f8f8f8; border-left: 4px solid #4285f4;">
          <tr>
            <td style="padding: 15px;">
              <p style="margin: 0;"><span style="font-weight: bold; color: #e53935;">$DeviceCount</span> devices have not synced in the past $DaysSinceLastSync days and do not have a primary user assigned. These devices require IT attention.</p>
            </td>
          </tr>
        </table>
        
        <!-- Device Table -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 20px 0;">
          <tr>
            <td>
              <table cellpadding="0" cellspacing="0" border="1" width="100%" style="border-collapse: collapse; border-color: #ddd;">
                <tr>
                  <th style="background-color: #f2f2f2; padding: 10px 8px; text-align: left; font-weight: bold;">Device Name</th>
                  <th style="background-color: #f2f2f2; padding: 10px 8px; text-align: left; font-weight: bold;">OS</th>
                  <th style="background-color: #f2f2f2; padding: 10px 8px; text-align: left; font-weight: bold;">Model</th>
                  <th style="background-color: #f2f2f2; padding: 10px 8px; text-align: left; font-weight: bold;">Serial Number</th>
                  <th style="background-color: #f2f2f2; padding: 10px 8px; text-align: left; font-weight: bold;">Last Sync Time</th>
                </tr>
                $DeviceTableRows
              </table>
            </td>
          </tr>
        </table>
        
        <!-- Action Items -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding-bottom: 15px; color: #555;">
              <p>These devices may need:</p>
              <table cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="padding-left: 20px; padding-bottom: 5px; color: #555;">• Primary user assignment</td>
                </tr>
                <tr>
                  <td style="padding-left: 20px; padding-bottom: 5px; color: #555;">• Manual sync initiation</td>
                </tr>
                <tr>
                  <td style="padding-left: 20px; padding-bottom: 5px; color: #555;">• Investigation for connectivity issues</td>
                </tr>
                <tr>
                  <td style="padding-left: 20px; padding-bottom: 5px; color: #555;">• Possible retirement if no longer in use</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        
        <!-- Additional Info -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding-bottom: 15px; color: #555;">
              <p>Please review these devices and take appropriate action to ensure they remain compliant and secure. Devices without assigned users cannot receive the automated sync reminder emails sent to end users.</p>
              <p>For a complete view of all devices, please check the Intune portal.</p>
            </td>
          </tr>
        </table>
        
        <!-- Footer -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px;">
          <tr>
            <td align="center" style="font-size: 0.9em; color: #666;">
              <p style="margin-bottom: 5px;">This is an automated message from the Intune Device Management System.</p>
              <p style="margin: 0;">Report generated on: $CurrentDate</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@
        
        if ($WhatIf) {
            Write-Log "Would send IT notification email from $From to $To with subject: $Subject" -Type "WHATIF"
            Write-Log "Email would contain information about $($DevicesWithoutUsers.Count) devices without primary users" -Type "WHATIF"
            return $true
        }
        
        Write-Log "Sending IT notification email to $To with information about $($DevicesWithoutUsers.Count) devices..."
        
        $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        
        $messageBody = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "HTML"
                    content = $Body
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $To
                        }
                    }
                )
            }
            saveToSentItems = $true
        }
        
        Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $messageBody -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "IT notification email sent to $To successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send IT notification email to $To`: $_" -Type "ERROR"
        return $false
    }
}

function Process-DeviceBatch {
    param (
        [string]$Token,
        [array]$Devices,
        [datetime]$SyncThreshold,
        [string]$EmailSender,
        [string]$LogoUrl,
        [string]$TestEmailAddress,
        [switch]$WhatIf,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds,
        [int]$MaxEmails,
        [string[]]$ExcludedCategories,
        [string]$SupportEmail,
        [string]$SupportPhone,
        [System.Collections.ArrayList]$DevicesWithoutUsers
    )
    
    $batchEmailCount = 0
    $batchSkippedCount = 0
    $batchErrorCount = 0
    
    foreach ($device in $Devices) {
        try {
            $deviceName = $device.deviceName
            $deviceId = $device.id
            $osType = $device.operatingSystem
            $lastSyncDateTime = [datetime]$device.lastSyncDateTime
            $deviceCategory = $device.deviceCategoryDisplayName
            
            if ($ExcludedCategories -contains $deviceCategory) {
                Write-Log "Skipping device $deviceName due to excluded category: $deviceCategory"
                $batchSkippedCount++
                $Stats.SkippedCategoryCount++
                continue
            }
            
            Write-Log "Processing device: $deviceName (ID: $deviceId, OS: $osType)"
            Write-Log "Last sync time: $($lastSyncDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
            
            if ($lastSyncDateTime -lt $SyncThreshold) {
                Write-Log "Device $deviceName has not synced since threshold date ($($SyncThreshold.ToString('yyyy-MM-dd')))" -Type "WARNING"
                
                $primaryUser = Get-DevicePrimaryUser -Token $Token -DeviceId $deviceId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                
                if ($null -ne $primaryUser -and (-not [string]::IsNullOrEmpty($primaryUser.mail) -or -not [string]::IsNullOrEmpty($primaryUser.userPrincipalName))) {
                    $userEmail = if (-not [string]::IsNullOrEmpty($primaryUser.mail)) { $primaryUser.mail } else { $primaryUser.userPrincipalName }
                    $userDisplayName = $primaryUser.displayName
                    $userFirstName = $primaryUser.givenName
                    
                    if ([string]::IsNullOrEmpty($userFirstName)) {
                        $userFirstName = $userDisplayName.Split(' ')[0]
                    }
                    
                    Write-Log "Found primary user $userDisplayName with email $userEmail"
                    
                    if ($Stats.EmailsSent -ge $MaxEmails) {
                        Write-Log "Maximum number of emails reached ($MaxEmails). Skipping remaining devices." -Type "WARNING"
                        $batchSkippedCount++
                        $Stats.SkippedMaxEmailsCount++
                        continue
                    }
                    
                    $recipientEmail = if (-not [string]::IsNullOrEmpty($TestEmailAddress)) { $TestEmailAddress } else { $userEmail }
                    
                    $lastSyncFormatted = $lastSyncDateTime.ToString('MMMM d, yyyy h:mm tt')
                    
                    $emailResult = Send-SyncReminderEmail -Token $Token -To $recipientEmail -Username $userFirstName -DeviceName $deviceName -LastSyncTime $lastSyncFormatted -From $EmailSender -LogoUrl $LogoUrl -SupportEmail $SupportEmail -SupportPhone $SupportPhone -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                    
                    if ($emailResult) {
                        $batchEmailCount++
                        $Stats.EmailsSent++
                        
                        if ($osType) {
                            if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                                $Stats.OSTypeStats[$osType] = @{
                                    "Total" = 0
                                    "EmailsSent" = 0
                                }
                            }
                            $Stats.OSTypeStats[$osType]["EmailsSent"]++
                        }
                    }
                    else {
                        $batchErrorCount++
                        $Stats.ErrorCount++
                    }
                }
                else {
                    Write-Log "No primary user with email found for device $deviceName. Adding to IT notification list." -Type "WARNING"
                    $deviceInfo = [PSCustomObject]@{
                        deviceName = $deviceName
                        operatingSystem = $osType
                        model = $device.model
                        serialNumber = $device.serialNumber
                        lastSyncDateTime = $lastSyncDateTime
                    }
                    [void]$DevicesWithoutUsers.Add($device)
                    $batchSkippedCount++
                    $Stats.SkippedNoUserCount++
                }
            }
            else {
                Write-Log "Device $deviceName has synced recently. Last sync: $lastSyncDateTime. Threshold: $SyncThreshold"
                $batchSkippedCount++
                $Stats.RecentlySyncedCount++
            }
            
            if ($osType) {
                if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                    $Stats.OSTypeStats[$osType] = @{
                        "Total" = 0
                        "EmailsSent" = 0
                    }
                }
                $Stats.OSTypeStats[$osType]["Total"]++
            }
        }
        catch {
            Write-Log "Error processing device $($device.deviceName): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
        }
    }
    
    return @{
        EmailCount = $batchEmailCount
        SkippedCount = $batchSkippedCount
        ErrorCount = $batchErrorCount
    }
}

# Main script execution starts here
try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO EMAILS WILL BE SENT ===" -Type "WHATIF"
    }
    
    Write-Log "=== Intune Device Sync Reminder Process Started ==="
    Write-Log "Days since last sync threshold: $DaysSinceLastSync"
    Write-Log "Email sender: $EmailSender"
    if (-not [string]::IsNullOrEmpty($TestEmailAddress)) {
        Write-Log "TEST MODE: All emails will be sent to $TestEmailAddress" -Type "WARNING"
    }
    if (-not [string]::IsNullOrEmpty($ITDepartmentEmail)) {
        Write-Log "IT Department email for devices without primary users: $ITDepartmentEmail"
    }
    
    $startTime = Get-Date
    $syncThreshold = (Get-Date).AddDays(-$DaysSinceLastSync)
    
    Write-Log "Sync threshold date: $($syncThreshold.ToString('yyyy-MM-dd'))"
    
    $token = Get-MsGraphToken
    
    $devices = Get-IntuneDevices -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    $stats = @{
        TotalDevices = 0
        OutdatedDevices = 0
        EmailsSent = 0
        ErrorCount = 0
        RecentlySyncedCount = 0
        SkippedNoUserCount = 0
        SkippedCategoryCount = 0
        SkippedMaxEmailsCount = 0
        OSTypeStats = @{}
    }
    
    $devicesWithoutUsers = New-Object System.Collections.ArrayList
    $stats.TotalDevices = $devices.Count
    $outdatedDevices = $devices | Where-Object { [datetime]$_.lastSyncDateTime -lt $syncThreshold }
    $stats.OutdatedDevices = $outdatedDevices.Count
    
    Write-Log "Found $($stats.OutdatedDevices) devices that haven't synced since $($syncThreshold.ToString('yyyy-MM-dd'))"
    
    if ($ExcludedDeviceCategories.Count -gt 0) {
        Write-Log "Excluded device categories: $($ExcludedDeviceCategories -join ', ')"
    }
    
    $totalBatches = [Math]::Ceiling($devices.Count / $BatchSize)
    Write-Log "Processing $($devices.Count) devices in $totalBatches batches of maximum $BatchSize devices"
    
    for ($batchNum = 0; $batchNum -lt $totalBatches; $batchNum++) {
        $start = $batchNum * $BatchSize
        $end = [Math]::Min(($batchNum + 1) * $BatchSize - 1, $devices.Count - 1)
        $currentBatchSize = $end - $start + 1
        
        Write-Log "Processing batch $($batchNum+1) of $totalBatches (devices $($start+1) to $($end+1) of $($devices.Count))"
        $currentBatch = $devices[$start..$end]
        $batchResult = Process-DeviceBatch -Token $token -Devices $currentBatch -SyncThreshold $syncThreshold -EmailSender $EmailSender `
            -LogoUrl $LogoUrl -TestEmailAddress $TestEmailAddress -WhatIf:$WhatIf -Stats $stats -MaxRetries $MaxRetries `
            -InitialBackoffSeconds $InitialBackoffSeconds -MaxEmails $MaxEmailsPerRun -ExcludedCategories $ExcludedDeviceCategories `
            -SupportEmail $SupportEmail -SupportPhone $SupportPhone
        
        Write-Log "Batch $($batchNum+1) results: $($batchResult.EmailCount) emails sent, $($batchResult.SkippedCount) skipped, $($batchResult.ErrorCount) errors"
        
        if (-not [string]::IsNullOrEmpty($ITDepartmentEmail) -and $devicesWithoutUsers.Count -gt 0) {
            Write-Log "Sending notification to IT department about $($devicesWithoutUsers.Count) devices without primary users"
            
            $itEmailResult = Send-ITNotificationEmail -Token $token -To $ITDepartmentEmail -DevicesWithoutUsers $devicesWithoutUsers `
                -From $EmailSender -LogoUrl $LogoUrl -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($itEmailResult) {
                Write-Log "Successfully sent IT notification email to $ITDepartmentEmail"
                if (-not $stats.ContainsKey("ITNotificationSent")) {
                    $stats.Add("ITNotificationSent", $true)
                } else {
                    $stats["ITNotificationSent"] = $true
                }
                
                if (-not $stats.ContainsKey("DevicesWithoutUsersCount")) {
                    $stats.Add("DevicesWithoutUsersCount", $devicesWithoutUsers.Count)
                } else {
                    $stats["DevicesWithoutUsersCount"] = $devicesWithoutUsers.Count
                }
            } else {
                Write-Log "Failed to send IT notification email" -Type "ERROR"
                if (-not $stats.ContainsKey("ITNotificationSent")) {
                    $stats.Add("ITNotificationSent", $false)
                } else {
                    $stats["ITNotificationSent"] = $false
                }
            }
        } else {
            if ([string]::IsNullOrEmpty($ITDepartmentEmail)) {
                Write-Log "No IT department email specified. Skipping notification about devices without primary users."
            } else {
                Write-Log "No devices without primary users found that need syncing."
            }
            if (-not $stats.ContainsKey("ITNotificationSent")) {
                $stats.Add("ITNotificationSent", $false)
            } else {
                $stats["ITNotificationSent"] = $false
            }
        }

        if ($stats.EmailsSent -ge $MaxEmailsPerRun) {
            Write-Log "Maximum number of emails reached ($MaxEmailsPerRun). Stopping processing." -Type "WARNING"
            break
        }
        
        if ($batchNum -lt $totalBatches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Intune Device Sync Reminder Process Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO EMAILS WERE SENT ===" -Type "WHATIF"
    }
    
    Write-Log "Overall Summary:"
    Write-Log "Total devices processed: $($stats.TotalDevices)"
    Write-Log "Devices with outdated sync: $($stats.OutdatedDevices)"
    
    if ($WhatIf) {
        Write-Log "Would have sent emails: $($stats.EmailsSent)" -Type "WHATIF"
    } else {
        Write-Log "Emails sent: $($stats.EmailsSent)"
    }
    
    Write-Log "Devices with recent sync: $($stats.RecentlySyncedCount)"
    Write-Log "Devices skipped due to no user: $($stats.SkippedNoUserCount)"
    Write-Log "Devices skipped due to excluded category: $($stats.SkippedCategoryCount)"
    Write-Log "Devices skipped due to maximum emails limit: $($stats.SkippedMaxEmailsCount)"
    Write-Log "Errors: $($stats.ErrorCount)"
    
    $outputObject = [PSCustomObject][ordered]@{
        TotalDevices = $stats.TotalDevices
        OutdatedDevices = $stats.OutdatedDevices
        EmailsSent = $stats.EmailsSent
        RecentlySyncedCount = $stats.RecentlySyncedCount
        SkippedNoUserCount = $stats.SkippedNoUserCount
        SkippedCategoryCount = $stats.SkippedCategoryCount
        SkippedMaxEmailsCount = $stats.SkippedMaxEmailsCount
        ErrorCount = $stats.ErrorCount
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        SyncThresholdDate = $syncThreshold.ToString('yyyy-MM-dd')
    }
    
    foreach ($os in $stats.OSTypeStats.Keys | Sort-Object) {
        Write-Log "$os Device Summary:"
        Write-Log "- Total $os devices: $($stats.OSTypeStats[$os]["Total"])"
        
        if ($WhatIf) {
            Write-Log "- Would have sent emails: $($stats.OSTypeStats[$os]["EmailsSent"])" -Type "WHATIF"
        } else {
            Write-Log "- Emails sent: $($stats.OSTypeStats[$os]["EmailsSent"])"
        }
        
        $outputObject | Add-Member -MemberType NoteProperty -Name "${os}Devices" -Value $stats.OSTypeStats[$os]["Total"]
        $outputObject | Add-Member -MemberType NoteProperty -Name "${os}EmailsSent" -Value $stats.OSTypeStats[$os]["EmailsSent"]
    }
    
    return $outputObject
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
}