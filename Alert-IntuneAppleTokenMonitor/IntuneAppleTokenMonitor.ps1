<#
.SYNOPSIS
    Monitors Intune Apple token and certificate expiration dates and sends notifications.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves expiration information for various Intune tokens (Apple Push Notification service certificates,
    VPP tokens, DEP tokens), and sends Teams and/or email notifications for tokens approaching expiration.
    
.PARAMETER WarningThresholdDays
    The number of days before expiration to start sending warning notifications.
    Default is 30 days.
    
.PARAMETER TeamsWebhookUrl
    Optional. Microsoft Teams webhook URL for sending notifications about token status.
    If not specified, Teams notifications will not be sent.
    
.PARAMETER SendEmailNotification
    Switch parameter. When specified, email notifications will be sent.
    Uses true/false to indicate whether to send email notifications.
    If specified, EmailSender and EmailRecipients must also be provided.
    
.PARAMETER EmailSender
    The email address that will be used as the sender for email notifications.
    Required if SendEmailNotification is specified.
    
.PARAMETER EmailRecipients
    A comma-separated list of email addresses that will receive the notifications.
    Required if SendEmailNotification is specified.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually send notifications.
    Use true to enable WhatIf mode.
    
.NOTES
    Author: Ryan Schultz
    Version: 1.2
    Created: 2025-04-17
    Updated: 2025-04-29
    
    Required Graph API Permissions for Managed Identity:
    - DeviceManagementServiceConfig.Read.All
    - DeviceManagementConfiguration.Read.All
    - DeviceManagementApps.Read.All
    - Organization.Read.All
    - Mail.Send (required for email notifications)
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$WarningThresholdDays = 30,
    
    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl,
    
    [Parameter(Mandatory = $false)]
    [switch]$SendEmailNotification,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailSender,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

Write-Output "=== Intune Apple Token Monitor Started ==="
Write-Output "Warning threshold: $WarningThresholdDays days"
$startTime = Get-Date
$notificationThreshold = (Get-Date).AddDays($WarningThresholdDays)
$tokenCollection = @()
$expiringTokens = @()

# Validate parameters for email notifications
if ($SendEmailNotification) {
    if ([string]::IsNullOrEmpty($EmailSender)) {
        Write-Output "ERROR: EmailSender parameter is required when SendEmailNotification is specified"
        throw "EmailSender parameter is required when SendEmailNotification is specified"
    }
    
    if ([string]::IsNullOrEmpty($EmailRecipients)) {
        Write-Output "ERROR: EmailRecipients parameter is required when SendEmailNotification is specified"
        throw "EmailRecipients parameter is required when SendEmailNotification is specified"
    }
    
    Write-Output "Email notifications will be sent from $EmailSender to $EmailRecipients"
}

if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
    Write-Output "Teams notifications will be sent to the specified webhook URL"
}

if (-not $SendEmailNotification -and [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
    Write-Output "WARNING: No notification method specified. No notifications will be sent."
}

# Connect to Microsoft Graph using Managed Identity
try {
    Write-Output "Connecting to Microsoft Graph using Managed Identity..."
    Connect-AzAccount -Identity | Out-Null
    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
    
    if ([string]::IsNullOrEmpty($token)) {
        throw "Failed to acquire token - token is empty"
    }
    Write-Output "Successfully connected to Microsoft Graph"
}
catch {
    Write-Output "Failed to connect to Microsoft Graph: $_"
    throw "Authentication failed: $_"
}

# Get organization information
try {
    Write-Output "Retrieving organization information..."
    $uri = "https://graph.microsoft.com/v1.0/organization"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $orgResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $orgDomain = ($orgResponse.value | 
                 Select-Object -ExpandProperty verifiedDomains | 
                 Where-Object { $_.isInitial } | 
                 Select-Object -ExpandProperty name)
    Write-Output "Organization domain: $orgDomain"
}
catch {
    Write-Output "Failed to retrieve organization information: $_"
    # Edit this line to set a default organization domain
    $orgDomain = "yourorganization.com"
}

# Check Apple Push Notification Certificate
try {
    Write-Output "Checking Apple Push Notification Certificate..."
    $uri = "https://graph.microsoft.com/beta/deviceManagement/applePushNotificationCertificate"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $applePushCert = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($applePushCert) {
        Write-Output "Found Apple Push Notification Certificate for $($applePushCert.appleIdentifier)"
        Write-Output "Certificate expires: $($applePushCert.expirationDateTime)"
        $expirationDate = [datetime]$applePushCert.expirationDateTime
        $daysLeft = ($expirationDate - (Get-Date)).Days
        Write-Output "Days until expiration: $daysLeft"
        $status = if ($daysLeft -le 0) { "Expired" } 
                 elseif ($daysLeft -le 7) { "Critical" } 
                 elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                 else { "OK" }
        $tokenInfo = [PSCustomObject]@{
            TokenType = "Apple Push Notification Certificate"
            Name = $applePushCert.appleIdentifier
            ExpirationDate = $expirationDate
            DaysUntilExpiration = $daysLeft
            Status = $status
        }
        $tokenCollection += $tokenInfo
        if ($notificationThreshold -ge $expirationDate) {
            Write-Output "Certificate will expire soon! Adding to notification list."
            $expiringTokens += $tokenInfo
        }
    }
    else {
        Write-Output "No Apple Push Notification Certificate found."
    }
}
catch {
    Write-Output "Error checking Apple Push Notification Certificate: $_"
}

# Check Apple VPP Tokens
try {
    Write-Output "Checking Apple VPP Tokens..."
    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/vppTokens"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $vppResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $vppTokens = $vppResponse.value
    if ($vppTokens -and $vppTokens.Count -gt 0) {
        Write-Output "Found $($vppTokens.Count) VPP tokens"
        foreach ($vppToken in $vppTokens) {
            Write-Output "Processing VPP token: $($vppToken.organizationName) - $($vppToken.appleId)"
            Write-Output "Token expires: $($vppToken.expirationDateTime)"
            $expirationDate = [datetime]$vppToken.expirationDateTime
            $daysLeft = ($expirationDate - (Get-Date)).Days
            Write-Output "Days until expiration: $daysLeft"
            $status = if ($daysLeft -le 0) { "Expired" } 
                     elseif ($daysLeft -le 7) { "Critical" } 
                     elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                     else { "OK" }
            $tokenInfo = [PSCustomObject]@{
                TokenType = "Apple VPP Token"
                Name = "$($vppToken.organizationName): $($vppToken.appleId)"
                ExpirationDate = $expirationDate
                DaysUntilExpiration = $daysLeft
                Status = $status
            }
            $tokenCollection += $tokenInfo
            if ($notificationThreshold -ge $expirationDate) {
                Write-Output "VPP token will expire soon! Adding to notification list."
                $expiringTokens += $tokenInfo
            }
        }
    }
    else {
        Write-Output "No Apple VPP tokens found."
    }
}
catch {
    Write-Output "Error checking Apple VPP tokens: $_"
}

# Check Apple DEP Tokens
try {
    Write-Output "Checking Apple DEP Tokens..."
    
    $uri = "https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    
    $depResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $depTokens = $depResponse.value
    if ($depTokens -and $depTokens.Count -gt 0) {
        Write-Output "Found $($depTokens.Count) DEP tokens"
        foreach ($depToken in $depTokens) {
            Write-Output "Processing DEP token: $($depToken.tokenName) - $($depToken.appleIdentifier)"
            Write-Output "Token expires: $($depToken.tokenExpirationDateTime)"
            if (-not [string]::IsNullOrEmpty($depToken.tokenExpirationDateTime)) {
                $expirationDate = [datetime]$depToken.tokenExpirationDateTime
                $daysLeft = ($expirationDate - (Get-Date)).Days
                Write-Output "Days until expiration: $daysLeft"
                
                $status = if ($daysLeft -le 0) { "Expired" } 
                         elseif ($daysLeft -le 7) { "Critical" } 
                         elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                         else { "OK" }
                
                $tokenInfo = [PSCustomObject]@{
                    TokenType = "Apple DEP Token"
                    Name = "$($depToken.tokenName): $($depToken.appleIdentifier)"
                    ExpirationDate = $expirationDate
                    DaysUntilExpiration = $daysLeft
                    Status = $status
                }
                $tokenCollection += $tokenInfo
                if ($notificationThreshold -ge $expirationDate) {
                    Write-Output "DEP token will expire soon! Adding to notification list."
                    $expiringTokens += $tokenInfo
                }
            }
            else {
                Write-Output "DEP token $($depToken.tokenName) has no expiration date."
            }
        }
    }
    else {
        Write-Output "No Apple DEP tokens found."
    }
}
catch {
    Write-Output "Error checking Apple DEP tokens: $_"
}

# Send Email notification
if ($expiringTokens.Count -gt 0 -and $SendEmailNotification) {
    try {
        Write-Output "Sending Email notification for $($expiringTokens.Count) expiring tokens..."
        
        $tokenTableRows = ""
        foreach ($token in $expiringTokens) {
            $statusColor = switch ($token.Status) {
                "OK" { "#4CAF50" }
                "Warning" { "#FFC107" }
                "Critical" { "#F44336" }
                "Expired" { "#F44336" }
                default { "#000000" }
            }
            
            $tokenTableRows += @"
<tr>
    <td style="padding: 8px; border: 1px solid #ddd; font-family: Arial, sans-serif;">$($token.TokenType)</td>
    <td style="padding: 8px; border: 1px solid #ddd; font-family: Arial, sans-serif;">$($token.Name)</td>
    <td style="padding: 8px; border: 1px solid #ddd; font-family: Arial, sans-serif;">$($token.ExpirationDate.ToString("yyyy-MM-dd"))</td>
    <td style="padding: 8px; border: 1px solid #ddd; font-family: Arial, sans-serif;">$($token.DaysUntilExpiration)</td>
    <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; color: $statusColor; font-family: Arial, sans-serif;">$($token.Status)</td>
</tr>
"@
}
        
        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body style="font-family: Arial, sans-serif; margin: 0; padding: 0; background-color: #f9f9f9;">
    <div style="max-width: 800px; margin: 0 auto; border: 1px solid #ddd;">
        <div style="background-color: #0078D4; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0; padding: 0; font-size: 24px;">Intune Apple Token Expiration Alert</h2>
        </div>
        <div style="padding: 20px;">
            <p style="font-size: 16px; line-height: 1.5; color: #333;">The following Apple tokens or certificates in your Intune environment are approaching expiration or have expired:</p>
            
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr>
                    <th style="background-color: #f2f2f2; text-align: left; padding: 8px; border: 1px solid #ddd; font-weight: bold;">Token Type</th>
                    <th style="background-color: #f2f2f2; text-align: left; padding: 8px; border: 1px solid #ddd; font-weight: bold;">Name</th>
                    <th style="background-color: #f2f2f2; text-align: left; padding: 8px; border: 1px solid #ddd; font-weight: bold;">Expiration Date</th>
                    <th style="background-color: #f2f2f2; text-align: left; padding: 8px; border: 1px solid #ddd; font-weight: bold;">Days Remaining</th>
                    <th style="background-color: #f2f2f2; text-align: left; padding: 8px; border: 1px solid #ddd; font-weight: bold;">Status</th>
                </tr>
                $tokenTableRows
            </table>
            
            <p style="font-size: 16px; line-height: 1.5; color: #333;"><strong>Action Required:</strong> Please take immediate action to renew these tokens before they expire to avoid service disruptions.</p>
            
            <h3 style="color: #333; margin-top: 25px; font-size: 18px;">Renewal Process:</h3>
            <ul style="margin-left: 0; padding-left: 20px;">
                <li style="margin-bottom: 10px;"><strong>For APNs Certificate:</strong> Generate a new CSR in the Intune admin center, use it to renew the certificate in the Apple Push Certificates Portal, and upload the renewed certificate back to Intune.</li>
                <li style="margin-bottom: 10px;"><strong>For VPP Tokens:</strong> Download a new token from Apple Business Manager and upload it to Intune.</li>
                <li style="margin-bottom: 10px;"><strong>For DEP Tokens:</strong> Generate a new server token in Apple Business Manager and upload it to Intune.</li>
            </ul>
            
            <p style="font-size: 16px; line-height: 1.5; color: #333;">Failure to renew these tokens before expiration may result in disruption to device management capabilities.</p>
        </div>
        <div style="background-color: #f2f2f2; padding: 10px; text-align: center; font-size: 12px; color: #666;">
            <p style="margin: 5px 0;">This is an automated notification from the Intune Monitoring System for $orgDomain</p>
            <p style="margin: 5px 0;">Report generated: $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))</p>
        </div>
    </div>
</body>
</html>
"@
        
        if ($WhatIf) {
            Write-Output "WHATIF: Would send Email notification to $EmailRecipients with subject 'Intune Apple Token Expiration Alert'"
            Write-Output "Email would contain information about $($expiringTokens.Count) expiring tokens"
        }
        else {
            $recipientList = @()
            foreach ($recipient in $EmailRecipients.Split(',')) {
                $recipientList += @{
                    emailAddress = @{
                        address = $recipient.Trim()
                    }
                }
            }
            
            $mailMessage = @{
                message = @{
                    subject = "Intune Apple Token Expiration Alert"
                    body = @{
                        contentType = "HTML"
                        content = $emailBody
                    }
                    toRecipients = $recipientList
                }
                saveToSentItems = $true
            }
            
            $emailUri = "https://graph.microsoft.com/v1.0/users/$EmailSender/sendMail"
            Invoke-RestMethod -Uri $emailUri -Headers $headers -Method Post -Body ($mailMessage | ConvertTo-Json -Depth 10)
            
            Write-Output "Email notification sent successfully to $EmailRecipients"
        }
    }
    catch {
        Write-Output "Error sending Email notification: $_"
        
        if ($null -ne $_.Exception) {
            Write-Output "Exception type: $($_.Exception.GetType().FullName)"
            
            if ($null -ne $_.Exception.Response) {
                Write-Output "Status code: $($_.Exception.Response.StatusCode)"
                
                try {
                    $responseStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    Write-Output "Response body: $responseBody"
                }
                catch {
                    Write-Output "Could not read response body: $_"
                }
            }
            
            if ($_.Exception.InnerException) {
                Write-Output "Inner exception: $($_.Exception.InnerException.Message)"
            }
        }
    }
}
elseif ($SendEmailNotification -and $expiringTokens.Count -eq 0) {
    Write-Output "No expiring tokens found. No email notifications sent."
}

# Send Teams notification
if ($expiringTokens.Count -gt 0 -and -not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
    try {
        Write-Output "Sending Teams notification for $($expiringTokens.Count) expiring tokens..."
        if ($WhatIf) {
            Write-Output "WHATIF: Would send Teams notification about expiring tokens"
        }
        else {
            $attachments = @()
            foreach ($token in $expiringTokens) {
                $cardAttachment = @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    content = @{
                        type = "AdaptiveCard"
                        version = "1.0"
                        body = @(
                            @{
                                type = "TextBlock"
                                text = "Intune Apple Token Expiration Alert"
                                size = "Large"
                                weight = "Bolder"
                            },
                            @{
                                type = "TextBlock"
                                text = "Token requires immediate attention:"
                                wrap = $true
                            },
                            @{
                                type = "FactSet"
                                facts = @(
                                    @{
                                        title = "Token Type:"
                                        value = $token.TokenType
                                    },
                                    @{
                                        title = "Token Name:"
                                        value = $token.Name
                                    },
                                    @{
                                        title = "Expiration Date:"
                                        value = $token.ExpirationDate.ToString("yyyy-MM-dd")
                                    },
                                    @{
                                        title = "Days Remaining:"
                                        value = "$($token.DaysUntilExpiration)"
                                    },
                                    @{
                                        title = "Status:"
                                        value = $token.Status
                                    }
                                )
                            },
                            @{
                                type = "TextBlock"
                                text = "Please take action to renew this token before it expires."
                                wrap = $true
                            }
                        )
                    }
                }
                
                $attachments += $cardAttachment
            }
            
            $message = @{
                attachments = $attachments
            }
            $jsonBody = ConvertTo-Json -InputObject $message -Depth 10
            Write-Output "Sending JSON payload to Teams webhook:"
            
            $params = @{
                Uri = $TeamsWebhookUrl
                Method = "POST"
                Body = $jsonBody
                ContentType = "application/json"
                UseBasicParsing = $true
            }
            
            $teamsRequest = Invoke-WebRequest @params
            Write-Output "Teams notification sent successfully"
            Write-Output "Response: $($teamsRequest.StatusCode) $($teamsRequest.StatusDescription)"
        }
    }
    catch {
        Write-Output "Error sending Teams notification: $_"
        
        if ($null -ne $_.Exception) {
            Write-Output "Exception type: $($_.Exception.GetType().FullName)"
            
            if ($null -ne $_.Exception.Response) {
                Write-Output "Status code: $($_.Exception.Response.StatusCode)"
                
                try {
                    $responseStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    Write-Output "Response body: $responseBody"
                }
                catch {
                    Write-Output "Could not read response body: $_"
                }
            }
            
            if ($_.Exception.InnerException) {
                Write-Output "Inner exception: $($_.Exception.InnerException.Message)"
            }
        }
    }
}
elseif (-not [string]::IsNullOrEmpty($TeamsWebhookUrl) -and $expiringTokens.Count -eq 0) {
    Write-Output "No expiring tokens found. No Teams notifications sent."
}

# Generate summary
$healthyTokens = $tokenCollection | Where-Object { $_.Status -eq "OK" }
$warningTokens = $tokenCollection | Where-Object { $_.Status -eq "Warning" }
$criticalTokens = $tokenCollection | Where-Object { $_.Status -eq "Critical" -or $_.Status -eq "Expired" }

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Output "=== Intune Apple Token Monitor Summary ==="
Write-Output "Total tokens checked: $($tokenCollection.Count)"
Write-Output "Healthy tokens: $($healthyTokens.Count)"
Write-Output "Warning tokens: $($warningTokens.Count)"
Write-Output "Critical/expired tokens: $($criticalTokens.Count)"
Write-Output "Execution time: $($duration.TotalMinutes.ToString("0.00")) minutes"

$notificationMethod = @()
if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) { $notificationMethod += "Teams" }
if ($SendEmailNotification) { $notificationMethod += "Email" }
if ($notificationMethod.Count -eq 0) { $notificationMethod += "None" }

$result = [PSCustomObject]@{
    TotalTokensChecked = $tokenCollection.Count
    HealthyTokens = $healthyTokens.Count
    WarningTokens = $warningTokens.Count
    CriticalTokens = $criticalTokens.Count
    ExpiringTokens = $expiringTokens.Count
    ExecutionTimeMinutes = $duration.TotalMinutes
    TokenCollection = $tokenCollection
    ExpiringTokenDetails = $expiringTokens
    NotificationMethod = $notificationMethod -join ", "
    EmailNotificationSent = ($SendEmailNotification -and $expiringTokens.Count -gt 0 -and -not $WhatIf)
    TeamsNotificationSent = (-not [string]::IsNullOrEmpty($TeamsWebhookUrl) -and $expiringTokens.Count -gt 0 -and -not $WhatIf)
}

return $result