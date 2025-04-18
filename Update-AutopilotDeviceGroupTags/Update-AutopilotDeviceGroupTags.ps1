# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
    Updates Windows Autopilot device group tags to match their corresponding Intune device categories.
    
.DESCRIPTION
    This Azure Runbook script authenticates to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all Intune device categories and Windows Autopilot devices, and updates the group tag of 
    each Autopilot device to match its corresponding device category in Intune. The script supports
    batch processing, handles API throttling, and includes detailed logging.
    
.PARAMETER WhatIf
    If specified, shows what changes would occur without actually making any updates.
    
.PARAMETER BatchSize
    Number of devices to process in each batch. Default is 50.
    
.PARAMETER BatchDelaySeconds
    Number of seconds to wait between processing batches. Default is 10.
    
.PARAMETER MaxRetries
    Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.NOTES
    File Name: Update-AutopilotDeviceGroupTags.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-10
    Updated: 2025-04-18
    
    Companion script to Update-IntuneDeviceCategories.ps1
    
    Required permissions:
    - DeviceManagementManagedDevices.Read.All
    - DeviceManagementManagedDevices.ReadWrite.All
    - DeviceManagementServiceConfig.ReadWrite.All
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchDelaySeconds = 10,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5
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

function Get-MsGraphToken {
    param (
        [switch]$UseManagedIdentity = $true
    )
    
    try {
        Write-Log "Acquiring Microsoft Graph token using Managed Identity..."
        
        Write-Log "Connecting to Azure with managed identity..."
        Connect-AzAccount -Identity | Out-Null
        
        Write-Log "Successfully connected to Azure, retrieving access token..."
        $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
        
        if (-not [string]::IsNullOrEmpty($token)) {
            Write-Log "Successfully acquired token using Managed Identity"
            return $token
        }
        else {
            throw "Failed to acquire token - token is empty"
        }
    }
    catch {
        Write-Log "Failed to acquire Microsoft Graph token using Managed Identity: $_" -Type "ERROR"
        throw "Authentication failed: $_"
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

function Get-IntuneDeviceCategories {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving device categories..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories"
        $categories = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        Write-Log "Retrieved $($categories.value.Count) device categories"
        
        return $categories.value
    }
    catch {
        Write-Log "Failed to retrieve device categories: $_" -Type "ERROR"
        throw "Failed to retrieve device categories: $_"
    }
}

function Get-AutopilotDevices {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving Windows Autopilot devices..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $devices += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of Autopilot devices..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $devices += $response.value
        }
        
        Write-Log "Retrieved $($devices.Count) Autopilot devices"
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve Autopilot devices: $_" -Type "ERROR"
        throw "Failed to retrieve Autopilot devices: $_"
    }
}

function Get-IntuneDevicesWithCategories {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving Intune devices with their categories..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=id,serialNumber,deviceCategoryDisplayName,deviceName"
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $devices += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of Intune devices..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $devices += $response.value
        }
        
        Write-Log "Retrieved $($devices.Count) Intune devices with category information"
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve Intune devices: $_" -Type "ERROR"
        throw "Failed to retrieve Intune devices: $_"
    }
}

function Update-AutopilotDeviceGroupTag {
    param (
        [string]$Token,
        [string]$DeviceId,
        [string]$SerialNumber,
        [string]$CurrentGroupTag,
        [string]$NewGroupTag,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would update Autopilot device group tag for device with serial number $SerialNumber from '$CurrentGroupTag' to '$NewGroupTag'" -Type "WHATIF"
            return $true
        }
        else {
            Write-Log "Updating Autopilot device group tag for device with serial number $SerialNumber from '$CurrentGroupTag' to '$NewGroupTag'"
            
            $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$DeviceId/updateDeviceProperties"
            $body = @{
                groupTag = $NewGroupTag
            }
            Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "PATCH" -Body $body -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            Write-Log "Successfully updated Autopilot device group tag"
            return $true
        }
    }
    catch {
        Write-Log "Failed to update Autopilot device group tag for device with serial number $SerialNumber`: $_" -Type "ERROR"
        return $false
    }
}

function Process-DeviceBatch {
    param (
        [string]$Token,
        [array]$AutopilotDevices,
        [array]$IntuneDevicesMap,
        [switch]$WhatIf,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds
    )
    
    $batchUpdateCount = 0
    $batchNoChangeCount = 0
    $batchErrorCount = 0
    $batchNoMatchCount = 0
    $batchNoCategoryCount = 0
    
    foreach ($device in $AutopilotDevices) {
        try {
            $serialNumber = $device.serialNumber
            $currentGroupTag = $device.groupTag
            $deviceId = $device.id
            Write-Log "Processing Autopilot device with serial number: $serialNumber"
            Write-Log "Current group tag: '$currentGroupTag'"
            $intuneDevice = $IntuneDevicesMap | Where-Object { $_.serialNumber -eq $serialNumber }
            if ($null -ne $intuneDevice) {
                $deviceCategory = $intuneDevice.deviceCategoryDisplayName
                if (-not [string]::IsNullOrEmpty($deviceCategory)) {
                    if ($currentGroupTag -ne $deviceCategory) {
                        $updateResult = Update-AutopilotDeviceGroupTag -Token $Token -DeviceId $deviceId -SerialNumber $serialNumber -CurrentGroupTag $currentGroupTag -NewGroupTag $deviceCategory -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                        
                        if ($updateResult) {
                            Write-Log "Updated group tag from '$currentGroupTag' to '$deviceCategory'"
                            $batchUpdateCount++
                            $Stats.UpdatedCount++
                        }
                        else {
                            Write-Log "Failed to update group tag" -Type "ERROR"
                            $batchErrorCount++
                            $Stats.ErrorCount++
                        }
                    }
                    else {
                        Write-Log "Group tag already matches device category. No update needed."
                        $batchNoChangeCount++
                        $Stats.NoChangeCount++
                    }
                }
                else {
                    Write-Log "Intune device has no category assigned. Skipping." -Type "WARNING"
                    $batchNoCategoryCount++
                    $Stats.NoCategoryCount++
                }
            }
            else {
                Write-Log "No matching Intune device found with serial number: $serialNumber. Skipping." -Type "WARNING"
                $batchNoMatchCount++
                $Stats.NoMatchCount++
            }
        }
        catch {
            Write-Log "Error processing Autopilot device with serial number $($device.serialNumber): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
        }
    }
    
    return @{
        UpdatedCount = $batchUpdateCount
        NoChangeCount = $batchNoChangeCount
        ErrorCount = $batchErrorCount
        NoMatchCount = $batchNoMatchCount
        NoCategoryCount = $batchNoCategoryCount
    }
}

try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO CHANGES WILL BE MADE ===" -Type "WHATIF"
    }
    Write-Log "=== Windows Autopilot Device Group Tag Update Started ==="
    Write-Log "Batch Size: $BatchSize"
    Write-Log "Batch Delay: $BatchDelaySeconds seconds"
    $startTime = Get-Date
    $token = Get-MsGraphToken
    $deviceCategories = Get-IntuneDeviceCategories -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    Write-Log "Available device categories:"
    foreach ($category in $deviceCategories) {
        Write-Log "- $($category.displayName) (ID: $($category.id))"
    }
    $autopilotDevices = Get-AutopilotDevices -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $intuneDevices = Get-IntuneDevicesWithCategories -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $stats = @{
        UpdatedCount = 0
        NoChangeCount = 0
        ErrorCount = 0
        NoMatchCount = 0
        NoCategoryCount = 0
        TotalDevices = $autopilotDevices.Count
    }
    
    $totalDevices = $autopilotDevices.Count
    $batches = [Math]::Ceiling($totalDevices / $BatchSize)
    Write-Log "Processing $totalDevices Autopilot devices in $batches batches of maximum $BatchSize devices"
    for ($batchNum = 0; $batchNum -lt $batches; $batchNum++) {
        $start = $batchNum * $BatchSize
        $end = [Math]::Min(($batchNum + 1) * $BatchSize - 1, $totalDevices - 1)
        $currentBatchSize = $end - $start + 1
        Write-Log "Processing batch $($batchNum + 1) of $batches (devices $($start + 1) to $($end + 1) of $totalDevices)"
        $currentBatch = $autopilotDevices[$start..$end]
        $batchResult = Process-DeviceBatch -Token $token -AutopilotDevices $currentBatch -IntuneDevicesMap $intuneDevices -WhatIf:$WhatIf -Stats $stats -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        Write-Log "Batch $($batchNum + 1) results: $($batchResult.UpdatedCount) updated, $($batchResult.NoChangeCount) already correct, $($batchResult.NoCategoryCount) no category, $($batchResult.NoMatchCount) no match, $($batchResult.ErrorCount) errors"
        if ($batchNum -lt $batches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Log "=== Windows Autopilot Device Group Tag Update Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO CHANGES WERE MADE ===" -Type "WHATIF"
    }
    Write-Log "Overall Summary:"
    Write-Log "Total Autopilot devices: $($stats.TotalDevices)"
    if ($WhatIf) {
        Write-Log "Would be updated: $($stats.UpdatedCount)" -Type "WHATIF"
    }
    else {
        Write-Log "Updated: $($stats.UpdatedCount)"
    }
    Write-Log "Already correct: $($stats.NoChangeCount)"
    Write-Log "No category assigned: $($stats.NoCategoryCount)"
    Write-Log "No matching Intune device: $($stats.NoMatchCount)"
    Write-Log "Errors: $($stats.ErrorCount)"
    $outputProperties = [ordered]@{
        TotalDevices = $stats.TotalDevices
        UpdatedCount = $stats.UpdatedCount
        NoChangeCount = $stats.NoChangeCount
        NoCategoryCount = $stats.NoCategoryCount
        NoMatchCount = $stats.NoMatchCount
        ErrorCount = $stats.ErrorCount
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        BatchesProcessed = $batches
    }
    $outputObject = [PSCustomObject]$outputProperties
    return $outputObject
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
}