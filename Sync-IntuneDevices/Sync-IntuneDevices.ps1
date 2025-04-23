<#
.SYNOPSIS
    Synchronizes all Microsoft Intune managed devices using Microsoft Graph API.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all managed devices from Intune, and initiates a sync command for each device.
    It supports throttling detection with retry logic and comprehensive logging.
    
.NOTES
    File Name: Sync-IntuneDevices.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-23
    
    Required permissions:
    - DeviceManagementManagedDevices.ReadWrite.All
    - DeviceManagementManagedDevices.PrivilegedOperations.All
#>

# Requires -Modules "Az.Accounts"

[CmdletBinding()]
param (
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
        default { 
            Write-Verbose $LogMessage -Verbose
        }
    }
}

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
        Write-Log "Failed to acquire Microsoft Graph token using Managed Identity: $_" -Type "ERROR"
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
        if ($Body -is [string]) {
            $params.Add("Body", $Body)
        } else {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
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
        Write-Log "Retrieving Intune managed devices..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
        
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $devices += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of devices..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $devices += $response.value
        }
        
        $deviceCount = $devices.Count
        Write-Log "Retrieved $deviceCount devices from Intune"
        
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve Intune devices: $_" -Type "ERROR"
        throw "Failed to retrieve Intune devices: $_"
    }
}

function Sync-IntuneDevice {
    param (
        [string]$Token,
        [string]$DeviceId,
        [string]$DeviceName,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Sending sync command to device: $DeviceName (ID: $DeviceId)"
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceId')/syncDevice"
        
        Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $null -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Successfully sent sync command to device: $DeviceName"
        return $true
    }
    catch {
        Write-Log "Failed to sync device $DeviceName`: $_" -Type "ERROR"
        return $false
    }
}

function Process-DeviceBatch {
    param (
        [string]$Token,
        [array]$Devices,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds
    )
    
    $batchSuccessCount = 0
    $batchErrorCount = 0
    
    foreach ($device in $Devices) {
        try {
            $deviceName = $device.deviceName
            $deviceId = $device.id
            
            $syncResult = Sync-IntuneDevice -Token $Token -DeviceId $deviceId -DeviceName $deviceName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($syncResult) {
                $batchSuccessCount++
                $Stats.SuccessCount++
            }
            else {
                $batchErrorCount++
                $Stats.ErrorCount++
            }
        }
        catch {
            Write-Log "Error processing device $($device.deviceName): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
        }
    }
    
    return @{
        SuccessCount = $batchSuccessCount
        ErrorCount = $batchErrorCount
    }
}

# Main script execution
try {
    Write-Log "=== Intune Device Sync Process Started ==="
    Write-Log "Batch Size: $BatchSize"
    Write-Log "Batch Delay: $BatchDelaySeconds seconds"
    
    $startTime = Get-Date
    
    $token = Get-MsGraphToken
    
    $devices = Get-IntuneDevices -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    $stats = @{
        TotalDevices = $devices.Count
        SuccessCount = 0
        ErrorCount = 0
    }
    
    $totalDevices = $devices.Count
    $batches = [Math]::Ceiling($totalDevices / $BatchSize)
    Write-Log "Processing $totalDevices devices in $batches batches of maximum $BatchSize devices"
    
    for ($batchNum = 0; $batchNum -lt $batches; $batchNum++) {
        $start = $batchNum * $BatchSize
        $end = [Math]::Min(($batchNum + 1) * $BatchSize - 1, $totalDevices - 1)
        $currentBatchSize = $end - $start + 1
        
        Write-Log "Processing batch $($batchNum+1) of $batches (devices $($start+1) to $($end+1) of $totalDevices)"
        
        $currentBatch = $devices[$start..$end]
        
        $batchResult = Process-DeviceBatch -Token $token -Devices $currentBatch -Stats $stats -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Batch $($batchNum+1) results: $($batchResult.SuccessCount) successful, $($batchResult.ErrorCount) errors"
        
        if ($batchNum -lt $batches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Intune Device Sync Process Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Total devices processed: $totalDevices"
    Write-Log "Successful syncs: $($stats.SuccessCount)"
    Write-Log "Errors: $($stats.ErrorCount)"
    
    $outputObject = [PSCustomObject]@{
        TotalDevices = $stats.TotalDevices
        SuccessCount = $stats.SuccessCount
        ErrorCount = $stats.ErrorCount
        DurationMinutes = $duration.TotalMinutes
        BatchesProcessed = $batches
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