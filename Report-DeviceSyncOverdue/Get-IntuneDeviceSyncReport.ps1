# Requires -Modules "Az.Accounts", "Az.Storage"
<#
.SYNOPSIS
    Identifies Intune devices that haven't synced in a specified period and creates a report stored in Azure Blob Storage.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all managed devices from Intune that haven't synced within the specified threshold period, 
    and generates a detailed report stored in an Azure Storage Blob container.
    
.PARAMETER DaysSinceLastSync
    The number of days to use as a threshold for determining "stale" devices that need to sync.
    Default is 7 days.
    
.PARAMETER StorageAccountName
    The name of the Azure Storage Account where the report will be stored.
    
.PARAMETER StorageContainerName
    The name of the Blob container in the Storage Account where the report will be stored.
    
.PARAMETER ExcludedDeviceCategories
    Optional. An array of device categories to exclude from the report.
    
.PARAMETER BatchSize
    Optional. Number of devices to process in each batch. Default is 50.
    
.PARAMETER BatchDelaySeconds
    Optional. Number of seconds to wait between processing batches. Default is 10.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually create or upload the report.
    
.PARAMETER ReportFormat
    Optional. Format of the generated report. Valid values are "CSV", "JSON", "HTML". Default is "CSV".
    
.PARAMETER IncludeDetailedDeviceInfo
    Optional. If specified, includes more detailed information about each device in the report.

.NOTES
    File Name: Get-IntuneDeviceSyncReport.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-30
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$DaysSinceLastSync = 7,
    
    [Parameter(Mandatory = $true)]
    [string]$StorageAccountName,
    
    [Parameter(Mandatory = $true)]
    [string]$StorageContainerName,
    
    [Parameter(Mandatory = $false)]
    [string[]]$ExcludedDeviceCategories = @(),
    
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
    [ValidateSet("CSV", "JSON", "HTML")]
    [string]$ReportFormat = "CSV",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedDeviceInfo
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

function Get-IntuneOutdatedDevices {
    param (
        [string]$Token,
        [datetime]$SyncThreshold,
        [bool]$IncludeDetailedInfo = $false,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving Intune devices that haven't synced since $($SyncThreshold.ToString('yyyy-MM-dd'))..."
        
        $thresholdDateString = $SyncThreshold.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filter = "(operatingSystem eq 'Windows' or operatingSystem eq 'iOS' or operatingSystem eq 'Android' or operatingSystem eq 'macOS') and lastSyncDateTime lt $thresholdDateString"
        $select = "id,deviceName,managedDeviceOwnerType,deviceType,operatingSystem,osVersion,complianceState,lastSyncDateTime,emailAddress,userPrincipalName,serialNumber,model,manufacturer,enrolledDateTime,userDisplayName,deviceCategoryDisplayName"
        
        if ($IncludeDetailedInfo) {
            $select += ",phoneNumber,wiFiMacAddress,imei,meid,complianceGracePeriodExpirationDateTime,managementAgent,managementState,isEncrypted,isSupervised,jailBroken,azureADRegistered,deviceEnrollmentType,deviceRegistrationState"
        }
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$filter&`$select=$select"
        
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $devices += $response.value
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of outdated devices..."
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
        
        Write-Log "Retrieved $($devices.Count) devices from Intune that haven't synced since threshold date ($osCountsString)"
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
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/users?`$select=id,displayName,mail,userPrincipalName,givenName,department,jobTitle,officeLocation"
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

function Process-DeviceBatch {
    param (
        [string]$Token,
        [array]$Devices,
        [bool]$IncludeDetailedInfo = $false,
        [switch]$WhatIf,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds,
        [string[]]$ExcludedCategories,
        [System.Collections.ArrayList]$ProcessedDevices
    )
    
    $batchProcessedCount = 0
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
            
            Write-Log "Processing outdated device: $deviceName (ID: $deviceId, OS: $osType)"
            Write-Log "Last sync time: $($lastSyncDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
            
            $deviceEntry = [PSCustomObject]@{
                DeviceName = $deviceName
                DeviceId = $deviceId
                OSType = $osType
                OSVersion = $device.osVersion
                LastSyncDateTime = $lastSyncDateTime
                DeviceCategory = $deviceCategory
                SerialNumber = $device.serialNumber
                Model = $device.model
                Manufacturer = $device.manufacturer
                EnrolledDateTime = $device.enrolledDateTime
                ComplianceState = $device.complianceState
                OwnerType = $device.managedDeviceOwnerType
                DaysSinceLastSync = ([datetime]::Now - $lastSyncDateTime).Days
                UserDisplayName = $device.userDisplayName
                UserEmail = $device.emailAddress
                UserPrincipalName = $device.userPrincipalName
                PrimaryUser = $null
                Department = $null
                JobTitle = $null
                OfficeLocation = $null
            }
            
            if ($IncludeDetailedInfo) {
                $detailedProperties = @{
                    PhoneNumber = $device.phoneNumber
                    WiFiMacAddress = $device.wiFiMacAddress
                    IMEI = $device.imei
                    MEID = $device.meid
                    ComplianceGraceExpiration = $device.complianceGracePeriodExpirationDateTime
                    ManagementAgent = $device.managementAgent
                    ManagementState = $device.managementState
                    IsEncrypted = $device.isEncrypted
                    IsSupervised = $device.isSupervised
                    JailBroken = $device.jailBroken
                    AzureADRegistered = $device.azureADRegistered
                    EnrollmentType = $device.deviceEnrollmentType
                    RegistrationState = $device.deviceRegistrationState
                }
                
                foreach ($prop in $detailedProperties.Keys) {
                    $deviceEntry | Add-Member -MemberType NoteProperty -Name $prop -Value $detailedProperties[$prop]
                }
            }
            
            $primaryUser = Get-DevicePrimaryUser -Token $Token -DeviceId $deviceId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($null -ne $primaryUser) {
                Write-Log "Found primary user $($primaryUser.displayName) for device $deviceName"
                $deviceEntry.PrimaryUser = $primaryUser.displayName
                $deviceEntry.Department = $primaryUser.department
                $deviceEntry.JobTitle = $primaryUser.jobTitle
                $deviceEntry.OfficeLocation = $primaryUser.officeLocation
                if ([string]::IsNullOrEmpty($deviceEntry.UserEmail) -and -not [string]::IsNullOrEmpty($primaryUser.mail)) {
                    $deviceEntry.UserEmail = $primaryUser.mail
                }
                
                if ([string]::IsNullOrEmpty($deviceEntry.UserPrincipalName) -and -not [string]::IsNullOrEmpty($primaryUser.userPrincipalName)) {
                    $deviceEntry.UserPrincipalName = $primaryUser.userPrincipalName
                }
                
                if ([string]::IsNullOrEmpty($deviceEntry.UserDisplayName) -and -not [string]::IsNullOrEmpty($primaryUser.displayName)) {
                    $deviceEntry.UserDisplayName = $primaryUser.displayName
                }
                
                if ($osType) {
                    if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                        $Stats.OSTypeStats[$osType] = @{
                            "Total" = 0
                            "HasPrimaryUser" = 0
                            "NoPrimaryUser" = 0
                        }
                    }
                    $Stats.OSTypeStats[$osType]["HasPrimaryUser"]++
                }
            }
            else {
                Write-Log "No primary user found for device $deviceName" -Type "WARNING"
                $Stats.NoUserCount++
                
                if ($osType) {
                    if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                        $Stats.OSTypeStats[$osType] = @{
                            "Total" = 0
                            "HasPrimaryUser" = 0
                            "NoPrimaryUser" = 0
                        }
                    }
                    $Stats.OSTypeStats[$osType]["NoPrimaryUser"]++
                }
            }
            
            if ($osType) {
                if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                    $Stats.OSTypeStats[$osType] = @{
                        "Total" = 0
                        "HasPrimaryUser" = 0
                        "NoPrimaryUser" = 0
                    }
                }
                $Stats.OSTypeStats[$osType]["Total"]++
            }
            
            [void]$ProcessedDevices.Add($deviceEntry)
            $batchProcessedCount++
        }
        catch {
            Write-Log "Error processing device $($device.deviceName): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
        }
    }
    
    return @{
        ProcessedCount = $batchProcessedCount
        SkippedCount = $batchSkippedCount
        ErrorCount = $batchErrorCount
    }
}

function Create-DeviceReport {
    param (
        [System.Collections.ArrayList]$DeviceData,
        [string]$OutputPath,
        [string]$ReportFormat = "CSV",
        [hashtable]$Stats,
        [datetime]$ReportTimestamp,
        [datetime]$SyncThreshold,
        [switch]$WhatIf
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would create $ReportFormat report at $OutputPath" -Type "WHATIF"
            return $true
        }
        
        Write-Log "Creating $ReportFormat report..."
        switch ($ReportFormat) {
            "CSV" {
                $DeviceData | Export-Csv -Path $OutputPath -NoTypeInformation
                Write-Log "CSV report created at $OutputPath"
                return $true
            }
            "JSON" {
                $jsonReport = @{
                    ReportMetadata = @{
                        GeneratedAt = $ReportTimestamp.ToString("yyyy-MM-dd HH:mm:ss")
                        SyncThreshold = $SyncThreshold.ToString("yyyy-MM-dd HH:mm:ss")
                        DaysSinceLastSync = $DaysSinceLastSync
                        TotalDevices = $Stats.TotalDevices
                        DevicesWithPrimaryUser = ($Stats.TotalDevices - $Stats.NoUserCount)
                        DevicesWithoutUser = $Stats.NoUserCount
                        OSStats = $Stats.OSTypeStats
                    }
                    Devices = $DeviceData
                }
                
                $jsonReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath
                Write-Log "JSON report created at $OutputPath"
                return $true
            }
            "HTML" {
                $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Intune Device Sync Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2, h3 { color: #0078D4; }
        .summary { background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #0078D4; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .no-user { background-color: #ffffcc; }
        .summary-section { margin-bottom: 15px; }
        .os-stats { display: flex; flex-wrap: wrap; }
        .os-stat-box { background-color: #f5f5f5; padding: 10px; margin: 10px; border-radius: 5px; min-width: 200px; }
        .days-critical { color: #d13438; font-weight: bold; }
        .days-warning { color: #ff8c00; font-weight: bold; }
        .days-normal { color: #107c10; }
    </style>
</head>
<body>
    <h1>Intune Device Sync Report</h1>
    
    <div class="summary">
        <div class="summary-section">
            <h2>Report Summary</h2>
            <p><strong>Generated:</strong> $($ReportTimestamp.ToString("yyyy-MM-dd HH:mm:ss"))</p>
            <p><strong>Sync Threshold:</strong> $($SyncThreshold.ToString("yyyy-MM-dd"))</p>
            <p><strong>Days Since Last Sync:</strong> $DaysSinceLastSync or more</p>
        </div>
        
        <div class="summary-section">
            <h3>Device Counts</h3>
            <p><strong>Total Outdated Devices:</strong> $($DeviceData.Count)</p>
            <p><strong>Devices With Primary User:</strong> $($DeviceData.Count - $Stats.NoUserCount)</p>
            <p><strong>Devices Without Primary User:</strong> $($Stats.NoUserCount)</p>
        </div>
        
        <div class="summary-section">
            <h3>OS Statistics</h3>
            <div class="os-stats">
"@

                foreach ($os in $Stats.OSTypeStats.Keys) {
                    $osStats = $Stats.OSTypeStats[$os]
                    $htmlHeader += @"
                <div class="os-stat-box">
                    <h4>$os</h4>
                    <p>Total: $($osStats.Total)</p>
                    <p>With User: $($osStats.HasPrimaryUser)</p>
                    <p>Without User: $($osStats.NoPrimaryUser)</p>
                </div>
"@
                }

                $htmlHeader += @"
            </div>
        </div>
    </div>
    
    <h2>Device Details</h2>
    <table>
        <tr>
            <th>Device Name</th>
            <th>OS</th>
            <th>Last Sync</th>
            <th>Days Overdue</th>
            <th>Primary User</th>
            <th>Department</th>
            <th>Model</th>
            <th>Serial Number</th>
        </tr>
"@

                $htmlRows = ""
                foreach ($device in $DeviceData) {
                    $rowClass = ""
                    if ([string]::IsNullOrEmpty($device.PrimaryUser)) {
                        $rowClass = "no-user"
                    }
                    
                    $daysSinceSync = $device.DaysSinceLastSync
                    $daysClass = "days-normal"
                    if ($daysSinceSync -ge 14) {
                        $daysClass = "days-critical"
                    } elseif ($daysSinceSync -ge 10) {
                        $daysClass = "days-warning"
                    }
                    
                    $htmlRows += @"
        <tr class="$rowClass">
            <td>$($device.DeviceName)</td>
            <td>$($device.OSType) $($device.OSVersion)</td>
            <td>$($device.LastSyncDateTime.ToString("yyyy-MM-dd HH:mm"))</td>
            <td class="$daysClass">$($device.DaysSinceLastSync)</td>
            <td>$($device.PrimaryUser)</td>
            <td>$($device.Department)</td>
            <td>$($device.Manufacturer) $($device.Model)</td>
            <td>$($device.SerialNumber)</td>
        </tr>
"@
                }

                $htmlFooter = @"
    </table>
    
    <div>
        <p><small>Report generated on $($ReportTimestamp.ToString("yyyy-MM-dd HH:mm:ss"))</small></p>
    </div>
</body>
</html>
"@

                $htmlReport = $htmlHeader + $htmlRows + $htmlFooter
                $htmlReport | Out-File -FilePath $OutputPath
                Write-Log "HTML report created at $OutputPath"
                return $true
            }
            default {
                Write-Log "Unsupported report format: $ReportFormat" -Type "ERROR"
                return $false
            }
        }
    }
    catch {
        Write-Log "Failed to create report: $_" -Type "ERROR"
        return $false
    }
}

function Upload-ToAzureBlob {
    param (
        [string]$StorageAccountName,
        [string]$ContainerName,
        [string]$FilePath,
        [string]$BlobName,
        [switch]$WhatIf
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would upload file $FilePath to blob container $ContainerName as $BlobName" -Type "WHATIF"
            return $true
        }
        
        Write-Log "Uploading file to Azure Storage..."
        
        $storageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
        $container = Get-AzStorageContainer -Name $ContainerName -Context $storageContext -ErrorAction SilentlyContinue
        if ($null -eq $container) {
            Write-Log "Container $ContainerName does not exist. Creating new container..." -Type "WARNING"
            New-AzStorageContainer -Name $ContainerName -Context $storageContext -Permission Off
        }
        
        $blobProperties = @{
            File      = $FilePath
            Container = $ContainerName
            Blob      = $BlobName
            Context   = $storageContext
            Properties = @{
                ContentType = switch ($ReportFormat) {
                    "CSV"  { "text/csv" }
                    "JSON" { "application/json" }
                    "HTML" { "text/html" }
                    default { "application/octet-stream" }
                }
            }
        }
        
        $blob = Set-AzStorageBlobContent @blobProperties -Force
        if ($blob) {
            Write-Log "File uploaded successfully to $($blob.Name)"
            
            $blobUrl = $blob.ICloudBlob.Uri.AbsoluteUri
            Write-Log "Blob URL: $blobUrl"
            
            return $blobUrl
        }
        else {
            Write-Log "Failed to upload file to Azure Storage" -Type "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error uploading file to Azure Storage: $_" -Type "ERROR"
        return $false
    }
}

# Main script starts here
try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO ACTUAL REPORTS WILL BE CREATED OR UPLOADED ===" -Type "WHATIF"
    }
    
    Write-Log "=== Intune Device Sync Report Process Started ==="
    Write-Log "Days since last sync threshold: $DaysSinceLastSync"
    Write-Log "Storage Account: $StorageAccountName"
    Write-Log "Container Name: $StorageContainerName"
    Write-Log "Report Format: $ReportFormat"
    Write-Log "Include Detailed Device Info: $IncludeDetailedDeviceInfo"
    
    $startTime = Get-Date
    $reportTimestamp = $startTime
    $syncThreshold = (Get-Date).AddDays(-$DaysSinceLastSync)
    Write-Log "Sync threshold date: $($syncThreshold.ToString('yyyy-MM-dd'))"
    $token = Get-MsGraphToken
    $devices = Get-IntuneOutdatedDevices -Token $token -SyncThreshold $syncThreshold -IncludeDetailedInfo $IncludeDetailedDeviceInfo -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $stats = @{
        TotalDevices = $devices.Count
        NoUserCount = 0
        SkippedCategoryCount = 0
        ErrorCount = 0
        OSTypeStats = @{}
    }
    
    $processedDevices = New-Object System.Collections.ArrayList
    if ($ExcludedDeviceCategories.Count -gt 0) {
        Write-Log "Excluded device categories: $($ExcludedDeviceCategories -join ', ')"
    }
    
    $totalBatches = [Math]::Ceiling($devices.Count / $BatchSize)
    Write-Log "Processing $($devices.Count) outdated devices in $totalBatches batches of maximum $BatchSize devices"
    for ($batchNum = 0; $batchNum -lt $totalBatches; $batchNum++) {
        $start = $batchNum * $BatchSize
        $end = [Math]::Min(($batchNum + 1) * $BatchSize - 1, $devices.Count - 1)
        $currentBatchSize = $end - $start + 1
        
        Write-Log "Processing batch $($batchNum+1) of $totalBatches (devices $($start+1) to $($end+1) of $($devices.Count))"
        $currentBatch = $devices[$start..$end]
        $batchResult = Process-DeviceBatch -Token $token -Devices $currentBatch `
            -IncludeDetailedInfo $IncludeDetailedDeviceInfo -WhatIf:$WhatIf -Stats $stats -MaxRetries $MaxRetries `
            -InitialBackoffSeconds $InitialBackoffSeconds -ExcludedCategories $ExcludedDeviceCategories -ProcessedDevices $processedDevices
        
        Write-Log "Batch $($batchNum+1) results: $($batchResult.ProcessedCount) processed, $($batchResult.SkippedCount) skipped, $($batchResult.ErrorCount) errors"
        if ($batchNum -lt $totalBatches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    $fileName = "IntuneOutdatedDevices_$(Get-Date -Format 'yyyyMMdd_HHmmss').$($ReportFormat.ToLower())"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $fileName)
    $reportCreated = Create-DeviceReport -DeviceData $processedDevices -OutputPath $tempPath -ReportFormat $ReportFormat `
                    -Stats $stats -ReportTimestamp $reportTimestamp -SyncThreshold $syncThreshold -WhatIf:$WhatIf
    
    if ($reportCreated) {
        $blobUrl = Upload-ToAzureBlob -StorageAccountName $StorageAccountName -ContainerName $StorageContainerName `
                -FilePath $tempPath -BlobName $fileName -WhatIf:$WhatIf
        
        if (Test-Path $tempPath) {
            Remove-Item -Path $tempPath -Force
            Write-Log "Temporary file removed"
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Log "=== Intune Device Sync Report Process Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO REPORT WAS CREATED OR UPLOADED ===" -Type "WHATIF"
    }
    
    Write-Log "Overall Summary:"
    Write-Log "Total outdated devices: $($stats.TotalDevices)"
    Write-Log "Devices without primary user: $($stats.NoUserCount)"
    Write-Log "Devices skipped due to excluded category: $($stats.SkippedCategoryCount)"
    Write-Log "Errors: $($stats.ErrorCount)"
    
    foreach ($os in $stats.OSTypeStats.Keys | Sort-Object) {
        Write-Log "$os Device Summary:"
        Write-Log "- Total $os devices: $($stats.OSTypeStats[$os]["Total"])"
        Write-Log "- With primary user: $($stats.OSTypeStats[$os]["HasPrimaryUser"])"
        Write-Log "- Without primary user: $($stats.OSTypeStats[$os]["NoPrimaryUser"])"
    }
    
    $outputObject = [PSCustomObject][ordered]@{
        OutdatedDevices = $stats.TotalDevices
        NoUserCount = $stats.NoUserCount
        SkippedCategoryCount = $stats.SkippedCategoryCount
        ErrorCount = $stats.ErrorCount
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        SyncThresholdDate = $syncThreshold.ToString('yyyy-MM-dd')
        ReportUrl = $blobUrl
        ReportFormat = $ReportFormat
    }
    
    foreach ($os in $stats.OSTypeStats.Keys | Sort-Object) {
        $outputObject | Add-Member -MemberType NoteProperty -Name "${os}Devices" -Value $stats.OSTypeStats[$os]["Total"]
        $outputObject | Add-Member -MemberType NoteProperty -Name "${os}WithUser" -Value $stats.OSTypeStats[$os]["HasPrimaryUser"]
        $outputObject | Add-Member -MemberType NoteProperty -Name "${os}NoUser" -Value $stats.OSTypeStats[$os]["NoPrimaryUser"]
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