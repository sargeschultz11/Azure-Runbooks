<#
.SYNOPSIS
    Updates Intune device categories (Windows, iOS, Android, and Linux) to match primary user's department.
.DESCRIPTION
    This Azure Runbook script authenticates to Microsoft Graph API using a client ID and secret,
    retrieves all Intune devices (Windows, iOS, Android, and Linux), and for devices with no category, sets the category
    to match the primary user's department. Includes a -WhatIf parameter for testing without making changes and
    an -OSType parameter to specify which types of devices to process.
    
    The script supports processing devices in batches with configurable batch size and delay to avoid API throttling.
    It also includes throttling detection and exponential backoff retry logic for handling Graph API rate limits.
.PARAMETER WhatIf
    If specified, shows what changes would occur without actually making any updates.
.PARAMETER OSType
    Specifies which operating systems to process. Valid values are "All", "Windows", "iOS", "Android", "Linux". Default is "All".
.PARAMETER BatchSize
    Number of devices to process in each batch. Default is 50.
.PARAMETER BatchDelaySeconds
    Number of seconds to wait between processing batches. Default is 10.
.PARAMETER MaxRetries
    Maximum number of retry attempts for throttled API requests. Default is 5.
.PARAMETER InitialBackoffSeconds
    Initial backoff period in seconds before retrying a throttled request. Default is 5.
.NOTES
    File Name: Update-IntuneDeviceCategories.ps1
    Author: Ryan Schultz
    Version: 2.0
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,

    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "Windows", "iOS", "Android", "Linux")]
    [string]$OSType = "All",
    
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
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    try {
        Write-Log "Attempting to acquire Microsoft Graph API token..."
        
        if ([string]::IsNullOrEmpty($TenantId) -or [string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret)) {
            Write-Log "Using Azure Automation variables for authentication"
            $TenantId = Get-AutomationVariable -Name 'TenantId'
            $ClientId = Get-AutomationVariable -Name 'ClientId'
            $ClientSecret = Get-AutomationVariable -Name 'ClientSecret'
        }
        
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        $body = @{
            client_id     = $ClientId
            scope         = "https://graph.microsoft.com/.default"
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }
        
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Log "Successfully acquired token" 
        return $response.access_token
    }
    catch {
        Write-Log "Failed to acquire token: $_" -Type "ERROR"
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
                } else {
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
        [string]$OSType,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        $filter = switch ($OSType) {
            "Windows" { "operatingSystem eq 'Windows'" }
            "iOS" { "operatingSystem eq 'iOS'" }
            "Android" { "operatingSystem eq 'Android'" }
            "Linux" { "operatingSystem eq 'Linux'" }
            "All" { "operatingSystem eq 'Windows' or operatingSystem eq 'iOS' or operatingSystem eq 'Android' or operatingSystem eq 'Linux'" }
            default { "operatingSystem eq 'Windows' or operatingSystem eq 'iOS' or operatingSystem eq 'Android' or operatingSystem eq 'Linux'" }
        }

        Write-Log "Retrieving Intune devices with filter: $filter"
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$filter"
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
        
        $categoryLookup = @{}
        foreach ($category in $categories.value) {
            $categoryLookup[$category.displayName] = $category.id
        }
        
        return $categoryLookup
    }
    catch {
        Write-Log "Failed to retrieve device categories: $_" -Type "ERROR"
        throw "Failed to retrieve device categories: $_"
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
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/users"
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

function Get-UserDetails {
    param (
        [string]$Token,
        [string]$UserId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving details for user $UserId..."
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId`?`$select=id,displayName,department"
        $user = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        return $user
    }
    catch {
        Write-Log "Failed to retrieve details for user $UserId`: $_" -Type "ERROR"
        return $null
    }
}

function Update-DeviceCategory {
    param (
        [string]$Token,
        [string]$DeviceId,
        [string]$CategoryId,
        [string]$CategoryName,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would update device category for device $DeviceId to: $CategoryName (ID: $CategoryId)" -Type "WHATIF"
            return $true
        }
        else {
            Write-Log "Updating device category for device $DeviceId to: $CategoryName (ID: $CategoryId)"
            
            $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCategory/`$ref"
            $body = @{
                "@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$CategoryId"
            }
            
            $jsonBody = $body | ConvertTo-Json -Depth 10
            Write-Log "Request body: $jsonBody"
            Write-Log "Request URI: $uri"
            
            Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "PUT" -Body $body -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            Write-Log "Successfully updated device category"
            return $true
        }
    }
    catch {
        Write-Log "Failed to update device category for device $DeviceId`: $_" -Type "ERROR"
        return $false
    }
}

function Process-DeviceBatch {
    param (
        [string]$Token,
        [array]$Devices,
        [hashtable]$CategoryLookup,
        [switch]$WhatIf,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds
    )
    
    $batchUpdateCount = 0
    $batchErrorCount = 0
    $batchSkippedCount = 0
    $batchMatchCount = 0
    
    foreach ($device in $Devices) {
        try {
            $deviceName = $device.deviceName
            $deviceId = $device.id
            $category = $device.deviceCategoryDisplayName
            $osType = $device.operatingSystem
            
            if ($osType) {
                if (-not $Stats.OSTypeStats.ContainsKey($osType)) {
                    $Stats.OSTypeStats[$osType] = @{
                        "Total" = 0
                        "Updated" = 0
                        "Matched" = 0
                        "Skipped" = 0
                        "Errors" = 0
                    }
                }
                $Stats.OSTypeStats[$osType]["Total"]++
            }
            
            Write-Log "Processing device: $deviceName (ID: $deviceId, OS: $osType)"
            Write-Log "Current Category: '$category'"
            
            $primaryUser = Get-DevicePrimaryUser -Token $Token -DeviceId $deviceId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($null -ne $primaryUser) {
                $userDetails = Get-UserDetails -Token $Token -UserId $primaryUser.id -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                
                if ($null -ne $userDetails -and -not [string]::IsNullOrEmpty($userDetails.department)) {
                    $userDepartment = $userDetails.department
                    Write-Log "Found department '$userDepartment' for user $($userDetails.displayName)"
                    
                    if ($CategoryLookup.ContainsKey($userDepartment)) {
                        $categoryId = $CategoryLookup[$userDepartment]
                        
                        if ([string]::IsNullOrEmpty($category) -or 
                            $category -eq "Unassigned" -or 
                            $category -eq "Unknown" -or 
                            $category -ne $userDepartment) {
                            
                            if (![string]::IsNullOrEmpty($category) -and 
                                $category -ne "Unassigned" -and 
                                $category -ne "Unknown" -and 
                                $category -ne $userDepartment) {
                                Write-Log "Device $deviceName has category '$category' which doesn't match user department '$userDepartment'. Updating..." -Type "WARNING"
                            } else {
                                Write-Log "Device $deviceName has no valid category assigned. Updating to match user department..." -Type "WARNING"
                            }
                            
                            $updateResult = Update-DeviceCategory -Token $Token -DeviceId $deviceId -CategoryId $categoryId -CategoryName $userDepartment -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                            
                            if ($updateResult) {
                                if ($WhatIf) {
                                    Write-Log "Would have updated device category to '$userDepartment' based on user department for device $deviceName" -Type "WHATIF"
                                }
                                else {
                                    Write-Log "Successfully updated device category to '$userDepartment' based on user department for device $deviceName"
                                }
                                $batchUpdateCount++
                                $Stats.UpdatedCount++
                                if ($Stats.OSTypeStats.ContainsKey($osType)) {
                                    $Stats.OSTypeStats[$osType]["Updated"]++
                                }
                            }
                            else {
                                Write-Log "Failed to update device category for device $deviceName" -Type "ERROR"
                                $batchErrorCount++
                                $Stats.ErrorCount++
                                if ($Stats.OSTypeStats.ContainsKey($osType)) {
                                    $Stats.OSTypeStats[$osType]["Errors"]++
                                }
                            }
                        }
                        else {
                            Write-Log "Device $deviceName already has category set to '$category' which matches user department. No action needed."
                            $batchMatchCount++
                            $Stats.MatchCount++
                            if ($Stats.OSTypeStats.ContainsKey($osType)) {
                                $Stats.OSTypeStats[$osType]["Matched"]++
                            }
                        }
                    }
                    else {
                        Write-Log "Department '$userDepartment' does not exist as a device category in Intune. Skipping." -Type "WARNING"
                        $batchSkippedCount++
                        $Stats.SkippedCount++
                        if ($Stats.OSTypeStats.ContainsKey($osType)) {
                            $Stats.OSTypeStats[$osType]["Skipped"]++
                        }
                    }
                }
                else {
                    Write-Log "No department information found for the primary user of device $deviceName. Skipping." -Type "WARNING"
                    $batchSkippedCount++
                    $Stats.SkippedCount++
                    if ($Stats.OSTypeStats.ContainsKey($osType)) {
                        $Stats.OSTypeStats[$osType]["Skipped"]++
                    }
                }
            }
            else {
                Write-Log "No primary user found for device $deviceName. Keeping existing category." -Type "WARNING"
                $batchSkippedCount++
                $Stats.SkippedCount++
                if ($Stats.OSTypeStats.ContainsKey($osType)) {
                    $Stats.OSTypeStats[$osType]["Skipped"]++
                }
            }
        }
        catch {
            Write-Log "Error processing device $($device.deviceName): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
            $osType = $device.operatingSystem
            if ($Stats.OSTypeStats.ContainsKey($osType)) {
                $Stats.OSTypeStats[$osType]["Errors"]++
            }
        }
    }
    
    return @{
        UpdatedCount = $batchUpdateCount
        SkippedCount = $batchSkippedCount
        MatchCount = $batchMatchCount
        ErrorCount = $batchErrorCount
    }
}

try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO CHANGES WILL BE MADE ===" -Type "WHATIF"
    }
    
    Write-Log "=== Intune Device Category Update Started ==="
    Write-Log "Operating System Type: $OSType"
    Write-Log "Batch Size: $BatchSize"
    Write-Log "Batch Delay: $BatchDelaySeconds seconds"
    
    $startTime = Get-Date
    
    $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    
    $categoryLookup = Get-IntuneDeviceCategories -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    Write-Log "Available device categories:"
    foreach ($cat in $categoryLookup.Keys) {
        Write-Log "- $cat (ID: $($categoryLookup[$cat]))"
    }
    
    $devices = Get-IntuneDevices -Token $token -OSType $OSType -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    $stats = @{
        UpdatedCount = 0
        ErrorCount = 0
        SkippedCount = 0
        MatchCount = 0
        OSTypeStats = @{}
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
        
        $batchResult = Process-DeviceBatch -Token $token -Devices $currentBatch -CategoryLookup $categoryLookup -WhatIf:$WhatIf -Stats $stats -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Batch $($batchNum+1) results: $($batchResult.UpdatedCount) updated, $($batchResult.MatchCount) already correct, $($batchResult.SkippedCount) skipped, $($batchResult.ErrorCount) errors"
        
        if ($batchNum -lt $batches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Intune Device Category Update Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO CHANGES WERE MADE ===" -Type "WHATIF"
    }
    
    Write-Log "Overall Summary:"
    Write-Log "Devices processed: $totalDevices"
    Write-Log "Already categorized: $($stats.MatchCount)"

    if ($WhatIf) {
        Write-Log "Would be updated: $($stats.UpdatedCount)" -Type "WHATIF"
    } else {
        Write-Log "Updated: $($stats.UpdatedCount)"
    }

    Write-Log "Skipped (no primary user, no department, or department not a category): $($stats.SkippedCount)"
    Write-Log "Errors: $($stats.ErrorCount)"
    
    foreach ($os in $stats.OSTypeStats.Keys | Sort-Object) {
        Write-Log "$os Device Summary:"
        Write-Log "- Total $os devices: $($stats.OSTypeStats[$os]["Total"])"
        Write-Log "- Already categorized: $($stats.OSTypeStats[$os]["Matched"])"
        
        if ($WhatIf) {
            Write-Log "- Would be updated: $($stats.OSTypeStats[$os]["Updated"])" -Type "WHATIF"
        } else {
            Write-Log "- Updated: $($stats.OSTypeStats[$os]["Updated"])"
        }
        
        Write-Log "- Skipped: $($stats.OSTypeStats[$os]["Skipped"])"
        Write-Log "- Errors: $($stats.OSTypeStats[$os]["Errors"])"
    }

    $outputProperties = [ordered]@{
        TotalDevices = $totalDevices
        AlreadyCategorized = $stats.MatchCount
        Updated = $stats.UpdatedCount
        Skipped = $stats.SkippedCount
        Errors = $stats.ErrorCount
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        BatchesProcessed = $batches
    }
    
    foreach ($os in $stats.OSTypeStats.Keys | Sort-Object) {
        $outputProperties["${os}Devices"] = $stats.OSTypeStats[$os]["Total"]
        $outputProperties["${os}Updated"] = $stats.OSTypeStats[$os]["Updated"]
        $outputProperties["${os}Matched"] = $stats.OSTypeStats[$os]["Matched"]
        $outputProperties["${os}Skipped"] = $stats.OSTypeStats[$os]["Skipped"]
        $outputProperties["${os}Errors"] = $stats.OSTypeStats[$os]["Errors"]
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