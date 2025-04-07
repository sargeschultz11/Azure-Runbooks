<#
.SYNOPSIS
    Generates a report of all devices with a specific application installed and uploads it to a SharePoint document library.

.DESCRIPTION
    This Azure Runbook script connects to the Microsoft Graph API using System-Assigned Managed Identity,
    retrieves all devices that have a specific application installed (identified by app ID), 
    exports the data to an Excel file, and uploads the file to a specified SharePoint document library.
    It includes robust error handling and implements throttling detection with retry logic.

.PARAMETER AppId
    The ID of the application to search for. This is the ID from the Intune Discovered Apps report.

.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the report will be uploaded.

.PARAMETER SharePointDriveId
    The ID of the document library drive where the report will be uploaded.

.PARAMETER FolderPath
    Optional. The folder path within the document library where the report will be uploaded.
    If not specified, the file will be uploaded to the root of the document library.

.PARAMETER IncludeDeviceDetails
    Optional. If specified, the script will include more detailed device information in the report.
    This requires additional API calls and may increase execution time.

.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.

.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.

.PARAMETER TeamsWebhookUrl
    Optional. Microsoft Teams webhook URL for sending notifications about the report.

.NOTES
    File Name: Get-DevicesWithAppReport.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-07

    Requires -Modules ImportExcel

#>

param(
    [Parameter(Mandatory = $true)]
    [string]$AppId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointDriveId,
    
    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDeviceDetails = $false,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl
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

function Get-AppDetails {
    param (
        [string]$Token,
        [string]$AppId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving details for application ID: $AppId"
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/$AppId"
        
        $app = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        if ($null -eq $app) {
            Write-Log "No application found with ID: $AppId" -Type "WARNING"
            return $null
        }
        
        Write-Log "Found application: $($app.displayName) (Publisher: $($app.publisher), Version: $($app.version))"
        return $app
    }
    catch {
        Write-Log "Failed to retrieve application details: $_" -Type "ERROR"
        return $null
    }
}

function Get-DevicesWithApp {
    param (
        [string]$Token,
        [string]$AppId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving devices with application ID: $AppId installed"
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/$AppId/managedDevices"
        
        $devices = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $devices += $response.value
        $totalDevices = $devices.Count
        Write-Log "Retrieved $totalDevices devices in first batch"
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of devices..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $devices += $response.value
            Write-Log "Retrieved $($response.value.Count) additional devices, total: $($devices.Count)"
        }
        
        Write-Log "Retrieved a total of $($devices.Count) devices with the application installed"
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve devices with application: $_" -Type "ERROR"
        throw "Failed to retrieve devices with application: $_"
    }
}

function Get-EnhancedDeviceDetails {
    param (
        [string]$Token,
        [array]$Devices,
        [PSObject]$AppDetails,
        [bool]$IncludeDetails = $false,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        $totalDevices = $Devices.Count
        $enhancedDevices = @()
        $counter = 0
        
        Write-Log "Enhancing device information for $totalDevices devices..."
        
        foreach ($device in $Devices) {
            $counter++
            $percentage = [math]::Round(($counter / $totalDevices) * 100, 1)
            
            if (($counter % 10 -eq 0) -or ($counter -eq $totalDevices)) {
                Write-Log "Processing device $counter of $totalDevices ($percentage%)"
            }
            
            $deviceId = $device.id
            $deviceUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$deviceId"
            
            try {
                $deviceDetails = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $deviceUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                
                $enhancedDevice = [PSCustomObject]@{
                    DeviceName = $deviceDetails.deviceName
                    DeviceId = $deviceDetails.id
                    SerialNumber = $deviceDetails.serialNumber
                    Model = $deviceDetails.model
                    Manufacturer = $deviceDetails.manufacturer
                    OS = $deviceDetails.operatingSystem
                    OSVersion = $deviceDetails.osVersion
                    ManagementState = $deviceDetails.managementState
                    PrimaryUser = $deviceDetails.userPrincipalName
                    UserDisplayName = $deviceDetails.userDisplayName
                    DeviceOwnership = $deviceDetails.managedDeviceOwnerType
                    LastSyncDateTime = $deviceDetails.lastSyncDateTime
                    EnrolledDateTime = $deviceDetails.enrolledDateTime
                    DeviceCategory = $deviceDetails.deviceCategoryDisplayName
                    AppName = $AppDetails.displayName
                    AppPublisher = $AppDetails.publisher
                    AppVersion = $AppDetails.version
                    AppId = $AppDetails.id
                    AppDetectionDateTime = $deviceDetails.lastSyncDateTime # Using last sync as proxy since actual detection time not available
                }
                
                # Add additional device details if requested
                if ($IncludeDetails) {
                    # Get device compliance info
                    $complianceUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$deviceId/deviceCompliancePolicyStates"
                    
                    try {
                        $compliance = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $complianceUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                        $complianceStatus = if ($compliance.value.Count -gt 0) {
                            ($compliance.value | ForEach-Object { $_.state }) -join ', '
                        } else {
                            "No compliance policies assigned"
                        }
                        
                        $enhancedDevice | Add-Member -MemberType NoteProperty -Name "ComplianceStatus" -Value $complianceStatus
                    }
                    catch {
                        Write-Log "Failed to retrieve compliance info for device $($deviceDetails.deviceName): $_" -Type "WARNING"
                        $enhancedDevice | Add-Member -MemberType NoteProperty -Name "ComplianceStatus" -Value "Error retrieving compliance info"
                    }
                }
                
                $enhancedDevices += $enhancedDevice
                
                if ($counter % 10 -eq 0 -and $counter -lt $totalDevices) {
                    Start-Sleep -Seconds 1
                }
            }
            catch {
                Write-Log "Failed to retrieve detailed information for device ID $deviceId`: $_" -Type "WARNING"
                
                $enhancedDevice = [PSCustomObject]@{
                    DeviceName = "Unknown"
                    DeviceId = $deviceId
                    SerialNumber = "Unknown"
                    Model = "Unknown"
                    Manufacturer = "Unknown"
                    OS = "Unknown"
                    OSVersion = "Unknown"
                    ManagementState = "Unknown"
                    PrimaryUser = "Unknown"
                    UserDisplayName = "Unknown"
                    DeviceOwnership = "Unknown"
                    LastSyncDateTime = "Unknown"
                    EnrolledDateTime = "Unknown"
                    DeviceCategory = "Unknown"
                    AppName = $AppDetails.displayName
                    AppPublisher = $AppDetails.publisher
                    AppVersion = $AppDetails.version
                    AppId = $AppDetails.id
                    AppDetectionDateTime = "Unknown"
                }
                
                if ($IncludeDetails) {
                    $enhancedDevice | Add-Member -MemberType NoteProperty -Name "ComplianceStatus" -Value "Error retrieving device details"
                }
                
                $enhancedDevices += $enhancedDevice
            }
        }
        
        Write-Log "Enhanced device information processing completed for $totalDevices devices"
        return $enhancedDevices
    }
    catch {
        Write-Log "Failed to enhance device information: $_" -Type "ERROR"
        throw "Failed to enhance device information: $_"
    }
}

function Export-DataToExcel {
    param (
        [array]$Data,
        [string]$FilePath,
        [PSObject]$AppDetails,
        [bool]$IncludedDetailedInfo = $false
    )
    
    try {
        Write-Log "Exporting data to Excel file: $FilePath"
        
        if (-not (Get-Module -Name ImportExcel)) {
            Import-Module ImportExcel -ErrorAction Stop
        }
        
        $currentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $reportInfo = [PSCustomObject]@{
            'Report Generated' = $currentDate
            'Generated By'     = $env:COMPUTERNAME
            'Number of Devices' = $Data.Count
            'Application Name' = $AppDetails.displayName
            'Publisher' = $AppDetails.publisher
            'Version' = $AppDetails.version
            'Device Count' = $AppDetails.deviceCount
        }
        
        $excelParams = @{
            Path          = $FilePath
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            AutoFilter    = $true
            WorksheetName = "Devices With App"
            TableName     = "DevicesWithApp"
            PassThru      = $true
        }
        
        $properties = @(
            'DeviceName', 
            'PrimaryUser', 
            'UserDisplayName',
            'OS', 
            'OSVersion',
            'Model',
            'Manufacturer',
            'SerialNumber',
            'DeviceOwnership',
            'LastSyncDateTime',
            'EnrolledDateTime',
            'DeviceCategory'
        )
        
        if ($IncludedDetailedInfo) {
            $properties += @('ComplianceStatus')
        }
        
        $properties += @(
            'AppName',
            'AppPublisher',
            'AppVersion',
            'AppId'
        )
        
        $excel = $Data | Select-Object $properties | Export-Excel @excelParams
        
        $summarySheet = $excel.Workbook.Worksheets.Add("Summary")
        $summarySheet.Cells["A1"].Value = "App Installation Report Summary"
        $summarySheet.Cells["A1:B1"].Merge = $true
        $summarySheet.Cells["A1:B1"].Style.Font.Bold = $true
        $summarySheet.Cells["A1:B1"].Style.Font.Size = 14
        
        $row = 3
        
        $summarySheet.Cells["A$row"].Value = "Report Generated"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Report Generated'
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Generated By"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Generated By'
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Number of Devices"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Number of Devices'
        $row++
        
        $row++
        $summarySheet.Cells["A$row"].Value = "Application Information"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Application Name"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Application Name'
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Publisher"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Publisher'
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Version"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Version'
        $row++
        
        $summarySheet.Cells["A$row"].Value = "Total Device Count"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Device Count'
        $row++
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Operating System Distribution"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $osSummary = $Data | Group-Object -Property OS | Sort-Object -Property Count -Descending
        
        $summarySheet.Cells["A$row"].Value = "OS"
        $summarySheet.Cells["B$row"].Value = "Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($os in $osSummary) {
            $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($os.Name)) { "(Unknown)" } else { $os.Name }
            $summarySheet.Cells["B$row"].Value = $os.Count
            $row++
        }
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Device Ownership Distribution"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $ownershipSummary = $Data | Group-Object -Property DeviceOwnership | Sort-Object -Property Count -Descending
        
        $summarySheet.Cells["A$row"].Value = "Ownership Type"
        $summarySheet.Cells["B$row"].Value = "Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($ownership in $ownershipSummary) {
            $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($ownership.Name)) { "(Unknown)" } else { $ownership.Name }
            $summarySheet.Cells["B$row"].Value = $ownership.Count
            $row++
        }
        
        $categorySummary = $Data | Group-Object -Property DeviceCategory | Where-Object { -not [string]::IsNullOrEmpty($_.Name) } | Sort-Object -Property Count -Descending
        
        if ($categorySummary.Count -gt 0) {
            $row += 2
            $summarySheet.Cells["A$row"].Value = "Device Category Distribution"
            $summarySheet.Cells["A$row"].Style.Font.Bold = $true
            $row++
            
            $summarySheet.Cells["A$row"].Value = "Category"
            $summarySheet.Cells["B$row"].Value = "Count"
            $summarySheet.Cells["A$row"].Style.Font.Bold = $true
            $summarySheet.Cells["B$row"].Style.Font.Bold = $true
            $row++
            
            foreach ($category in $categorySummary) {
                $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($category.Name)) { "(Not Categorized)" } else { $category.Name }
                $summarySheet.Cells["B$row"].Value = $category.Count
                $row++
            }
        }
        
        if ($IncludedDetailedInfo) {
            $complianceSummary = $Data | Group-Object -Property ComplianceStatus | Sort-Object -Property Count -Descending
            
            if ($complianceSummary.Count -gt 0) {
                $row += 2
                $summarySheet.Cells["A$row"].Value = "Compliance Status Distribution"
                $summarySheet.Cells["A$row"].Style.Font.Bold = $true
                $row++
                
                $summarySheet.Cells["A$row"].Value = "Compliance Status"
                $summarySheet.Cells["B$row"].Value = "Count"
                $summarySheet.Cells["A$row"].Style.Font.Bold = $true
                $summarySheet.Cells["B$row"].Style.Font.Bold = $true
                $row++
                
                foreach ($compliance in $complianceSummary) {
                    $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($compliance.Name)) { "(Unknown)" } else { $compliance.Name }
                    $summarySheet.Cells["B$row"].Value = $compliance.Count
                    $row++
                }
            }
        }
        
        $summarySheet.Column(1).AutoFit()
        $summarySheet.Column(2).AutoFit()
        
        try {
            $excel.Workbook.Worksheets[0].View.TabSelected = $false
            $summarySheet.View.TabSelected = $true
            $excel.Workbook.View.ActiveTab = 1
        }
        catch {
            Write-Log "Could not set the active sheet, but this is not critical for report generation" -Type "WARNING"
        }
        
        $excel.Save()
        $excel.Dispose()
        
        Write-Log "Excel file created successfully at: $FilePath"
    }
    catch {
        Write-Log "Failed to export data to Excel: $_" -Type "ERROR"
        throw "Failed to export data to Excel: $_"
    }
}

function Send-TeamsNotification {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebhookUrl,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ReportData
    )
    
    try {
        Write-Log "Sending notification to Microsoft Teams..."
        
        $executionTime = [math]::Round($ReportData.ExecutionTimeMinutes, 2)
        
        $adaptiveCard = @{
            type        = "message"
            attachments = @(
                @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    contentUrl  = $null
                    content     = @{
                        "$schema" = "http://adaptivecards.io/schemas/adaptive-card.json"
                        type      = "AdaptiveCard"
                        version   = "1.2"
                        msTeams   = @{
                            width = "full"
                        }
                        body      = @(
                            @{
                                type   = "TextBlock"
                                size   = "Large"
                                weight = "Bolder"
                                text   = "App Installation Report"
                                wrap   = $true
                                color  = "Default"
                            },
                            @{
                                type     = "TextBlock"
                                spacing  = "None"
                                text     = "Report generated on $($ReportData.Timestamp)"
                                wrap     = $true
                                isSubtle = $true
                            },
                            @{
                                type  = "FactSet"
                                facts = @(
                                    @{
                                        title = "Application:"
                                        value = "$($ReportData.AppName) $($ReportData.AppVersion)"
                                    },
                                    @{
                                        title = "Publisher:"
                                        value = $ReportData.AppPublisher
                                    },
                                    @{
                                        title = "Devices Found:"
                                        value = $ReportData.DevicesCount.ToString()
                                    },
                                    @{
                                        title = "Execution Time:"
                                        value = "$executionTime minutes"
                                    }
                                )
                            }
                        )
                        actions   = @(
                            @{
                                type  = "Action.OpenUrl"
                                title = "View Report"
                                url   = $ReportData.ReportUrl
                            }
                        )
                    }
                }
            )
        }
        
        $body = ConvertTo-Json -InputObject $adaptiveCard -Depth 20
        
        $params = @{
            Uri         = $WebhookUrl
            Method      = "POST"
            Body        = $body
            ContentType = "application/json"
        }
        
        $response = Invoke-RestMethod @params
        
        Write-Log "Teams notification sent successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send Teams notification: $_" -Type "WARNING"
        return $false
    }
}

function Upload-FileToSharePoint {
    param (
        [string]$Token,
        [string]$SiteId,
        [string]$DriveId,
        [string]$FolderPath,
        [string]$FilePath,
        [string]$FileName,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Uploading file to SharePoint..."
        
        if (-not (Test-Path $FilePath)) {
            throw "File does not exist at path: $FilePath"
        }
        
        $fileInfo = Get-Item -Path $FilePath
        $fileSize = $fileInfo.Length
        
        Write-Log "File size: $fileSize bytes"
        
        if ($fileSize -gt 4000000) {
            Write-Log "Using large file upload session approach for file over 4MB"
            
            $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
                $FileName
            }
            else {
                "$FolderPath/$FileName"
            }
            
            $createSessionUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$uploadPath`:/createUploadSession"
            $createSessionBody = @{
                item = @{
                    "@microsoft.graph.conflictBehavior" = "replace"
                }
            }
            
            $uploadSession = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $createSessionUri -Method "POST" -Body $createSessionBody -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if (-not $uploadSession -or -not $uploadSession.uploadUrl) {
                throw "Failed to create upload session"
            }
            
            $chunkSize = 3 * 1024 * 1024
            $fileStream = [System.IO.File]::OpenRead($FilePath)
            $buffer = New-Object byte[] $chunkSize
            $bytesRead = 0
            $position = 0
            
            try {
                while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    if ($bytesRead -lt $buffer.Length) {
                        $actualBuffer = New-Object byte[] $bytesRead
                        [Array]::Copy($buffer, $actualBuffer, $bytesRead)
                        $buffer = $actualBuffer
                    }
                    
                    $contentRange = "bytes $position-$($position + $bytesRead - 1)/$fileSize"
                    $headers = @{
                        "Authorization" = "Bearer $Token"
                        "Content-Range" = $contentRange
                    }
                    
                    $uploadChunkParams = @{
                        Uri         = $uploadSession.uploadUrl
                        Method      = "PUT"
                        Headers     = $headers
                        Body        = $buffer
                        ContentType = "application/octet-stream"
                    }
                    
                    Write-Log "Uploading chunk: $contentRange"
                    
                    $retryCount = 0
                    $success = $false
                    
                    while (-not $success -and $retryCount -lt $MaxRetries) {
                        try {
                            $response = Invoke-RestMethod @uploadChunkParams
                            $success = $true
                            
                            if ($response.id) {
                                Write-Log "File upload completed. WebUrl: $($response.webUrl)"
                                return $response
                            }
                        }
                        catch {
                            $retryCount++
                            $backoffSeconds = $InitialBackoffSeconds * [Math]::Pow(2, $retryCount - 1)
                            
                            if ($retryCount -lt $MaxRetries) {
                                Write-Log "Chunk upload failed. Retrying in $backoffSeconds seconds. Attempt $retryCount of $MaxRetries. Error: $_" -Type "WARNING"
                                Start-Sleep -Seconds $backoffSeconds
                            }
                            else {
                                throw $_
                            }
                        }
                    }
                    
                    $position += $bytesRead
                }
            }
            finally {
                $fileStream.Close()
                $fileStream.Dispose()
            }
            
            throw "File upload did not complete properly"
        }
        else {
            Write-Log "Using direct upload approach for smaller file"
            
            $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
                $FileName
            }
            else {
                "$FolderPath/$FileName"
            }
            
            $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$uploadPath`:/content"
            
            Write-Log "Uploading file to: $uploadUri"
            
            $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
            
            $headers = @{
                "Authorization" = "Bearer $Token"
            }
            
            $params = @{
                Uri         = $uploadUri
                Method      = "PUT"
                Headers     = $headers
                Body        = $fileBytes
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
            
            $response = Invoke-RestMethod @params
            
            Write-Log "File uploaded successfully. WebUrl: $($response.webUrl)"
            return $response
        }
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Type "ERROR"
        throw "Failed to upload file to SharePoint: $_"
    }
}

# Main execution block
try {
    $startTime = Get-Date
    Write-Log "=== Intune App Installation Report Generation Started ==="
    Write-Log "Application ID: $AppId"
    Write-Log "Include Detailed Device Info: $IncludeDeviceDetails"
    
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "ImportExcel module not found. Installing..." -Type "WARNING"
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
    
    $token = Get-MsGraphToken
    
    $appDetails = Get-AppDetails -Token $token -AppId $AppId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    if ($null -eq $appDetails) {
        throw "Could not find application with ID: $AppId. Please verify the App ID is correct."
    }
    
    Write-Log "Processing application: $($appDetails.displayName) (Publisher: $($appDetails.publisher), Version: $($appDetails.version))"
    
    $devicesWithApp = Get-DevicesWithApp -Token $token -AppId $AppId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    if ($devicesWithApp.Count -eq 0) {
        Write-Log "No devices found with the application installed" -Type "WARNING"
        
        $result = [PSCustomObject]@{
            AppId              = $AppId
            AppName            = $appDetails.displayName
            AppPublisher       = $appDetails.publisher
            AppVersion         = $appDetails.version
            DevicesCount       = 0
            ReportName         = "No devices found"
            ExecutionTimeMinutes = (Get-Date - $startTime).TotalMinutes
            Timestamp          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ReportUrl          = "N/A"
            NotificationSent   = $false
        }
        
        return $result
    }
    
    Write-Log "Found $($devicesWithApp.Count) devices with the application installed"
    
    $enhancedDevices = Get-EnhancedDeviceDetails -Token $token -Devices $devicesWithApp -AppDetails $appDetails -IncludeDetails $IncludeDeviceDetails -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    $currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $appNameSafe = $appDetails.displayName -replace '[\\\/\:\*\?\"\<\>\|]', '_'
    $reportName = "App_Installation_Report_${appNameSafe}_$currentDate.xlsx"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $reportName)
    
    Export-DataToExcel -Data $enhancedDevices -FilePath $tempPath -AppDetails $appDetails -IncludedDetailedInfo $IncludeDeviceDetails
    
    $uploadResult = Upload-FileToSharePoint -Token $token -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FolderPath $FolderPath -FilePath $tempPath -FileName $reportName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    if (Test-Path -Path $tempPath) {
        Remove-Item -Path $tempPath -Force
        Write-Log "Temporary file removed: $tempPath"
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Intune App Installation Report Generation Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Report URL: $($uploadResult.webUrl)"
    
    $result = [PSCustomObject]@{
        AppId                = $AppId
        AppName              = $appDetails.displayName
        AppPublisher         = $appDetails.publisher
        AppVersion           = $appDetails.version
        DevicesCount         = $enhancedDevices.Count
        ReportName           = $reportName
        ExecutionTimeMinutes = $duration.TotalMinutes
        Timestamp            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ReportUrl            = $uploadResult.webUrl
        NotificationSent     = $false
    }
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        $notificationSent = Send-TeamsNotification -WebhookUrl $TeamsWebhookUrl -ReportData $result
        $result.NotificationSent = $notificationSent
    }
    
    return $result
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    
    if ($tempPath -and (Test-Path -Path $tempPath)) {
        Remove-Item -Path $tempPath -Force
        Write-Log "Temporary file removed: $tempPath"
    }
    
    throw $_
}
finally {
    Write-Log "Script execution completed"
}