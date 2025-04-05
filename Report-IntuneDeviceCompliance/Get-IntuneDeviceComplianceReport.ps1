<#
.SYNOPSIS
    Generates a report of device compliance status from Microsoft Intune and uploads it to a SharePoint document library.

.DESCRIPTION
    This Azure Runbook script connects to the Microsoft Graph API using client credentials (App Registration),
    retrieves device compliance status from Intune, exports the data to an Excel file,
    and uploads the file to a specified SharePoint document library.
    It supports batch processing and implements throttling detection with retry logic.

.PARAMETER TenantId
    The Azure AD tenant ID. If not provided, will be retrieved from Automation variables.

.PARAMETER ClientId
    The App Registration's client ID. If not provided, will be retrieved from Automation variables.

.PARAMETER ClientSecret
    The App Registration's client secret. If not provided, will be retrieved from Automation variables.
    
.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the report will be uploaded.

.PARAMETER SharePointDriveId
    The ID of the document library drive where the report will be uploaded.

.PARAMETER FolderPath
    Optional. The folder path within the document library where the report will be uploaded.
    If not specified, the file will be uploaded to the root of the document library.

.PARAMETER BatchSize
    Optional. Number of records to process in each batch. Default is 100.

.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.

.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.

.NOTES
    File Name: Get-IntuneDeviceComplianceReport.ps1
    Version: 1.0
    Created: 2025-04-04

    Requires -Modules ImportExcel

#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseManagedIdentity,

    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointDriveId,
    
    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "",
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 100,
    
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
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [switch]$UseManagedIdentity
    )
    
    try {
        Write-Log "Attempting to acquire Microsoft Graph API token..."
        
        if ($UseManagedIdentity) {
            Write-Log "Using Managed Identity for authentication"
            
            try {
                if (Get-Module -ListAvailable -Name Az.Accounts) {
                    Import-Module Az.Accounts
                    Write-Log "Using Az module for Managed Identity authentication"
                    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
                }
                elseif (Get-Module -ListAvailable -Name AzureRM.Profile) {
                    Import-Module AzureRM.Profile
                    Write-Log "Using AzureRM module for Managed Identity authentication"
                    $token = (Get-AzureRmAccessToken -ResourceUrl "https://graph.microsoft.com").AccessToken
                }
                else {
                    Write-Log "Using Azure Automation's managed identity capabilities"
                    Connect-AzAccount -Identity
                    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
                }
                
                if ([string]::IsNullOrEmpty($token)) {
                    throw "Could not acquire token using Managed Identity"
                }
                
                Write-Log "Successfully acquired token using Managed Identity"
                return $token
            }
            catch {
                throw "Failed to authenticate using Managed Identity: $_"
            }
        }
        else {
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
            Write-Log "Successfully acquired token using App Registration" 
            return $response.access_token
        }
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
    $headers = @{
        Authorization = "Bearer $Token"
        ConsistencyLevel = "eventual"
    }
    
    $params = @{
        Uri         = $Uri
        Headers     = $headers
        Method      = $Method
        ContentType = $ContentType
    }
    
    if ($null -ne $Body -and $Method -ne "GET") {
        if ($ContentType -eq "application/json") {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
        else {
            $params.Add("Body", $Body)
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

function Get-IntuneDeviceComplianceStatus {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5,
        [int]$BatchSize = 100
    )
    
    try {
        Write-Log "Retrieving device compliance status from Intune..."
        
        $devices = @()
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=id,deviceName,managedDeviceOwnerType,deviceType,operatingSystem,osVersion,complianceState,lastSyncDateTime,emailAddress,userPrincipalName,serialNumber,model,manufacturer,enrolledDateTime,userDisplayName&`$top=$BatchSize"
        
        $count = 0
        $batchCount = 0
        
        do {
            $batchCount++
            Write-Log "Retrieving batch $batchCount of devices..."
            
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($response.value.Count -gt 0) {
                $devices += $response.value
                $count += $response.value.Count
                Write-Log "Retrieved $($response.value.Count) devices in this batch, total count: $count"
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($null -ne $uri)
        
        Write-Log "Retrieved a total of $($devices.Count) devices"
        
        Write-Log "Retrieving compliance policies..."
        $policies = Invoke-MsGraphRequestWithRetry -Token $Token -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $policyLookup = @{}
        foreach ($policy in $policies.value) {
            $policyLookup[$policy.id] = $policy.displayName
        }
        
        Write-Log "Enhancing device information with compliance policy details..."
        $enhancedDevices = @()
        
        foreach ($device in $devices) {
            $enhancedDevice = $device.PSObject.Copy()
            
            $deviceId = $device.id
            $policyStatesUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceId/deviceCompliancePolicyStates?`$expand=settings"
            
            try {
                $policyStates = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $policyStatesUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                
                if ($policyStates.value.Count -gt 0) {
                    $policyNames = @()
                    $policyStatuses = @()
                    
                    foreach ($policyState in $policyStates.value) {
                        $policyId = $policyState.policyId
                        if ($policyLookup.ContainsKey($policyId)) {
                            $policyName = $policyLookup[$policyId]
                        } else {
                            $policyName = "Unknown Policy ($policyId)"
                        }
                        
                        $policyNames += $policyName
                        $policyStatuses += "$policyName : $($policyState.state)"
                    }
                    
                    $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyNames" -Value ($policyNames -join ", ")
                    $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyStatuses" -Value ($policyStatuses -join ", ")
                } else {
                    $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyNames" -Value "No Policies Assigned"
                    $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyStatuses" -Value "No Policy States"
                }
            }
            catch {
                Write-Log "Error retrieving compliance policy states for device $deviceId : $_" -Type "WARNING"
                $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyNames" -Value "Error Retrieving Policies"
                $enhancedDevice | Add-Member -MemberType NoteProperty -Name "compliancePolicyStatuses" -Value "Error Retrieving Policy States"
            }
            
            $enhancedDevices += $enhancedDevice
        }
        
        return $enhancedDevices
    }
    catch {
        Write-Log "Failed to retrieve device compliance status: $_" -Type "ERROR"
        throw "Failed to retrieve device compliance status: $_"
    }
}

function Export-DataToExcel {
    param (
        [array]$Data,
        [string]$FilePath
    )
    
    try {
        Write-Log "Exporting data to Excel file: $FilePath"
        
        $reportInfo = [PSCustomObject]@{
            'Report Generated'   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            'Generated By'       = $env:COMPUTERNAME
            'Number of Devices'  = $Data.Count
        }
        
        $excelParams = @{
            Path          = $FilePath
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            AutoFilter    = $true
            WorksheetName = "Device Compliance"
            TableName     = "DeviceComplianceStatus"
            PassThru      = $true
        }
        
        $excel = $Data | Select-Object @{Name='Device Name';Expression={$_.deviceName}}, 
                                      @{Name='User';Expression={$_.userDisplayName}}, 
                                      @{Name='Email';Expression={$_.emailAddress}}, 
                                      @{Name='Device Owner';Expression={$_.managedDeviceOwnerType}}, 
                                      @{Name='Device Type';Expression={$_.deviceType}}, 
                                      @{Name='OS';Expression={$_.operatingSystem}}, 
                                      @{Name='OS Version';Expression={$_.osVersion}}, 
                                      @{Name='Compliance State';Expression={$_.complianceState}}, 
                                      @{Name='Compliance Policies';Expression={$_.compliancePolicyNames}}, 
                                      @{Name='Policy Statuses';Expression={$_.compliancePolicyStatuses}}, 
                                      @{Name='Last Sync';Expression={$_.lastSyncDateTime}}, 
                                      @{Name='Enrolled Date';Expression={$_.enrolledDateTime}}, 
                                      @{Name='Serial Number';Expression={$_.serialNumber}}, 
                                      @{Name='Model';Expression={$_.model}}, 
                                      @{Name='Manufacturer';Expression={$_.manufacturer}} | 
                 Export-Excel @excelParams
        
        $summarySheet = $excel.Workbook.Worksheets.Add("Summary")
        $summarySheet.Cells["A1"].Value = "Compliance Report Summary"
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
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Compliance Status Summary"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $complianceSummary = $Data | Group-Object -Property complianceState | 
                             Sort-Object -Property Name
        
        $summarySheet.Cells["A$row"].Value = "Status"
        $summarySheet.Cells["B$row"].Value = "Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($status in $complianceSummary) {
            $summarySheet.Cells["A$row"].Value = $status.Name
            $summarySheet.Cells["B$row"].Value = $status.Count
            $row++
        }
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Device Type Summary"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $deviceTypeSummary = $Data | Group-Object -Property deviceType | 
                             Sort-Object -Property Count -Descending
        
        $summarySheet.Cells["A$row"].Value = "Device Type"
        $summarySheet.Cells["B$row"].Value = "Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($deviceType in $deviceTypeSummary) {
            $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($deviceType.Name)) { "(Unknown)" } else { $deviceType.Name }
            $summarySheet.Cells["B$row"].Value = $deviceType.Count
            $row++
        }
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "OS Summary"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $osSummary = $Data | Group-Object -Property operatingSystem | 
                     Sort-Object -Property Count -Descending
        
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
        
        $summarySheet.Column(1).AutoFit()
        $summarySheet.Column(2).AutoFit()
        
        try {
            $excel.Workbook.Worksheets[0].View.TabSelected = $false
            $summarySheet.View.TabSelected = $true
        }
        catch {
            try {
                $excel.Workbook.View.ActiveTab = 1
            }
            catch {
                Write-Log "Could not set the active sheet, but this is not critical for report generation" -Type "WARNING"
            }
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
        
        $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
        $fileSize = $fileContent.Length
        
        $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
            $FileName
        } else {
            "$FolderPath/$FileName"
        }
        
        $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$uploadPath`:/content"
        
        Write-Log "Uploading file to: $uploadUri"
        Write-Log "File size: $fileSize bytes"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uploadUri -Method "PUT" -Body $fileContent -ContentType "application/octet-stream" -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "File uploaded successfully. WebUrl: $($response.webUrl)"
        return $response
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Type "ERROR"
        throw "Failed to upload file to SharePoint: $_"
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
        
        $compliantCount = ($ReportData.ComplianceSummary | Where-Object { $_.Status -eq "compliant" }).Count
        if (-not $compliantCount) { $compliantCount = 0 }
        
        $totalDevices = $ReportData.DevicesCount
        $complianceRate = if ($totalDevices -gt 0) { [math]::Round(($compliantCount / $totalDevices) * 100, 1) } else { 0 }
        
        $adaptiveCard = @{
            type = "message"
            attachments = @(
                @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    contentUrl = $null
                    content = @{
                        "$schema" = "http://adaptivecards.io/schemas/adaptive-card.json"
                        type = "AdaptiveCard"
                        version = "1.2"
                        msTeams = @{
                            width = "full"
                        }
                        body = @(
                            @{
                                type = "TextBlock"
                                size = "Large"
                                weight = "Bolder"
                                text = "Intune Device Compliance Report"
                                wrap = $true
                                color = "Default"
                            },
                            @{
                                type = "TextBlock"
                                spacing = "None"
                                text = "Report generated on $($ReportData.Timestamp)"
                                wrap = $true
                                isSubtle = $true
                            },
                            @{
                                type = "FactSet"
                                facts = @(
                                    @{
                                        title = "Report Name:"
                                        value = $ReportData.ReportName
                                    },
                                    @{
                                        title = "Total Devices:"
                                        value = $ReportData.DevicesCount.ToString()
                                    },
                                    @{
                                        title = "Compliance Rate:"
                                        value = "$complianceRate%"
                                    },
                                    @{
                                        title = "Execution Time:"
                                        value = "$executionTime minutes"
                                    }
                                )
                            }
                        )
                        actions = @(
                            @{
                                type = "Action.OpenUrl"
                                title = "View Report"
                                url = $ReportData.ReportUrl
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

try {
    $startTime = Get-Date
    Write-Log "=== Intune Device Compliance Report Generation Started ==="
    
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "ImportExcel module not found. Installing..." -Type "WARNING"
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }
    Import-Module ImportExcel
    
    if ($UseManagedIdentity) {
        $token = Get-MsGraphToken -UseManagedIdentity
    } else {
        $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    }
    
    $devices = Get-IntuneDeviceComplianceStatus -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds -BatchSize $BatchSize
    
    if ($devices.Count -eq 0) {
        Write-Log "No enrolled devices found in Intune" -Type "WARNING"
        return
    }
    
    $currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $reportName = "Intune_Device_Compliance_Report_$currentDate.xlsx"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $reportName)
    
    Export-DataToExcel -Data $devices -FilePath $tempPath
    
    $uploadResult = Upload-FileToSharePoint -Token $token -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FolderPath $FolderPath -FilePath $tempPath -FileName $reportName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    Remove-Item -Path $tempPath -Force
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    $complianceSummary = $devices | Group-Object -Property complianceState | 
                         Select-Object @{Name='Status';Expression={$_.Name}}, @{Name='Count';Expression={$_.Count}}
    
    Write-Log "=== Intune Device Compliance Report Generation Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Report URL: $($uploadResult.webUrl)"
    
    $result = [PSCustomObject]@{
        ReportName      = $reportName
        DevicesCount    = $devices.Count
        ReportUrl       = $uploadResult.webUrl
        ExecutionTimeMinutes = $duration.TotalMinutes
        Timestamp       = $currentDate
        ComplianceSummary = $complianceSummary
    }
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        $notificationSent = Send-TeamsNotification -WebhookUrl $TeamsWebhookUrl -ReportData $result
        if ($notificationSent) {
            $result | Add-Member -MemberType NoteProperty -Name "NotificationSent" -Value $true
        } else {
            $result | Add-Member -MemberType NoteProperty -Name "NotificationSent" -Value $false
        }
    }
    
    return $result
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    
    if ($tempPath -and (Test-Path -Path $tempPath)) {
        Remove-Item -Path $tempPath -Force
    }
    
    throw $_
}
finally {
    Write-Log "Script execution completed"
}