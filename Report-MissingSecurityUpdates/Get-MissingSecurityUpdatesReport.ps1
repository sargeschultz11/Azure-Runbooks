<#
.SYNOPSIS
    Generates a report of Windows devices missing multiple security updates from Log Analytics and uploads it to SharePoint.

.DESCRIPTION
    This Azure Runbook script connects to Log Analytics using a System-Assigned Managed Identity,
    retrieves all Windows devices with multiple missing security updates, exports the data
    to an Excel file, and uploads the file to a specified SharePoint document library.
    It also sends a notification to Microsoft Teams with a link to the report.

.PARAMETER WorkspaceId
    The Log Analytics Workspace ID to query for missing security updates data.

.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the report will be uploaded.
    
.PARAMETER SharePointDriveId
    The ID of the document library drive where the report will be uploaded.
    
.PARAMETER FolderPath
    Optional. The folder path within the document library where the report will be uploaded.
    If not specified, the file will be uploaded to the root of the document library.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER TeamsWebhookUrl
    Optional. Microsoft Teams webhook URL for sending notifications about the report.

.NOTES
    File Name: Get-MissingSecurityUpdatesReport.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-22
    
    Requires -Modules ImportExcel, Az.Accounts, Az.OperationalInsights
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointDriveId,
    
    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "",
    
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

function Get-LogAnalyticsToken {
    try {
        Write-Log "Acquiring Log Analytics token using Managed Identity..."
        Connect-AzAccount -Identity | Out-Null
        
        $token = Get-AzAccessToken -ResourceUrl "https://api.loganalytics.io"
        
        if ($token.Token -is [System.Security.SecureString]) {
            Write-Log "Token is SecureString, converting to plain text..."
            $tokenValue = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($token.Token)
            )
        } else {
            Write-Log "Token is plain string, no conversion needed."
            $tokenValue = $token.Token
        }
        
        if (-not [string]::IsNullOrEmpty($tokenValue)) {
            Write-Log "Log Analytics token acquired successfully."
            return $tokenValue
        } else {
            throw "Log Analytics token was empty."
        }
    }
    catch {
        Write-Log "Failed to acquire Log Analytics token: $_" -Type "ERROR"
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

function Query-LogAnalytics {
    param (
        [string]$WorkspaceId,
        [string]$Query,
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Querying Log Analytics Workspace: $WorkspaceId"
        Write-Log "Query: $Query"
        
        $apiVersion = "2022-09-01"
        $uri = "https://api.loganalytics.io/v1/workspaces/$WorkspaceId/query"
        
        $body = @{
            query = $Query
        }
        
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type" = "application/json"
        }
        
        $retryCount = 0
        $backoffSeconds = $InitialBackoffSeconds
        
        while ($true) {
            try {
                $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body ($body | ConvertTo-Json)
                
                if ($response.tables.Count -gt 0 -and $response.tables[0].rows.Count -gt 0) {
                    Write-Log "Retrieved $($response.tables[0].rows.Count) rows from Log Analytics"
                    
                    $columns = $response.tables[0].columns
                    $rows = $response.tables[0].rows
                    
                    $resultObjects = @()
                    
                    foreach ($row in $rows) {
                        $resultObject = [PSCustomObject]@{}
                        
                        for ($i = 0; $i -lt $columns.Count; $i++) {
                            $columnName = $columns[$i].name
                            $resultObject | Add-Member -MemberType NoteProperty -Name $columnName -Value $row[$i]
                        }
                        
                        $resultObjects += $resultObject
                    }
                    
                    return $resultObjects
                }
                else {
                    Write-Log "No data returned from Log Analytics query" -Type "WARNING"
                    return @()
                }
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
                        Write-Log "Request throttled by Log Analytics API (429). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                    }
                    else {
                        Write-Log "Server error (5xx). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                    }
                    
                    Start-Sleep -Seconds $retryAfter
                    
                    $retryCount++
                    $backoffSeconds = $backoffSeconds * 2
                }
                else {
                    Write-Log "Log Analytics query failed: $_" -Type "ERROR"
                    throw $_
                }
            }
        }
    }
    catch {
        Write-Log "Failed to query Log Analytics: $_" -Type "ERROR"
        throw "Failed to query Log Analytics: $_"
    }
}

function Export-DataToExcel {
    param (
        [array]$Data,
        [string]$FilePath
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
        }
        
        $excelParams = @{
            Path          = $FilePath
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            AutoFilter    = $true
            WorksheetName = "Missing Security Updates"
            TableName     = "MissingUpdatesTable"
            PassThru      = $true
        }
        
        $selectedProperties = @(
            @{Name = 'Device Name'; Expression = { $_.DeviceName }},
            @{Name = 'Azure AD Device ID'; Expression = { $_.AzureADDeviceId }},
            @{Name = 'Alert ID'; Expression = { $_.AlertId }},
            @{Name = 'Alert Generated'; Expression = { $_.TimeGenerated }},
            @{Name = 'Description'; Expression = { $_.Description }},
            @{Name = 'Recommendation'; Expression = { $_.Recommendation }}
        )
        
        $excel = $Data | Select-Object $selectedProperties | Export-Excel @excelParams
        
        $summarySheet = $excel.Workbook.Worksheets.Add("Summary")
        $summarySheet.Cells["A1"].Value = "Windows Devices Missing Security Updates Report"
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
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Alert Age Distribution"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $currentTime = Get-Date
        $ageGroups = @(
            @{Name = "Less than 24 hours"; Count = 0},
            @{Name = "1-3 days"; Count = 0},
            @{Name = "4-7 days"; Count = 0},
            @{Name = "8-14 days"; Count = 0},
            @{Name = "15-30 days"; Count = 0},
            @{Name = "More than 30 days"; Count = 0}
        )
        
        foreach ($device in $Data) {
            $alertTime = [datetime]$device.TimeGenerated
            $ageDays = ($currentTime - $alertTime).Days
            
            if ($ageDays -lt 1) {
                $ageGroups[0].Count++
            }
            elseif ($ageDays -ge 1 -and $ageDays -le 3) {
                $ageGroups[1].Count++
            }
            elseif ($ageDays -ge 4 -and $ageDays -le 7) {
                $ageGroups[2].Count++
            }
            elseif ($ageDays -ge 8 -and $ageDays -le 14) {
                $ageGroups[3].Count++
            }
            elseif ($ageDays -ge 15 -and $ageDays -le 30) {
                $ageGroups[4].Count++
            }
            else {
                $ageGroups[5].Count++
            }
        }
        
        $summarySheet.Cells["A$row"].Value = "Age"
        $summarySheet.Cells["B$row"].Value = "Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($ageGroup in $ageGroups) {
            $summarySheet.Cells["A$row"].Value = $ageGroup.Name
            $summarySheet.Cells["B$row"].Value = $ageGroup.Count
            $row++
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
        
        $timeDistribution = ""
        if ($ReportData.AgeDistribution) {
            foreach ($item in $ReportData.AgeDistribution) {
                $timeDistribution += "$($item.Name): $($item.Count)\n"
            }
        }
        
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
                                text   = "Windows Security Updates Alert"
                                wrap   = $true
                                color  = "Attention"
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
                                        title = "Devices Missing Multiple Updates:"
                                        value = $ReportData.DevicesCount.ToString()
                                    },
                                    @{
                                        title = "Execution Time:"
                                        value = "$executionTime minutes"
                                    }
                                )
                            },
                            @{
                                type   = "TextBlock"
                                text   = "These devices require attention to maintain security compliance. Please review the report for details."
                                wrap   = $true
                                weight = "Bolder"
                            }
                        )
                        actions   = @(
                            @{
                                type  = "Action.OpenUrl"
                                title = "View Full Report"
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

# Main script
try {
    $startTime = Get-Date
    Write-Log "=== Windows Missing Security Updates Report Generation Started ==="
    
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "ImportExcel module not found. Installing..." -Type "WARNING"
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
    
    Write-Log "Authenticating and retrieving tokens..."
    $graphToken = Get-MsGraphToken
    $logAnalyticsToken = Get-LogAnalyticsToken
    
    $query = @"
UCDeviceAlert
| where AlertSubtype == "MultipleSecurityUpdatesMissing"
| where AlertStatus == "Active"
| summarize arg_max(TimeGenerated, *) by DeviceName
| project DeviceName, AzureADDeviceId, AlertId, TimeGenerated, Description, Recommendation
| order by TimeGenerated desc
"@
    
    Write-Log "Executing Log Analytics query..."
    $missingUpdatesData = Query-LogAnalytics -WorkspaceId $WorkspaceId -Query $query -Token $logAnalyticsToken -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    if ($missingUpdatesData.Count -eq 0) {
        Write-Log "No devices found with missing security updates" -Type "WARNING"
        
        $result = [PSCustomObject]@{
            DevicesCount         = 0
            ReportName           = "No devices found"
            ExecutionTimeMinutes = (Get-Date - $startTime).TotalMinutes
            Timestamp            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ReportUrl            = "N/A"
            NotificationSent     = $false
            AgeDistribution      = $null
        }
        
        return $result
    }
    
    Write-Log "Found $($missingUpdatesData.Count) devices with missing security updates"
    
    $currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $reportName = "Windows_Missing_Security_Updates_Report_$currentDate.xlsx"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $reportName)
    
    Export-DataToExcel -Data $missingUpdatesData -FilePath $tempPath
    
    $uploadResult = Upload-FileToSharePoint -Token $graphToken -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FolderPath $FolderPath -FilePath $tempPath -FileName $reportName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    if (Test-Path -Path $tempPath) {
        Remove-Item -Path $tempPath -Force
        Write-Log "Temporary file removed: $tempPath"
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Windows Missing Security Updates Report Generation Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Report URL: $($uploadResult.webUrl)"
    
    $currentTime = Get-Date
    $ageGroups = @(
        @{Name = "Less than 24 hours"; Count = 0},
        @{Name = "1-3 days"; Count = 0},
        @{Name = "4-7 days"; Count = 0},
        @{Name = "8-14 days"; Count = 0},
        @{Name = "15-30 days"; Count = 0},
        @{Name = "More than 30 days"; Count = 0}
    )
    
    foreach ($device in $missingUpdatesData) {
        $alertTime = [datetime]$device.TimeGenerated
        $ageDays = ($currentTime - $alertTime).Days
        
        if ($ageDays -lt 1) {
            $ageGroups[0].Count++
        }
        elseif ($ageDays -ge 1 -and $ageDays -le 3) {
            $ageGroups[1].Count++
        }
        elseif ($ageDays -ge 4 -and $ageDays -le 7) {
            $ageGroups[2].Count++
        }
        elseif ($ageDays -ge 8 -and $ageDays -le 14) {
            $ageGroups[3].Count++
        }
        elseif ($ageDays -ge 15 -and $ageDays -le 30) {
            $ageGroups[4].Count++
        }
        else {
            $ageGroups[5].Count++
        }
    }
    
    $result = [PSCustomObject]@{
        DevicesCount         = $missingUpdatesData.Count
        ReportName           = $reportName
        ExecutionTimeMinutes = $duration.TotalMinutes
        Timestamp            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ReportUrl            = $uploadResult.webUrl
        NotificationSent     = $false
        AgeDistribution      = $ageGroups
    }
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        $notificationSent = Send-TeamsNotification -WebhookUrl $TeamsWebhookUrl -ReportData $result
        if ($notificationSent) {
            $result.NotificationSent = $true
            Write-Log "Teams notification sent successfully"
        } else {
            Write-Log "Failed to send Teams notification" -Type "WARNING"
        }
    } else {
        Write-Log "Teams webhook URL not provided. Skipping notification."
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