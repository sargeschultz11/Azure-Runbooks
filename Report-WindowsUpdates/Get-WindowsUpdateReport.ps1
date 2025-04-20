# Requires -Modules "ImportExcel", "Az.Accounts", "Az.OperationalInsights"
<#
.SYNOPSIS
    Generates a Windows Update compliance dashboard using Log Analytics data.

.DESCRIPTION
    This Azure Automation runbook connects to a Log Analytics workspace using a System-Assigned Managed Identity,
    queries Windows Update activity, and maintains a persistent Excel dashboard file in SharePoint.
    The dashboard tracks update compliance over time, extending beyond the Log Analytics retention period.

.PARAMETER MonthsToQuery
    The number of months to include in the query (counting backwards from current month). Default is 1.
    Note: This is limited by your Log Analytics workspace retention period (typically 30 days).

.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the dashboard will be stored.

.PARAMETER SharePointDriveId
    The ID of the document library drive where the dashboard will be stored.

.PARAMETER FolderPath
    Optional. The folder path within the document library where the dashboard will be stored.
    If not specified, the file will be stored in the root of the document library.

.PARAMETER TeamsWebhookUrl
    Optional. Microsoft Teams webhook URL for sending notifications about dashboard updates.

.PARAMETER WorkspaceId
    The ID of the Log Analytics workspace to query for update data.

.NOTES
    File Name: Get-WindowsUpdateReport.ps1
    Author: Ryan Schultz (Enhanced by Claude)
    Version: 4.0
    Updated: 2025-04-19

    Requires -Modules ImportExcel, Az.Accounts, Az.OperationalInsights
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$MonthsToQuery = 1,

    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteId,

    [Parameter(Mandatory = $true)]
    [string]$SharePointDriveId,

    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "",

    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl = "",

    [Parameter(Mandatory = $true)]
    [string]$WorkspaceId
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
        Write-Log "Acquiring Microsoft Graph API token..."
        # Handle the breaking change warning by using the current approach but planning for the future
        $graphToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" 
        
        if ($null -ne $graphToken -and $null -ne $graphToken.Token) {
            Write-Log "Successfully acquired Microsoft Graph token"
            # Convert to string to handle the token properly
            $tokenStr = $graphToken.Token
            return $tokenStr
        }
        else {
            throw "Failed to acquire valid token - token is null or empty"
        }
    }
    catch {
        Write-Log "Failed to acquire Microsoft Graph token: $_" -Type "ERROR"
        throw "Authentication failed: $_"
    }
}

function Invoke-MsGraphRequest {
    param (
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json"
    )
    
    $maxRetries = 3
    $retryCount = 0
    $backoffSeconds = 2
    
    $params = @{
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        Method      = $Method
    }
    
    if ($ContentType) {
        $params.Add("ContentType", $ContentType)
    }
    
    if ($null -ne $Body -and $Method -ne "GET") {
        if ($Body -is [byte[]]) {
            $params.Add("Body", $Body)
        }
        else {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
    }
    
    while ($retryCount -le $maxRetries) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            
            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $retryCount -lt $maxRetries) {
                $retryCount++
                Write-Log "Request failed with status code $statusCode. Retrying in $backoffSeconds seconds..." -Type "WARNING"
                Start-Sleep -Seconds $backoffSeconds
                $backoffSeconds *= 2
            }
            else {
                Write-Log "Request failed with status code $statusCode`: $_" -Type "ERROR"
                throw $_
            }
        }
    }
}

function Get-SharePointFile {
    param (
        [string]$Token,
        [string]$SiteId,
        [string]$DriveId,
        [string]$FilePath,
        [string]$OutputPath
    )
    
    try {
        $fileUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$FilePath"
        Write-Log "Checking if file exists in SharePoint: $FilePath"
        
        $fileInfo = Invoke-MsGraphRequest -Token $Token -Uri $fileUri
        
        if ($fileInfo) {
            Write-Log "File found. Downloading..."
            $downloadUrl = $fileInfo.'@microsoft.graph.downloadUrl'
            
            Invoke-WebRequest -Uri $downloadUrl -OutFile $OutputPath
            Write-Log "File downloaded successfully to: $OutputPath"
            return $true
        }
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            Write-Log "File not found in SharePoint. Will create new file." -Type "INFO"
        }
        else {
            Write-Log "Error checking/downloading file: $_" -Type "WARNING"
        }
        return $false
    }
}

function Upload-SharePointFile {
    param (
        [string]$Token,
        [string]$SiteId,
        [string]$DriveId,
        [string]$FilePath,
        [string]$SourcePath
    )
    
    try {
        $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$FilePath"+":/content"
        Write-Log "Uploading file to SharePoint..."
        
        $fileBytes = [System.IO.File]::ReadAllBytes($SourcePath)
        
        $response = Invoke-MsGraphRequest -Token $Token -Uri $uploadUri -Method PUT -Body $fileBytes -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        Write-Log "File uploaded successfully to SharePoint"
        return $response
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Type "ERROR"
        throw $_
    }
}

function Get-UpdateDataFromLogAnalytics {
    param (
        [string]$WorkspaceId,
        [int]$MonthsToQuery
    )
    
    try {
        Write-Log "Querying Log Analytics for Windows Update data..."
        $startDate = (Get-Date).AddMonths(-$MonthsToQuery)
        $startDateStr = $startDate.ToString("yyyy-MM-dd")
        
        # List of possible Windows Update tables to check
        $tables = @("UCClientUpdateStatus", "Update", "WindowsUpdates", "Update_CL")
        $validTable = $null
        
        # Find the first valid table that exists
        foreach ($table in $tables) {
            try {
                Write-Log "Testing table: $table..."
                $testQuery = "$table | take 1"
                $testResult = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $testQuery -ErrorAction Stop
                
                if ($testResult.Results.Count -gt 0) {
                    Write-Log "Found valid Windows Update table: $table"
                    $validTable = $table
                    break
                }
            }
            catch {
                Write-Log "Table $table not found or not accessible" -Type "WARNING"
            }
        }
        
        if (-not $validTable) {
            throw "No valid Windows Update table found in workspace"
        }
        
        # Build a query to get update status by device and month
        $kqlQuery = @"
$validTable
| where TimeGenerated >= datetime('$startDateStr')
| extend Month = format_datetime(TimeGenerated, 'yyyy-MM')
| summarize UpdateCount = count(), LastSeen = max(TimeGenerated) by DeviceName, Month
"@
        
        $updateData = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $kqlQuery -ErrorAction Stop
        
        if (-not $updateData.Results -or $updateData.Results.Count -eq 0) {
            Write-Log "No update data found in table $validTable" -Type "WARNING"
            throw "No update data found"
        }
        
        Write-Log "Successfully retrieved update data with $($updateData.Results.Count) records"
        return $updateData.Results
    }
    catch {
        Write-Log "Failed to retrieve update data: $_" -Type "ERROR"
        throw $_
    }
}

function Format-MonthDisplay {
    param([string]$MonthStr)
    
    try {
        $date = [datetime]::ParseExact($MonthStr, "yyyy-MM", [System.Globalization.CultureInfo]::InvariantCulture)
        return $date.ToString("MMM yyyy")
    }
    catch {
        return $MonthStr
    }
}

function Create-ComplianceReport {
    param (
        [array]$UpdateData,
        [array]$ExistingHistoricalData,
        [string]$OutputPath
    )
    
    try {
        # Process data to build device-month matrix
        $matrix = @{}
        $months = @()
        
        foreach ($row in $UpdateData) {
            $device = $row.DeviceName
            $month = $row.Month
            
            if (-not $months.Contains($month)) { 
                $months += $month 
            }
            
            if (-not $matrix.ContainsKey($device)) { 
                $matrix[$device] = @{} 
            }
            
            $matrix[$device][$month] = @{
                UpdateCount = $row.UpdateCount
                LastSeen = $row.LastSeen
            }
        }
        
        # Sort months chronologically
        $months = @($months | Sort-Object)
        
        # Create month display mapping (yyyy-MM to MMM yyyy)
        $monthDisplay = @{}
        foreach ($month in $months) {
            $monthDisplay[$month] = Format-MonthDisplay $month
        }
        
        # Build report for device update status
        $deviceData = @()
        foreach ($device in $matrix.Keys | Sort-Object) {
            $row = [ordered]@{ "Device Name" = $device }
            
            foreach ($month in $months) {
                $display = $monthDisplay[$month]
                $hasUpdates = $matrix[$device].ContainsKey($month)
                
                # Use "Updated" instead of "Yes/No" for clarity
                $row[$display] = if ($hasUpdates) { "Updated" } else { "Not Updated" }
                
                # Add update count column for each month
                $row["$display (Count)"] = if ($hasUpdates) { $matrix[$device][$month].UpdateCount } else { 0 }
            }
            
            $deviceData += [PSCustomObject]$row
        }
        
        # Calculate summary data
        $summaryData = @()
        foreach ($month in $months) {
            $display = $monthDisplay[$month]
            $total = $matrix.Keys.Count
            $updated = ($matrix.Keys | Where-Object { $matrix[$_].ContainsKey($month) }).Count
            $notUpdated = $total - $updated
            $percent = if ($total -gt 0) { [math]::Round(($updated / $total) * 100, 2) } else { 0 }
            
            $summaryData += [PSCustomObject]@{
                Month = $display
                TotalDevices = $total
                UpdatedDevices = $updated
                NotUpdatedDevices = $notUpdated
                CompliancePercent = "$percent%"
                ComplianceValue = $percent  # Numeric value for charting
            }
        }
        
        # Merge with historical data if available
        $currentMonth = (Get-Date).ToString("MMM yyyy")
        $mergedHistoricalData = @()
        
        if ($ExistingHistoricalData -and ($ExistingHistoricalData | Measure-Object).Count -gt 0) {
            Write-Log "Merging with $(($ExistingHistoricalData | Measure-Object).Count) historical records..."
            
            # Remove current month from historical data to avoid duplicates
            $filteredHistorical = @($ExistingHistoricalData | Where-Object { $_.Month -ne $currentMonth })
            
            # Combine historical and new data
            if ($filteredHistorical.Count -gt 0) {
                $mergedHistoricalData = @($filteredHistorical)
                foreach ($item in $summaryData) {
                    $mergedHistoricalData += $item
                }
            } else {
                $mergedHistoricalData = @($summaryData)
            }
            
            # Remove any duplicates by month
            $mergedHistoricalData = @($mergedHistoricalData | Sort-Object -Property Month -Unique)
        }
        else {
            $mergedHistoricalData = @($summaryData)
        }
        
        # Sort historical data chronologically
        $mergedHistoricalData = @($mergedHistoricalData | Sort-Object -Property {
            try {
                [datetime]::ParseExact(($_.Month -replace '^(.{3}) (\d{4})$', '$1 01, $2'), "MMM dd, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            }
            catch {
                [datetime]::MinValue
            }
        })
        
        # Create Excel file - check if it already exists first
        $excel = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $OutputPath
        
        # Check if worksheets already exist and remove them to avoid the 'already exists' error
        try {
            if ($excel.Workbook.Worksheets["Dashboard"]) {
                $excel.Workbook.Worksheets.Delete("Dashboard")
                Write-Log "Removed existing Dashboard worksheet"
            }
            if ($excel.Workbook.Worksheets["Device Status"]) {
                $excel.Workbook.Worksheets.Delete("Device Status")
                Write-Log "Removed existing Device Status worksheet"
            }
            if ($excel.Workbook.Worksheets["Historical Data"]) {
                $excel.Workbook.Worksheets.Delete("Historical Data")
                Write-Log "Removed existing Historical Data worksheet"
            }
            if ($excel.Workbook.Worksheets["Compliance Trend"]) {
                $excel.Workbook.Worksheets.Delete("Compliance Trend")
                Write-Log "Removed existing Compliance Trend worksheet"
            }
        }
        catch {
            Write-Log "Warning when clearing existing worksheets: $_" -Type "WARNING"
        }
        
        # Create dashboard sheet
        $dashboardSheet = $excel.Workbook.Worksheets.Add("Dashboard")
        $dashboardSheet.Cells["A1"].Value = "Windows Update Compliance Dashboard"
        $dashboardSheet.Cells["A1"].Style.Font.Size = 16
        $dashboardSheet.Cells["A1"].Style.Font.Bold = $true
        
        $dashboardSheet.Cells["A3"].Value = "Last Updated:"
        $dashboardSheet.Cells["B3"].Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        
        $dashboardSheet.Cells["A5"].Value = "Current Month Compliance:"
        $latestData = $mergedHistoricalData | Select-Object -Last 1
        $dashboardSheet.Cells["B5"].Value = $latestData.CompliancePercent
        $dashboardSheet.Cells["B5"].Style.Font.Bold = $true
        
        $dashboardSheet.Cells["A7"].Value = "Summary Statistics"
        $dashboardSheet.Cells["A7"].Style.Font.Bold = $true
        
        $dashboardSheet.Cells["A9"].Value = "Total Devices:"
        $dashboardSheet.Cells["B9"].Value = $latestData.TotalDevices
        
        $dashboardSheet.Cells["A10"].Value = "Updated Devices:"
        $dashboardSheet.Cells["B10"].Value = $latestData.UpdatedDevices
        
        $dashboardSheet.Cells["A11"].Value = "Not Updated Devices:"
        $dashboardSheet.Cells["B11"].Value = $latestData.NotUpdatedDevices
        
        # Add compliance history table to dashboard
        $dashboardSheet.Cells["A14"].Value = "Monthly Compliance History"
        $dashboardSheet.Cells["A14"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A14"].Style.Font.Size = 12
        
        # Add headers for compliance table
        $dashboardSheet.Cells["A16"].Value = "Month"
        $dashboardSheet.Cells["B16"].Value = "Compliance %"
        $dashboardSheet.Cells["C16"].Value = "Updated Devices"
        $dashboardSheet.Cells["D16"].Value = "Total Devices"
        
        $dashboardSheet.Cells["A16:D16"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A16:D16"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $dashboardSheet.Cells["A16:D16"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
        
        # Add compliance history data (last 6 months)
        $rowIndex = 17
        foreach ($record in ($mergedHistoricalData | Select-Object -Last 6)) {
            $dashboardSheet.Cells[$rowIndex, 1].Value = $record.Month
            $dashboardSheet.Cells[$rowIndex, 2].Value = $record.CompliancePercent
            $dashboardSheet.Cells[$rowIndex, 3].Value = $record.UpdatedDevices
            $dashboardSheet.Cells[$rowIndex, 4].Value = $record.TotalDevices
            
            # Add conditional formatting - green for high compliance, yellow for medium, red for low
            $complianceValue = [double]($record.CompliancePercent -replace '%', '')
            if ($complianceValue -ge 95) {
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))
            }
            elseif ($complianceValue -ge 80) {
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 156))
            }
            else {
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $dashboardSheet.Cells[$rowIndex, 2].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206))
            }
            
            $rowIndex++
        }
        
        # Auto-size columns
        for ($i = 1; $i -le 4; $i++) {
            $dashboardSheet.Column($i).AutoFit()
        }
        
        # Add mini chart to dashboard
        if ($mergedHistoricalData.Count -gt 1) {
            # Add chart data
            $chartDataStartRow = 27
            $dashboardSheet.Cells[$chartDataStartRow, 1].Value = "Month"
            $dashboardSheet.Cells[$chartDataStartRow, 2].Value = "Compliance"
            $rowIndex = $chartDataStartRow + 1
            foreach ($record in $mergedHistoricalData) {
                $dashboardSheet.Cells[$rowIndex, 1].Value = $record.Month
                $dashboardSheet.Cells[$rowIndex, 2].Value = $record.ComplianceValue
                $rowIndex++
            }
            # Create mini chart
            $miniChart = $dashboardSheet.Drawings.AddChart("DashboardComplianceChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Line)
            $miniChart.Title.Text = "Compliance Trend"
            $miniChart.SetPosition(13, 0, 5, 0)
            $miniChart.SetSize(450, 250)
            $series = $miniChart.Series.Add("B" + ($chartDataStartRow + 1) + ":B" + ($rowIndex - 1), 
                                            "A" + ($chartDataStartRow + 1) + ":A" + ($rowIndex - 1))
            $series.Header = "Compliance %"
            
            # Format Y axis to show percentages
            $miniChart.YAxis.Format = "0\%"
            $miniChart.YAxis.MaxValue = 100
            $miniChart.YAxis.MinValue = 0
        }
        
        # Create device status sheet
        $deviceSheet = $excel.Workbook.Worksheets.Add("Device Status")
        
        # Write headers
        $colIndex = 1
        foreach ($key in ($deviceData[0].PSObject.Properties.Name)) {
            $deviceSheet.Cells[1, $colIndex].Value = $key
            $deviceSheet.Cells[1, $colIndex].Style.Font.Bold = $true
            $colIndex++
        }
        
        # Write data rows
        $rowIndex = 2
        foreach ($device in $deviceData) {
            $colIndex = 1
            foreach ($key in ($device.PSObject.Properties.Name)) {
                $deviceSheet.Cells[$rowIndex, $colIndex].Value = $device.$key
                
                # Apply conditional formatting (green for Updated, red for Not Updated)
                if ($device.$key -eq "Updated") {
                    $deviceSheet.Cells[$rowIndex, $colIndex].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $deviceSheet.Cells[$rowIndex, $colIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))
                }
                elseif ($device.$key -eq "Not Updated") {
                    $deviceSheet.Cells[$rowIndex, $colIndex].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $deviceSheet.Cells[$rowIndex, $colIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206))
                }
                
                $colIndex++
            }
            $rowIndex++
        }
        
        # Auto-size columns
        for ($i = 1; $i -le $colIndex - 1; $i++) {
            $deviceSheet.Column($i).AutoFit()
        }
        
        # Create historical data sheet
        $historySheet = $excel.Workbook.Worksheets.Add("Historical Data")
        
        # Write headers
        $colIndex = 1
        foreach ($key in ($mergedHistoricalData[0].PSObject.Properties.Name)) {
            # Skip ComplianceValue as it's just for charting
            if ($key -ne "ComplianceValue") {
                $historySheet.Cells[1, $colIndex].Value = $key
                $historySheet.Cells[1, $colIndex].Style.Font.Bold = $true
                $colIndex++
            }
        }
        
        # Write data rows
        $rowIndex = 2
        foreach ($record in $mergedHistoricalData) {
            $colIndex = 1
            foreach ($key in ($record.PSObject.Properties.Name)) {
                # Skip ComplianceValue as it's just for charting
                if ($key -ne "ComplianceValue") {
                    $historySheet.Cells[$rowIndex, $colIndex].Value = $record.$key
                    $colIndex++
                }
            }
            $rowIndex++
        }
        
        # Auto-size columns
        for ($i = 1; $i -le $colIndex - 1; $i++) {
            $historySheet.Column($i).AutoFit()
        }
        
        # Create compliance chart sheet
        if ($mergedHistoricalData.Count -gt 1) {
            $chartSheet = $excel.Workbook.Worksheets.Add("Compliance Trend")
            
            # Add month labels
            $chartSheet.Cells["A1"].Value = "Month"
            $chartSheet.Cells["A1"].Style.Font.Bold = $true
            
            $chartSheet.Cells["B1"].Value = "Compliance %"
            $chartSheet.Cells["B1"].Style.Font.Bold = $true
            
            $rowIndex = 2
            foreach ($record in $mergedHistoricalData) {
                $chartSheet.Cells[$rowIndex, 1].Value = $record.Month
                $chartSheet.Cells[$rowIndex, 2].Value = $record.ComplianceValue
                $rowIndex++
            }
            
            # Create chart
            $chart = $chartSheet.Drawings.AddChart("ComplianceChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Line)
            $chart.Title.Text = "Monthly Update Compliance Trend"
            $chart.SetPosition(4, 0, 4, 0)
            $chart.SetSize(800, 400)
            
            $series = $chart.Series.Add("B2:B$($rowIndex-1)", "A2:A$($rowIndex-1)")
            $series.Header = "Compliance %"
            
            # Format Y axis to show percentages
            $chart.YAxis.Format = "0\%"
            $chart.YAxis.MaxValue = 100
            $chart.YAxis.MinValue = 0
        }
        
        # Save and close the Excel package
        $excel.Save()
        $excel.Dispose()
        
        Write-Log "Excel report created successfully at: $OutputPath"
        
        return @{
            DeviceData = $deviceData
            SummaryData = $latestData
            HistoricalData = $mergedHistoricalData
        }
    }
    catch {
        Write-Log "Failed to create compliance report: $_" -Type "ERROR"
        
        if ($excel) {
            $excel.Dispose()
        }
        
        throw $_
    }
}

function Send-TeamsNotification {
    param (
        [string]$WebhookUrl,
        [PSCustomObject]$SummaryData,
        [string]$DashboardUrl
    )
    
    try {
        Write-Log "Sending notification to Microsoft Teams..."
        
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
                        body = @(
                            @{
                                type = "TextBlock"
                                size = "Large"
                                weight = "Bolder"
                                text = "Windows Update Compliance Dashboard"
                                wrap = $true
                            },
                            @{
                                type = "TextBlock"
                                text = "Dashboard updated on $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                                wrap = $true
                                isSubtle = $true
                            },
                            @{
                                type = "FactSet"
                                facts = @(
                                    @{
                                        title = "Current Compliance:"
                                        value = $SummaryData.CompliancePercent
                                    },
                                    @{
                                        title = "Total Devices:"
                                        value = "$($SummaryData.TotalDevices)"
                                    },
                                    @{
                                        title = "Updated Devices:"
                                        value = "$($SummaryData.UpdatedDevices)"
                                    },
                                    @{
                                        title = "Not Updated:"
                                        value = "$($SummaryData.NotUpdatedDevices)"
                                    }
                                )
                            }
                        )
                        actions = @(
                            @{
                                type = "Action.OpenUrl"
                                title = "View Dashboard"
                                url = $DashboardUrl
                            }
                        )
                    }
                }
            )
        }
        
        $notificationBody = ConvertTo-Json -InputObject $adaptiveCard -Depth 10
        
        Invoke-RestMethod -Uri $WebhookUrl -Method POST -Body $notificationBody -ContentType 'application/json'
        
        Write-Log "Teams notification sent successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send Teams notification: $_" -Type "WARNING"
        return $false
    }
}

# Main script execution
try {
    Write-Log "=== Windows Update Dashboard Generation Started ==="
    Write-Log "Parameters: MonthsToQuery=$MonthsToQuery, WorkspaceId=$WorkspaceId"
    
    $dashboardFileName = "Windows_Update_Dashboard.xlsx"
    $dashboardTempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $dashboardFileName)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    if (Test-Path -Path $dashboardTempPath) {
        Remove-Item -Path $dashboardTempPath -Force
    }
    
    Import-Module -Name ImportExcel -ErrorAction Stop
    
    Write-Log "Connecting to Azure with Managed Identity..."
    Connect-AzAccount -Identity | Out-Null
    
    $token = Get-MsGraphToken
    
    $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
        $dashboardFileName
    } else {
        "$FolderPath/$dashboardFileName"
    }
    
    $existingHistoricalData = @()
    $fileExists = Get-SharePointFile -Token $token -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FilePath $uploadPath -OutputPath $dashboardTempPath
    
    if ($fileExists) {
        try {
            $existingHistoricalData = Import-Excel -Path $dashboardTempPath -WorksheetName "Historical Data" -ErrorAction Stop
            Write-Log "Successfully imported historical data with $($existingHistoricalData.Count) records"
        }
        catch {
            Write-Log "Error reading historical data: $_" -Type "WARNING"
            Write-Log "Will create new historical data"
        }
    }
    
    $updateData = Get-UpdateDataFromLogAnalytics -WorkspaceId $WorkspaceId -MonthsToQuery $MonthsToQuery
    
    $reportData = Create-ComplianceReport -UpdateData $updateData -ExistingHistoricalData $existingHistoricalData -OutputPath $dashboardTempPath
    
    $uploadResult = Upload-SharePointFile -Token $token -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FilePath $uploadPath -SourcePath $dashboardTempPath
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        $notificationSent = Send-TeamsNotification -WebhookUrl $TeamsWebhookUrl -SummaryData $reportData.SummaryData -DashboardUrl $uploadResult.webUrl
    }
    
    if (Test-Path -Path $dashboardTempPath) {
        Remove-Item -Path $dashboardTempPath -Force
    }
    
    Write-Log "=== Windows Update Dashboard Generation Completed ==="
    
    return @{
        DashboardName = $dashboardFileName
        DashboardUrl = $uploadResult.webUrl
        DeviceCount = $reportData.SummaryData.TotalDevices
        CompliancePercent = $reportData.SummaryData.CompliancePercent
        LastUpdated = $timestamp
    }
}
catch {
    $errorMessage = $_.Exception.Message
    $errorType = $_.Exception.GetType().FullName
    $errorLine = $_.InvocationInfo.ScriptLineNumber
    Write-Log "Script execution failed at line $errorLine ($errorType): $errorMessage" -Type "ERROR"
    
    if ($dashboardTempPath -and (Test-Path -Path $dashboardTempPath)) {
        try {
            Remove-Item -Path $dashboardTempPath -Force -ErrorAction SilentlyContinue
            Write-Log "Cleaned up temp file: $dashboardTempPath"
        }
        catch {
            Write-Log "Failed to clean up temp file: $_" -Type "WARNING"
        }
    }
    
    throw $_
}