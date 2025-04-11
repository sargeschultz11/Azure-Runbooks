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
    File Name: Get-WindowsUpdateDashboard.ps1
    Author: Ryan Schultz
    Version: 3.0
    Updated: 2025-04-11

    Requires -Modules ImportExcel, Az.Accounts
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
        }
        "WARNING" { 
            Write-Warning $Message 
        }
        default { 
            Write-Output $LogMessage
        }
    }
}

$dashboardFileName = "Windows_Update_Dashboard.xlsx"
$dashboardFilePath = $null
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Main script execution
try {
    Write-Log "=== Windows Update Dashboard Generation Started ==="
    Write-Log "MonthsToQuery parameter value: $MonthsToQuery"
    
    $dashboardFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $dashboardFileName)
    
    Import-Module ImportExcel -ErrorAction Stop
    Connect-AzAccount -Identity | Out-Null
    $graphToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $token = $graphToken.Token
    $startDate = (Get-Date).AddMonths(-$MonthsToQuery)
    $startDateStr = $startDate.ToString("yyyy-MM-dd")
    $tables = @("UCClientUpdateStatus", "Update", "WindowsUpdates", "Update_CL")
    $validTable = $null
    $updateData = $null
    
    foreach ($table in $tables) {
        try {
            Write-Log "Trying to query table: $table..."
            
            $testQuery = "$table | take 1"
            $testResult = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $testQuery -ErrorAction Stop
            
            if ($testResult.Results.Count -gt 0) {
                Write-Log "Found valid table: $table"
                $validTable = $table
                break
            }
        }
        catch {
            Write-Log "Table $table not found or not accessible: $_" -Type "WARNING"
        }
    }
    
    if (-not $validTable) {
        Write-Log "No valid Windows Update table found in the workspace. Please verify your Log Analytics configuration." -Type "ERROR"
        throw "No valid Windows Update table found"
    }
    
    $kqlQuery = @"
$validTable
| where TimeGenerated >= datetime($startDateStr)
| extend Month = format_datetime(TimeGenerated, 'yyyy-MM')
| summarize UpdateCount = count(), LastSeen = max(TimeGenerated) by DeviceName, Month
"@
    
    try {
        $updateData = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $kqlQuery -ErrorAction Stop
        
        if (-not $updateData.Results -or $updateData.Results.Count -eq 0) {
            Write-Log "No update data found in table $validTable" -Type "WARNING"
            throw "No update data found"
        }
        
        Write-Log "Successfully retrieved update data from $validTable table"
    }
    catch {
        Write-Log "Failed to query update data: $_" -Type "ERROR"
        throw $_
    }
    
    $matrix = @{}
    $months = @()
    foreach ($row in $updateData.Results) {
        $device = $row.DeviceName
        $month = $row.Month
        
        if (-not $months.Contains($month)) { 
            $months += $month 
        }
        
        if (-not $matrix.ContainsKey($device)) { 
            $matrix[$device] = @{} 
        }
        
        $matrix[$device][$month] = @{
            HasUpdates = $true
            UpdateCount = $row.UpdateCount
            LastSeen = $row.LastSeen
        }
    }
    $months = $months | Sort-Object
    $monthDisplay = @{}
    foreach ($month in $months) {
        try {
            $date = [datetime]::ParseExact($month, "yyyy-MM", [System.Globalization.CultureInfo]::InvariantCulture)
            $monthDisplay[$month] = $date.ToString("MMM yyyy")
        }
        catch {
            $monthDisplay[$month] = $month
        }
    }
    
    Write-Log "Months found in query results: $($months -join ', ')"
    
    $reportData = @()
    foreach ($device in $matrix.Keys) {
        $deviceData = [ordered]@{ "Device Name" = $device }
        
        foreach ($month in $months) {
            $display = $monthDisplay[$month]
            $deviceData[$display] = if ($matrix[$device].ContainsKey($month)) { "Yes" } else { "No" }
        }
        
        $reportData += [PSCustomObject]$deviceData
    }
    
    $summaryData = @()
    foreach ($month in $months) {
        $display = $monthDisplay[$month]
        $updated = ($matrix.Keys | Where-Object { $matrix[$_].ContainsKey($month) }).Count
        $total = $matrix.Keys.Count
        $notUpdated = $total - $updated
        $percent = if ($total -gt 0) { [math]::Round(($updated / $total) * 100, 2) } else { 0 }
        $summaryData += [PSCustomObject]@{
            Month = $display
            TotalDevices = $total
            UpdatedDevices = $updated
            NotUpdatedDevices = $notUpdated
            CompliancePercent = "$percent%"
        }
    }
    
    $currentMonth = (Get-Date).ToString("yyyy-MM")
    $currentMonthDisplay = (Get-Date).ToString("MMM yyyy")
    $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
        $dashboardFileName
    } else {
        "$FolderPath/$dashboardFileName"
    }
    
    $historicalData = @()
    
    $fileExists = $false
    try {
        $downloadUri = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/drives/$SharePointDriveId/root:/$uploadPath"
        Write-Log "Checking if dashboard file exists in SharePoint..."
        $checkResponse = Invoke-RestMethod -Uri $downloadUri -Headers @{ Authorization = "Bearer $token" } -Method GET -ErrorAction Stop
        
        if ($checkResponse) {
            Write-Log "Existing dashboard found. Downloading..."
            $downloadUrl = $checkResponse.'@microsoft.graph.downloadUrl'
            Invoke-WebRequest -Uri $downloadUrl -OutFile $dashboardFilePath
            Write-Log "Downloaded existing dashboard file"
            $fileExists = $true
            
            try {
                Write-Log "Attempting to read historical data from existing file..."
                $importedData = Import-Excel -Path $dashboardFilePath -WorksheetName "Historical Data" -ErrorAction Stop
                
                if ($importedData) {
                    Write-Log "Successfully read historical data. Found $($importedData.Count) records."
                    $historicalData = $importedData
                }
            }
            catch {
                Write-Log "Error reading historical data: $_" -Type "WARNING"
                Write-Log "Will create new historical data"
            }
        }
    }
    catch {
        Write-Log "Dashboard file doesn't exist yet, will create new one" -Type "INFO"
        $fileExists = $false
    }
    if ($historicalData.Count -gt 0) {
        $updatedHistoricalData = $historicalData | Where-Object { $_.Month -ne $currentMonthDisplay }
        $updatedHistoricalData += $summaryData
        $historicalData = $updatedHistoricalData | Sort-Object -Property Month -Unique
    }
    else {
        $historicalData = $summaryData
    }
    
    $historicalData = $historicalData | Sort-Object -Property { 
        try {
            [datetime]::ParseExact(($_.Month -replace '^(.{3}) (\d{4})$', '$1 01, $2'), "MMM dd, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        catch {
            [datetime]::MinValue
        }
    }
    
    $latestData = $historicalData | Select-Object -Last 1
    $reportData | Export-Excel -Path $dashboardFilePath -WorksheetName "Current Month" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    $historicalData | Export-Excel -Path $dashboardFilePath -WorksheetName "Historical Data" -AutoSize -FreezeTopRow -BoldTopRow
    $dashboardHtml = @"
<table>
    <tr>
        <td colspan="2" style="font-size: 16pt; font-weight: bold;">Windows Update Compliance Dashboard</td>
    </tr>
    <tr>
        <td colspan="2">Last Updated: $timestamp</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
        <td style="font-weight: bold;">Current Month Compliance:</td>
        <td>$($latestData.CompliancePercent)</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
        <td style="font-weight: bold;">Summary Statistics</td>
        <td></td>
    </tr>
    <tr>
        <td>Total Devices:</td>
        <td>$($latestData.TotalDevices)</td>
    </tr>
    <tr>
        <td>Updated Devices:</td>
        <td>$($latestData.UpdatedDevices)</td>
    </tr>
    <tr>
        <td>Not Updated Devices:</td>
        <td>$($latestData.NotUpdatedDevices)</td>
    </tr>
</table>
"@
    
    $dashboardCsv = $dashboardHtml -replace '<tr>', '' -replace '</tr>', "`n" -replace '<td[^>]*>', '' -replace '</td>', '|' -replace '<[^>]+>', ''
    $dashboardCsvPath = [System.IO.Path]::ChangeExtension($dashboardFilePath, "csv")
    $dashboardCsv | Out-File -FilePath $dashboardCsvPath -Encoding utf8
    
    Import-Csv -Path $dashboardCsvPath -Delimiter '|' | Export-Excel -Path $dashboardFilePath -WorksheetName "Dashboard" -MoveToStart -AutoSize
    
    Remove-Item -Path $dashboardCsvPath -Force
    
    
    if ($historicalData.Count -gt 1) {
        Write-Log "Adding compliance trend chart to dashboard..."
        
        $excel = Open-ExcelPackage -Path $dashboardFilePath
        
        Add-Worksheet -ExcelPackage $excel -WorksheetName "ComplianceChart"
        
        $pivotParams = @{
            PivotTableName    = "CompliancePivot"
            Address           = "ComplianceChart!A1"
            SourceWorksheet   = "Historical Data"
            PivotRows         = "Month"
            PivotData         = @{"CompliancePercent" = "Average"}
            PivotChartType    = "Line"
            ChartTitle        = "Monthly Update Compliance"
            Width             = 800
            Height            = 400
            Activate          = $false
        }
        
        try {
            Add-PivotChart -ExcelPackage $excel @pivotParams -ErrorAction Stop
            Write-Log "Successfully added compliance chart"
        }
        catch {
            Write-Log "Error adding compliance chart: $_" -Type "WARNING"
            Write-Log "Continuing without chart"
        }
        
        Close-ExcelPackage $excel
    }
    else {
        Write-Log "Not enough historical data for chart (need at least 2 months)"
    }
    
    try {
        $fileBytes = [System.IO.File]::ReadAllBytes($dashboardFilePath)
        
        $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/drives/$SharePointDriveId/root:/$uploadPath`:/content"
        
        Write-Log "Uploading dashboard to SharePoint..."
        
        $response = Invoke-RestMethod -Uri $uploadUri -Headers @{ Authorization = "Bearer $token" } -Method PUT -Body $fileBytes -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        Write-Log "Dashboard uploaded successfully to SharePoint"
        
        if (Test-Path -Path $dashboardFilePath) {
            Remove-Item -Path $dashboardFilePath -Force
        }
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Type "ERROR"
        throw $_
    }
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        try {
            $notificationBody = @{
                text = "Windows Update Compliance Dashboard"
                title = "Windows Update Compliance Dashboard Updated"
                themeColor = "0076D7"
                sections = @(
                    @{
                        facts = @(
                            @{
                                name = "Total Devices:"
                                value = $latestData.TotalDevices.ToString()
                            },
                            @{
                                name = "Updated Devices ($($latestData.Month)):"
                                value = $latestData.UpdatedDevices.ToString()
                            },
                            @{
                                name = "Compliance:"
                                value = $latestData.CompliancePercent
                            }
                        )
                    }
                )
                potentialAction = @(
                    @{
                        "@type" = "OpenUri"
                        name = "View Dashboard"
                        targets = @(
                            @{
                                os = "default"
                                uri = $response.webUrl
                            }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 10
            
            Write-Log "Sending Teams notification..."
            Invoke-RestMethod -Uri $TeamsWebhookUrl -Method POST -Body $notificationBody -ContentType 'application/json'
            Write-Log "Teams notification sent successfully"
        }
        catch {
            Write-Log "Failed to send Teams notification: $_" -Type "WARNING"
        }
    }
    
    Write-Log "=== Windows Update Dashboard Generation Completed ==="
    
    return @{
        DashboardName = $dashboardFileName
        DashboardUrl = $response.webUrl
        DeviceCount = $latestData.TotalDevices
        CompliancePercent = $latestData.CompliancePercent
    }
}
catch {
    $errorMessage = $_.Exception.Message
    $errorType = $_.Exception.GetType().FullName
    $errorLine = $_.InvocationInfo.ScriptLineNumber
    Write-Log "Script execution failed at line $errorLine ($errorType): $errorMessage" -Type "ERROR"
    
    if ($dashboardFilePath -and (Test-Path -Path $dashboardFilePath)) {
        try {
            Remove-Item -Path $dashboardFilePath -Force -ErrorAction SilentlyContinue
            Write-Log "Cleaned up temp file: $dashboardFilePath" -Type "INFO"
        }
        catch {
            Write-Log "Failed to clean up temp file: $_" -Type "WARNING"
        }
    }
    
    throw $_
}