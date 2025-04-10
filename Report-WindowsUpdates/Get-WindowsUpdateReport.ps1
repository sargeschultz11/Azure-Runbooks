<#
.SYNOPSIS
    Generates a monthly compliance report from Windows Update for Business data using Log Analytics.

.DESCRIPTION
    This Azure Automation runbook connects to a Log Analytics workspace using a System-Assigned Managed Identity,
    runs a Kusto query against the UCClientUpdateStatus table, and aggregates device-level update data across
    multiple months. The script builds a Yes/No update matrix per device per month, creates summary statistics,
    and exports the results to an Excel workbook. The final report is uploaded to a SharePoint document library,
    and an optional Teams notification can be sent.

.PARAMETER MonthsToReport
    The number of months to include in the report (counting backwards from current month). Default is 3.

.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the report will be uploaded.

.PARAMETER SharePointDriveId
    The ID of the document library drive where the report will be uploaded.

.PARAMETER FolderPath
    Optional. The folder path within the document library where the report will be uploaded.
    If not specified, the file will be uploaded to the root of the document library.

.PARAMETER TeamsWebhookUrl
    Optional. Microsoft Teams webhook URL for sending notifications about the report.

.PARAMETER WorkspaceId
    The ID of the Log Analytics workspace to query for update data.

.NOTES
    File Name: Get-WindowsUpdateReport.ps1
    Author: Ryan Schultz
    Version: 1.1
    Updated: 2025-04-10

    Requires -Modules ImportExcel, Az.Accounts
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$MonthsToReport = 3,

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

# Connect to Azure
Connect-AzAccount -Identity | Out-Null

# Calculate start date for KQL
$startDate = (Get-Date).AddMonths(-$MonthsToReport)
$startDateStr = $startDate.ToString("yyyy-MM-dd")

# Build KQL Query using real WUfB table: UCClientUpdateStatus
$kqlQuery = @"
UCClientUpdateStatus
| where TimeGenerated >= datetime($startDateStr)
| extend Month = format_datetime(TimeGenerated, 'yyyy-MM')
| summarize UpdateCount = count(), LastSeen = max(TimeGenerated) by DeviceName, Month
"@

# Run the query against Log Analytics
$results = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $kqlQuery
if (-not $results.Results) {
    Write-Output "No update data found."
    return
}

# Pivot into update matrix
$matrix = @{}
$months = @()
foreach ($row in $results.Results) {
    $device = $row.DeviceName
    $month = $row.Month
    if (-not $months.Contains($month)) { $months += $month }
    if (-not $matrix.ContainsKey($device)) { $matrix[$device] = @{} }
    $matrix[$device][$month] = "Yes"
}
$months = $months | Sort-Object

# Build Excel rows
$reportData = @()
foreach ($device in $matrix.Keys) {
    $row = [ordered]@{ "Device Name" = $device }
    foreach ($month in $months) {
        $row[$month] = $matrix[$device].ContainsKey($month) ? "Yes" : "No"
    }
    $reportData += [pscustomobject]$row
}

# Create summary stats
$summary = @()
foreach ($month in $months) {
    $updated = $reportData | Where-Object { $_.$month -eq "Yes" } | Measure-Object | Select-Object -ExpandProperty Count
    $total = $reportData.Count
    $percent = if ($total -gt 0) { [math]::Round(($updated / $total) * 100, 2) } else { 0 }
    $summary += [pscustomobject]@{
        Month = $month
        TotalDevices = $total
        UpdatedDevices = $updated
        CompliancePercent = "$percent%"
    }
}

# Export to Excel
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$path = "$env:TEMP\UpdateComplianceReport_$timestamp.xlsx"
$reportData | Export-Excel -Path $path -WorksheetName "Update Matrix" -AutoSize
$summary | Export-Excel -Path $path -WorksheetName "Summary" -AutoSize

# Upload to SharePoint
try {
    Connect-AzAccount -Identity | Out-Null
    $graphToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $token = $graphToken.Token
    if (-not $token) {
        throw "Failed to retrieve Graph access token using Managed Identity."
    }
} catch {
    Write-Error "Unable to authenticate with Microsoft Graph: $_"
    return
}
$fileBytes = [System.IO.File]::ReadAllBytes($path)
$uploadUrl = "https://graph.microsoft.com/v1.0/sites/$SharePointSiteId/drives/$SharePointDriveId/root:/$($FolderPath)/UpdateComplianceReport_$timestamp.xlsx:/content"
Invoke-RestMethod -Uri $uploadUrl -Headers @{ Authorization = "Bearer $token" } -Method PUT -Body $fileBytes -ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Optional Teams Notification
if ($TeamsWebhookUrl) {
    $body = @{ text = "Windows Update Compliance report generated and uploaded to SharePoint: UpdateComplianceReport_$timestamp.xlsx" } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Uri $TeamsWebhookUrl -Method POST -Body $body -ContentType 'application/json'
}

Write-Output "Done! Excel report uploaded and notifications sent (if configured)."
