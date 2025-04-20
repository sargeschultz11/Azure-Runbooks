# Get-WindowsUpdateReport.ps1

## Overview
This Azure Automation runbook automatically generates a comprehensive Windows Update compliance dashboard using Log Analytics data and maintains it as a persistent Excel file in SharePoint. The dashboard provides visibility into Windows Update compliance across your device fleet, helping you track update status over time even beyond the Log Analytics retention period.

## Key Features
- **Persistent Compliance Tracking**: Maintains historical update compliance data beyond Log Analytics retention limits
- **SharePoint Integration**: Automatically stores and updates the dashboard in a SharePoint library
- **Rich Excel Dashboard**: Generates a multi-tab Excel dashboard with detailed device status
- **Compliance Trending**: Visualizes compliance rates over time with interactive charts
- **Teams Notifications**: Optionally sends Teams messages when the dashboard is updated

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The following PowerShell modules imported into your Azure Automation account:
  - ImportExcel
  - Az.Accounts
  - Az.OperationalInsights
- A Log Analytics workspace that collects Windows Update data
- A SharePoint site and document library for storing the dashboard
- Optional: A Teams webhook URL for sending notifications

## Required Permissions
The Managed Identity for your Azure Automation account must have:
- **Log Analytics**: Contributor or Reader access to the Log Analytics workspace
- **SharePoint**: Permission to read and write files to the specified document library
- **Microsoft Graph API**: The necessary permissions to access SharePoint via Graph API

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| MonthsToQuery | Int | No | Number of months to include in the query (counting backwards from current month). Default is 1. Limited by your Log Analytics workspace retention period. |
| SharePointSiteId | String | Yes | The ID of the SharePoint site where the dashboard will be stored. |
| SharePointDriveId | String | Yes | The ID of the document library drive where the dashboard will be stored. |
| FolderPath | String | No | The folder path within the document library. Default is the root folder. |
| TeamsWebhookUrl | String | No | Microsoft Teams webhook URL for sending notifications about dashboard updates. |
| WorkspaceId | String | Yes | The ID of the Log Analytics workspace to query for update data. |

## Dashboard Contents
The Excel dashboard consists of four sheets:

### 1. Dashboard
A summary view showing:
- Current month's compliance percentage
- Total, updated, and non-updated device counts
- Monthly compliance history table
- Compliance trend mini-chart

### 2. Device Status
A detailed matrix showing:
- All devices in your environment
- Update status for each device by month
- Color-coded cells (green for updated, red for not updated)
- Update count for each device by month

### 3. Historical Data
A table containing:
- Monthly compliance percentages
- Device counts by month
- Raw data for historical tracking

### 4. Compliance Trend
A full-page chart showing:
- Update compliance percentage trending over time
- Visual indication of compliance improvements or declines

## Usage Scenarios

### Basic Monitoring
Scheduled to run monthly to maintain an up-to-date dashboard of Windows Update compliance.

### Compliance Reporting
Generate reports for security teams or management to demonstrate update compliance over time.

### Trend Analysis
Identify patterns in update adoption and target devices or departments that consistently lag.

### Governance
Provide evidence of compliance for audits or regulatory requirements.

## Setup Instructions

### 1. Get SharePoint Site ID and Drive ID
1. Navigate to your SharePoint site
2. Use Microsoft Graph Explorer or PowerShell to retrieve:
   - Site ID format: `sitecollections/{site-collection-id}/sites/{site-id}`
   - Drive ID format: `b!{encoded-drive-id}`

### 2. Set Up Teams Webhook (Optional)
1. In Microsoft Teams, navigate to the channel for notifications
2. Add a new Webhook connector
3. Copy the webhook URL for the `TeamsWebhookUrl` parameter

### 3. Import Required PowerShell Modules
1. In your Azure Automation account, go to "Modules"
2. Import the modules:
   - ImportExcel
   - Az.Accounts
   - Az.OperationalInsights

### 4. Import the Runbook
1. In your Automation account, go to "Runbooks" > "Import a runbook"
2. Upload the Get-WindowsUpdateReport.ps1 file
3. Set the runbook type to PowerShell

### 5. Configure Parameter Values
When creating a schedule or starting the runbook, provide the following parameter values:
- `WorkspaceId`: Your Log Analytics workspace ID
- `SharePointSiteId`: The ID of your SharePoint site
- `SharePointDriveId`: The ID of your document library
- Other parameters as needed

### 6. Schedule the Runbook
1. Create a schedule (typically monthly)
2. Link it to the runbook with your parameter values

## Troubleshooting

### Log Analytics Table Detection
The script automatically tries several common Windows Update table names:
- UCClientUpdateStatus
- Update
- WindowsUpdates
- Update_CL

If your environment uses a different table name, you'll need to modify the script.

### Excel File Creation Issues
- Check that the ImportExcel module is properly imported
- Verify permissions to the SharePoint document library
- Ensure the FolderPath parameter specifies an existing folder

### No Update Data
- Verify that the Log Analytics workspace is collecting Windows Update data
- Confirm that the queried time period falls within your Log Analytics retention period
- Check if devices are properly reporting update status to Log Analytics

## Customization Options
- Modify the KQL query to filter for specific device groups
- Adjust the dashboard layout or add additional sheets
- Customize the Teams notification adaptive card
- Add department or business unit segmentation to the reports

## Notes
- Dashboard performance may decrease with very large device fleets
- Consider data retention in Log Analytics when setting the MonthsToQuery parameter
- The dashboard uses conditional formatting to highlight compliance status

## Output
The runbook returns a PowerShell object with:
- DashboardName: Name of the Excel file
- DashboardUrl: SharePoint URL to the dashboard
- DeviceCount: Total number of devices in the report
- CompliancePercent: Current compliance percentage
- LastUpdated: Timestamp of the dashboard update