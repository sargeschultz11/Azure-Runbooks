# Get-IntuneDiscoveredAppsReport.ps1

## Overview
This Azure Automation runbook script automatically generates a report of all discovered applications in Microsoft Intune. It exports the data to an Excel spreadsheet and uploads it to a specified SharePoint document library. The report includes application details, installation counts, and a summary analysis of the most common publishers. An optional Teams webhook alert can be enabled as well if you choose.

## Purpose
The primary purpose of this script is to provide regular reporting and visibility into applications present on managed devices by:
- Retrieving all detected applications from Intune with their installation counts
- Organizing the data into a structured Excel report with summary analytics
- Automating the report distribution via SharePoint
- Implementing robust error handling and API throttling management
- Optionally sending notifications via Microsoft Teams webhooks

This automation helps IT administrators maintain better visibility into their application landscape across managed devices, identify unauthorized software, and support software license compliance efforts.

## Prerequisites
- An Azure Automation account with a System-Assigned Managed Identity enabled
- The ImportExcel PowerShell module installed in the Automation account
- The Az.Accounts module installed in the Automation account
- Proper Microsoft Graph API permissions assigned to the Managed Identity:
  - `DeviceManagementManagedDevices.Read.All` or `DeviceManagementManagedDevices.ReadWrite.All`
  - `DeviceManagementApps.Read.All`
  - `Sites.ReadWrite.All` (for SharePoint upload functionality)
  - `Reports.Read.All`
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| UseManagedIdentity | Switch | No | When specified, uses the System-Assigned Managed Identity for authentication. Default is `$true`. |
| TenantId | String | No | Your Azure AD tenant ID. Only needed if not using Managed Identity. |
| SharePointSiteId | String | Yes | The ID of the SharePoint site where the report will be uploaded. |
| SharePointDriveId | String | Yes | The ID of the document library drive where the report will be uploaded. |
| FolderPath | String | No | The folder path within the document library for upload. Default is root. |
| BatchSize | Int | No | Number of apps to retrieve in each batch. Default is 100. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying. Default is 5. |
| TeamsWebhookUrl | String | No | Optional. Microsoft Teams webhook URL for sending notifications about the report. |

## Report Contents
The generated Excel report includes:

### "Discovered Apps" Tab
A table with the following columns:
- Application Name
- Publisher
- Version
- Device Count
- Platform
- Size in Bytes
- App ID

### "Summary" Tab
- Report metadata (generation date, system info)
- Total number of discovered apps
- Top 10 publishers with app counts
- Platform summary (distribution of apps by platform)

## Setup Instructions

### 1. Configure System-Assigned Managed Identity
1. In your Azure Automation account, navigate to Identity
2. Enable the System-assigned identity
3. Copy the Object ID of the Managed Identity for later use

### 2. Assign Graph API Permissions
Run the included `Add-GraphPermissions.ps1` script to assign the necessary permissions to your Automation Account's Managed Identity:

```powershell
.\Add-GraphPermissions.ps1 -AutomationMSI_ID "<AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
```

The script will assign the following permissions:
- DeviceManagementManagedDevices.Read.All
- DeviceManagementApps.Read.All
- Sites.ReadWrite.All
- Reports.Read.All

### 3. Get SharePoint Site and Drive IDs
1. You'll need the SharePoint site ID and document library drive ID where reports will be uploaded
2. These can be obtained using Graph Explorer or PowerShell
   - Site ID format: `sitecollections/{site-collection-id}/sites/{site-id}`
   - Drive ID format: `b!{encoded-drive-id}`

### 4. Set Up Azure Automation Account
1. Create or use an existing Azure Automation account with System-Assigned Managed Identity enabled
2. Import the required modules:
   - ImportExcel module
     - Browse to Modules > Browse gallery > Search for "ImportExcel" > Import
   - Az.Accounts module
     - Browse to Modules > Browse gallery > Search for "Az.Accounts" > Import

### 5. Import the Runbook
1. In the Automation account, go to Runbooks > Import a runbook
2. Upload the Get-IntuneDiscoveredAppsReport.ps1 file
3. Set the runbook type to PowerShell

### 6. Schedule the Runbook
1. Navigate to the runbook > Schedules > Add a schedule
2. Create a new schedule or link to an existing one
3. Configure the parameters, including SharePointSiteId and SharePointDriveId

### 7. Optional: Set Up Teams Notification
1. Create a Teams webhook connector in your desired Teams channel
2. Copy the webhook URL
3. Add the TeamsWebhookUrl parameter when scheduling the runbook

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the System-Assigned Managed Identity.
2. **Data Retrieval**: Gets all discovered apps from Intune using the direct API endpoint.
3. **Excel Report Generation**: Creates the Excel report with app data and summaries.
4. **SharePoint Upload**: Uploads the report to the specified SharePoint location.
5. **Teams Notification**: Optionally sends a notification card to Teams with report details.
6. **Cleanup**: Removes temporary files and returns execution summary.

## Managed Identity Authentication
The script uses multiple approaches to acquire a token using Managed Identity:
- First attempts standard token acquisition with `Get-AzAccessToken`
- Falls back to token cache inspection if the standard approach fails
- Final fallback attempts token exchange using managed identity credentials

This multi-level approach ensures robust authentication in various Azure environments.

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Exponential Backoff**: Implements exponential backoff for throttled requests
- **Retry Logic**: Automatically retries failed requests with increasing backoff periods
- **Retry-After Header**: Respects the Retry-After header from Microsoft Graph API when provided

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| ReportName | Name of the generated report file |
| AppsCount | Total number of apps in the report |
| ReportUrl | SharePoint URL to the uploaded report |
| ExecutionTimeMinutes | Total execution time in minutes |
| Timestamp | Report generation timestamp |
| NotificationSent | Boolean indicating whether Teams notification was sent successfully (only if TeamsWebhookUrl is provided) |

## Logging
The script utilizes verbose logging to provide detailed information about each step:
- All log entries include timestamps and log levels (INFO, WARNING, ERROR)
- Detailed error information when issues occur

## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API throttling is handled gracefully with exponential backoff
- File system operations are wrapped in try-catch blocks
- Temporary files are cleaned up even when errors occur
- Module dependencies are checked and installed if missing

## Notes
- The ImportExcel and Az.Accounts modules must be imported into the Azure Automation account
- The report includes applications detected across all managed device platforms (Windows, iOS, Android, MacOS)
- Make sure the SharePoint folder path exists before running the script
- Teams notifications include an adaptive card with a direct link to the report
