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
- An Azure Automation account
- The ImportExcel PowerShell module installed in the Automation account
- An Azure AD App Registration with the following:
  - Client ID
  - Client Secret
  - Proper Microsoft Graph API permissions:
    - `DeviceManagementManagedDevices.Read.All` or `DeviceManagementManagedDevices.ReadWrite.All`
    - `Sites.ReadWrite.All` (for SharePoint upload functionality)
- The following variables defined in the Automation account:
  - `TenantId`: Your Azure AD tenant ID
  - `ClientId`: The App Registration's client ID
  - `ClientSecret`: The App Registration's client secret (stored as an encrypted variable)
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| TenantId | String | No | Your Azure AD tenant ID. If not provided, will be retrieved from Automation variables. |
| ClientId | String | No | The App Registration's client ID. If not provided, will be retrieved from Automation variables. |
| ClientSecret | String | No | The App Registration's client secret. If not provided, will be retrieved from Automation variables. |
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
- Size in Bytes
- App ID

### "Summary" Tab
- Report metadata (generation date, system info)
- Total number of discovered apps
- Top 10 publishers with app counts

## Setup Instructions

### 1. Create an Azure AD Application Registration
1. In the Azure Portal, navigate to Azure Active Directory > App registrations > New registration
2. Name the application (e.g., "Intune Reporting")
3. Select the appropriate supported account type (typically Single tenant)
4. Click Register

### 2. Assign API Permissions
1. In the app registration, navigate to API permissions
2. Click "Add a permission" > Microsoft Graph > Application permissions
3. Add the following permissions:
   - DeviceManagementManagedDevices.Read.All (or ReadWrite.All)
   - Sites.ReadWrite.All
4. Click "Grant admin consent"

### 3. Create a Client Secret
1. In the app registration, navigate to Certificates & secrets
2. Create a new client secret with an appropriate expiration
3. Copy the secret value (you won't be able to retrieve it later)

### 4. Get SharePoint Site and Drive IDs
1. You'll need the SharePoint site ID and document library drive ID where reports will be uploaded
2. These can be obtained using Graph Explorer or PowerShell
   - Site ID format: `sitecollections/{site-collection-id}/sites/{site-id}`
   - Drive ID format: `b!{encoded-drive-id}`

### 5. Set Up Azure Automation Account
1. Create or use an existing Azure Automation account
2. Import the ImportExcel module
   - Browse to Modules > Browse gallery > Search for "ImportExcel" > Import
3. Create the following Automation variables:
   - Name: TenantId, Type: String, Value: Your tenant ID
   - Name: ClientId, Type: String, Value: Your application's client ID
   - Name: ClientSecret, Type: String, Value: Your client secret, Encrypted: Yes

### 6. Import the Runbook
1. In the Automation account, go to Runbooks > Import a runbook
2. Upload the Get-IntuneDiscoveredAppsReport.ps1 file
3. Set the runbook type to PowerShell

### 7. Schedule the Runbook
1. Navigate to the runbook > Schedules > Add a schedule
2. Create a new schedule or link to an existing one
3. Configure the parameters, including SharePointSiteId and SharePointDriveId

### 8. Optional: Set Up Teams Notification
1. Create a Teams webhook connector in your desired Teams channel
2. Copy the webhook URL
3. Add the TeamsWebhookUrl parameter when scheduling the runbook

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API.
2. **Data Retrieval**: Gets all discovered apps from Intune.
3. **Excel Report Generation**: Creates the Excel report with app data and summaries.
4. **SharePoint Upload**: Uploads the report to the specified SharePoint location.
5. **Teams Notification**: Optionally sends a notification card to Teams with report details.
6. **Cleanup**: Removes temporary files and returns execution summary.

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Retrieves apps in configurable batches
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
- Progress indicators for batch processing
- Detailed error information when issues occur

## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API throttling is handled gracefully with exponential backoff
- File system operations are wrapped in try-catch blocks
- Temporary files are cleaned up even when errors occur
- Module dependencies are checked and installed if missing

## Notes
- For large environments with thousands of applications, consider adjusting the BatchSize parameter
- The ImportExcel module must be imported into the Azure Automation account
- The report includes applications detected across all managed device platforms (Windows, iOS, Android, MacOS)
- Make sure the SharePoint folder path exists before running the script
- Teams notifications include an adaptive card with a direct link to the report

## Author Information
- **Author**: Ryan Schultz
- **Version**: 1.0
- **Creation Date**: 2025-04-04