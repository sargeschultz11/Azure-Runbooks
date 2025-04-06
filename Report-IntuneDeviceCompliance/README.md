# Get-IntuneDeviceComplianceReport.ps1

## Overview
This Azure Automation runbook script automatically generates a report of all enrolled device compliance statuses in Microsoft Intune. It exports the data to an Excel spreadsheet and uploads it to a specified SharePoint document library. The report includes detailed device compliance information, compliance states, and summary analytics. An optional Teams webhook alert can be configured to notify stakeholders when new reports are generated.

## Purpose
The primary purpose of this script is to provide regular reporting and visibility into device compliance across your Intune-managed environment by:
- Retrieving all enrolled devices from Intune with their compliance status
- Collecting details about compliance policies applied to each device
- Organizing the data into a structured Excel report with compliance statistics
- Automating the report distribution via SharePoint
- Implementing robust error handling and API throttling management
- Optionally sending notifications via Microsoft Teams webhooks

This automation helps IT administrators maintain better visibility into device compliance, identify non-compliant devices, and ensure security requirements are met across the organization.

## Prerequisites
- An Azure Automation account
- The ImportExcel PowerShell module installed in the Automation account
- Authentication using either:
  - **Recommended: System-assigned Managed Identity** with the following:
    - Managed Identity enabled on the Azure Automation account
    - Microsoft Graph API permissions assigned to the Managed Identity using the included `Add-GraphPermissions.ps1` helper script
  - **Alternative: Azure AD App Registration** (legacy approach)
- For Azure Automation with Managed Identity, the Az.Accounts module installed
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| UseManagedIdentity | Switch | No | When specified, the script will use the Managed Identity of the Azure Automation account for authentication. This is now the default and recommended authentication method. |
| SharePointSiteId | String | Yes | The ID of the SharePoint site where the report will be uploaded. |
| SharePointDriveId | String | Yes | The ID of the document library drive where the report will be uploaded. |
| FolderPath | String | No | The folder path within the document library for upload. Default is root. |
| BatchSize | Int | No | Number of devices to retrieve in each batch. Default is 100. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying. Default is 5. |
| TeamsWebhookUrl | String | No | Optional. Microsoft Teams webhook URL for sending notifications about the report. |

## Report Contents
The generated Excel report includes:

### "Device Compliance" Tab
A table with the following columns:
- Device Name
- User
- Email
- UPN (User Principal Name)
- Device Owner
- Device Type
- OS
- OS Version
- Compliance State
- Compliance Policies
- Policy Statuses
- Last Sync
- Enrolled Date
- Serial Number
- Model
- Manufacturer

### "Summary" Tab
- Report metadata (generation date, system info)
- Total number of enrolled devices
- Compliance status breakdown (compliant, non-compliant, etc.)
- Device type distribution
- Operating system distribution

## Setup Instructions

The recommended authentication method is using a System-assigned Managed Identity.

### Using Managed Identity (Recommended)

#### 1. Enable System-assigned Managed Identity
1. Navigate to your Azure Automation account
2. Go to Identity under Settings
3. Switch the Status to "On" under the System assigned tab
4. Click Save

#### 2. Assign API Permissions to the Managed Identity
Use the included `Add-GraphPermissions.ps1` script to assign the required permissions:

1. Install the Microsoft.Graph.Applications module if not already installed:
   ```powershell
   Install-Module -Name Microsoft.Graph.Applications -Force
   ```

2. Run the script with your Automation Account's MSI Object ID:
   ```powershell
   .\Add-GraphPermissions.ps1 -AutomationMSI_ID "<YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
   ```

This script will assign the following permissions:
- DeviceManagementManagedDevices.Read.All
- DeviceManagementConfiguration.Read.All
- Sites.ReadWrite.All

#### 3. Install Required Modules in Azure Automation
1. In your Azure Automation account, go to Modules
2. Add the following modules if not already installed:
   - Az.Accounts
   - ImportExcel

### 4. Get SharePoint Site and Drive IDs
1. You'll need the SharePoint site ID and document library drive ID where reports will be uploaded
2. These can be obtained using Graph Explorer or PowerShell
   - Site ID format: `sitecollections/{site-collection-id}/sites/{site-id}`
   - Drive ID format: `b!{encoded-drive-id}`

### 5. Import the Runbook
1. In the Automation account, go to Runbooks > Import a runbook
2. Upload the Get-IntuneDeviceComplianceReport.ps1 file
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
1. **Authentication**: The script authenticates to Microsoft Graph API using the system-assigned Managed Identity.
2. **Data Retrieval**: Gets all managed devices from Intune with their compliance states.
3. **Policy Lookup**: Retrieves compliance policy details and matches them to devices.
4. **Excel Report Generation**: Creates the Excel report with device data and compliance summaries.
5. **SharePoint Upload**: Uploads the report to the specified SharePoint location.
6. **Teams Notification**: Optionally sends a notification card to Teams with compliance statistics.
7. **Cleanup**: Removes temporary files and returns execution summary.

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Retrieves devices in configurable batches
- **Exponential Backoff**: Implements exponential backoff for throttled requests
- **Retry Logic**: Automatically retries failed requests with increasing backoff periods
- **Retry-After Header**: Respects the Retry-After header from Microsoft Graph API when provided

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| ReportName | Name of the generated report file |
| DevicesCount | Total number of devices in the report |
| ReportUrl | SharePoint URL to the uploaded report |
| ExecutionTimeMinutes | Total execution time in minutes |
| Timestamp | Report generation timestamp |
| ComplianceSummary | Array of compliance states and their counts |
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

## Security Best Practices

### Managed Identity vs App Registration

Using a managed identity is the recommended authentication method for Azure Automation because it eliminates the need to provision or rotate secrets and is managed by the Azure platform itself. Here are some key advantages of using managed identities:

- **No Secret Management**: Managed identities eliminate the need for developers to manage credentials when connecting to resources that support Microsoft Entra authentication.
- **Enhanced Security**: When granting permissions to a managed identity, always apply the principle of least privilege by granting only the minimal permissions needed to perform required actions.
- **Reduced Administrative Overhead**: There's no need to manually rotate secrets or manage certificate expirations
- **Simplified Deployment**: Once enabled, the system-assigned managed identity is registered with Microsoft Entra ID and can be used to access other resources protected by Microsoft Entra ID.

### Implementation Considerations

- When using Managed Identity, ensure that the Az.Accounts module is installed in your Azure Automation account
- For hybrid worker scenarios, you may need to grant additional permissions for the managed identity
- Follow the principle of least privilege and carefully assign only permissions required to execute your runbooks
- System-assigned identities are automatically deleted when the resource is deleted