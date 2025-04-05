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
  - **Option 1: Azure AD App Registration** with the following:
    - Client ID
    - Client Secret
    - Proper Microsoft Graph API permissions:
      - `DeviceManagementManagedDevices.Read.All` or `DeviceManagementManagedDevices.ReadWrite.All`
      - `Sites.ReadWrite.All` (for SharePoint upload functionality)
  - **Option 2: System-assigned Managed Identity** with the following:
    - Managed Identity enabled on the Azure Automation account
    - The same Microsoft Graph API permissions assigned to the Managed Identity
- If using App Registration, the following variables defined in the Automation account:
  - `TenantId`: Your Azure AD tenant ID
  - `ClientId`: The App Registration's client ID
  - `ClientSecret`: The App Registration's client secret (stored as an encrypted variable)
- For Azure Automation with Managed Identity, the Az.Accounts module installed
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| TenantId | String | No | Your Azure AD tenant ID. If not provided, will be retrieved from Automation variables. Not needed if using Managed Identity. |
| ClientId | String | No | The App Registration's client ID. If not provided, will be retrieved from Automation variables. Not needed if using Managed Identity. |
| ClientSecret | String | No | The App Registration's client secret. If not provided, will be retrieved from Automation variables. Not needed if using Managed Identity. |
| UseManagedIdentity | Switch | No | When specified, the script will use the Managed Identity of the Azure Automation account for authentication instead of App Registration credentials. |
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
- Device Owner
- Device Type
- OS
- OS Version
- Compliance State
- Compliance Policies
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

You can choose between two authentication methods: App Registration (service principal) or Managed Identity.

### Option 1: Using App Registration

#### 1. Create an Azure AD Application Registration
1. In the Azure Portal, navigate to Azure Active Directory > App registrations > New registration
2. Name the application (e.g., "Intune Compliance Reporting")
3. Select the appropriate supported account type (typically Single tenant)
4. Click Register

#### 2. Assign API Permissions
1. In the app registration, navigate to API permissions
2. Click "Add a permission" > Microsoft Graph > Application permissions
3. Add the following permissions:
   - DeviceManagementManagedDevices.Read.All (or ReadWrite.All)
   - Sites.ReadWrite.All
4. Click "Grant admin consent"

#### 3. Create a Client Secret
1. In the app registration, navigate to Certificates & secrets
2. Create a new client secret with an appropriate expiration
3. Copy the secret value (you won't be able to retrieve it later)

### Option 2: Using Managed Identity (Recommended)

#### 1. Enable System-assigned Managed Identity
1. Navigate to your Azure Automation account
2. Go to Identity under Settings
3. Switch the Status to "On" under the System assigned tab
4. Click Save

#### 2. Assign API Permissions to the Managed Identity
1. Go to Azure Active Directory > Enterprise applications
2. Find the managed identity (it will have the same name as your Automation account)
3. Go to Permissions > Add permission > Microsoft Graph > Application permissions
4. Add the following permissions:
   - DeviceManagementManagedDevices.Read.All (or ReadWrite.All)
   - Sites.ReadWrite.All
5. Click "Grant admin consent"

#### 3. Install Required Modules
1. In your Azure Automation account, go to Modules
2. Add the Az.Accounts module if it's not already installed

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
2. Upload the Get-IntuneDeviceComplianceReport.ps1 file
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

## Notes
- For large environments with thousands of devices, consider adjusting the BatchSize parameter
- The ImportExcel module must be imported into the Azure Automation account
- The report includes all enrolled devices across platforms (Windows, iOS, Android, MacOS)
- Make sure the SharePoint folder path exists before running the script
- Teams notifications include an adaptive card with compliance rate and a direct link to the report
- The script uses the Microsoft Graph beta endpoint to retrieve more detailed compliance information

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
- Follow the principle of least privilege and carefully assign only permissions required to execute your runbooks.
- System-assigned identities are automatically deleted when the resource is deleted, while user-assigned identities have independent lifecycles.
- Role assignments aren't automatically deleted when managed identities are deleted, so remember to clean up permissions when they're no longer needed
