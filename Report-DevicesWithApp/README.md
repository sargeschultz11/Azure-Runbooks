# Get-DevicesWithAppReport.ps1

## Overview
This Azure Automation runbook generates a comprehensive report of all devices that have a specific application installed in your Intune environment. The script identifies devices with the specified app (using the App ID from the [Intune Discovered Apps report](../Report-DiscoveredApps/README.md) also available in this repository), creates a detailed Excel report, and uploads it to a SharePoint document library for easy access.

> **Note:** This solution is particularly valuable for smaller organizations or those without Enterprise/Premium licensing tiers that include Microsoft Defender for Endpoint or Microsoft Sentinel. Organizations with these advanced security products can use their built-in query capabilities for more efficient app discovery reporting.

## Purpose
The primary purpose of this script is to provide visibility into application distribution across your device fleet by:
- Identifying all devices that have a specific application installed
- Gathering detailed device information for each device with the app
- Creating a comprehensive Excel report with device and application details
- Providing summary analytics about the devices (OS distribution, ownership types, etc.)
- Automating the report distribution via SharePoint
- Optionally sending notifications via Microsoft Teams

This automation helps IT administrators track software usage, manage license compliance, identify unauthorized software installations, and support application lifecycle management.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The ImportExcel PowerShell module installed in the Automation account
- The Az.Accounts module installed in the Automation account
- The following Microsoft Graph API permissions assigned to the Managed Identity:
  - `DeviceManagementManagedDevices.Read.All`
  - `DeviceManagementApps.Read.All`
  - `DeviceManagementConfiguration.Read.All` (if including detailed device information)
  - `Sites.ReadWrite.All` (for SharePoint upload functionality)
  - `User.Read.All`
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| AppId | String | Yes | The ID of the application to search for. This is the App ID from the Intune Discovered Apps report. |
| SharePointSiteId | String | Yes | The ID of the SharePoint site where the report will be uploaded. |
| SharePointDriveId | String | Yes | The ID of the document library drive where the report will be uploaded. |
| FolderPath | String | No | The folder path within the document library for upload. Default is root. |
| IncludeDeviceDetails | Switch | No | When specified, includes more detailed device information (compliance status, etc.) in the report. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying. Default is 5. |
| TeamsWebhookUrl | String | No | Optional. Microsoft Teams webhook URL for sending notifications about the report. |

## Report Contents
The generated Excel report includes:

### "Devices With App" Tab
A table with detailed information about each device with the application installed:
- Device Name
- Primary User
- User Display Name
- Operating System and Version
- Model and Manufacturer
- Serial Number
- Device Ownership
- Last Sync DateTime
- Enrollment DateTime
- Device Category
- App Name, Publisher, Version, and ID
- Compliance Status (if IncludeDeviceDetails is specified)

### "Summary" Tab
- Report metadata (generation date, system info)
- Application details (name, publisher, version)
- Total number of devices with the app installed
- OS distribution breakdown
- Device ownership distribution
- Device category distribution
- Compliance status distribution (if IncludeDeviceDetails is specified)

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
2. Upload the Get-DevicesWithAppReport.ps1 file
3. Set the runbook type to PowerShell

### 6. Create a Schedule or use with Webhooks
1. For regular reporting, create a schedule with the parameters needed
2. For on-demand usage, configure a webhook that can accept the AppId as a parameter

### 7. Optional: Set Up Teams Notification
1. Create a Teams webhook connector in your desired Teams channel
2. Copy the webhook URL
3. Add the TeamsWebhookUrl parameter when scheduling the runbook

## Usage Examples

### Scheduled Report for a Specific Application
```powershell
$params = @{
    AppId = "12345678-1234-1234-1234-123456789012" # The App ID from Intune Discovered Apps
    SharePointSiteId = "contoso.sharepoint.com,guid,guid"
    SharePointDriveId = "b!encoded_drive_id"
    FolderPath = "IT Reports/Application Reports"
    IncludeDeviceDetails = $true
    TeamsWebhookUrl = "https://contoso.webhook.office.com/webhookb2/..."
}

Start-AzAutomationRunbook -Name "Get-DevicesWithAppReport" -Parameters $params
```

### Integration with Discovered Apps Report
This runbook can be used as a follow-up to the Get-IntuneDiscoveredAppsReport runbook, allowing IT admins to first review the discovered apps report and then generate device lists for apps of interest.

## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API throttling is handled gracefully with exponential backoff
- File system operations are wrapped in try-catch blocks
- Temporary files are cleaned up even when errors occur
- Module dependencies are checked and installed if missing

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| AppId | The ID of the application that was searched for |
| AppName | The display name of the application |
| AppPublisher | The publisher of the application |
| AppVersion | The version of the application |
| DevicesCount | Total number of devices with the application installed |
| ReportName | Name of the generated report file |
| ExecutionTimeMinutes | Total execution time in minutes |
| Timestamp | Report generation timestamp |
| ReportUrl | SharePoint URL to the uploaded report |
| NotificationSent | Boolean indicating whether Teams notification was sent successfully (only if TeamsWebhookUrl is provided) |

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Exponential Backoff**: Implements exponential backoff for throttled requests
- **Retry Logic**: Automatically retries failed requests with increasing backoff periods
- **Retry-After Header**: Respects the Retry-After header from Microsoft Graph API when provided

## Alternative Solutions

For organizations with Enterprise or Premium licensing, Microsoft provides built-in tools that may be more efficient for app discovery reporting:

1. **Microsoft Defender for Endpoint** - Advanced Hunting feature allows running KQL queries to identify devices with specific software installed. This provides near real-time data through the existing security agent infrastructure.

2. **Microsoft Sentinel** - Offers more powerful correlation capabilities to join software inventory data with other security telemetry.

These built-in solutions offer significant advantages over custom scripts for organizations with the appropriate licensing:
- Near real-time data without manual report generation
- Integration with the broader security ecosystem
- More powerful query capabilities and visualization options
- Data retention according to your existing policies
- Built-in access controls and permissions

## Notes
- When using IncludeDeviceDetails, the report generation may take longer due to additional API calls
- For large environments with thousands of devices, consider running the script during off-hours
- App IDs can be obtained from the Intune portal or from the Get-IntuneDiscoveredAppsReport runbook
- This runbook complements the existing runbooks in the Azure-Runbooks repo, particularly the Get-IntuneDiscoveredAppsReport.ps1
- Consider this solution primarily when your organization lacks Microsoft Defender for Endpoint or Sentinel licensing, or when you need specific reporting automation capabilities not available in those tools