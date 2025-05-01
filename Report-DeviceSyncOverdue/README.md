# Report-DeviceSyncOverdue

## Overview
This Azure Automation runbook identifies and reports on Intune-managed devices that haven't synced within a specified time threshold. Unlike the Alert-DeviceSyncReminder solution (which sends notifications to users), this solution creates a comprehensive report that's stored in Azure Blob Storage for administrative review and record-keeping.

## Purpose
The primary purpose of this solution is to provide IT administrators with visibility into devices that may be disconnected from management by:
- Identifying all managed devices (Windows, iOS, Android, macOS) that haven't synced within a configurable timeframe
- Collecting detailed device and user information for reporting and analysis
- Generating structured reports in multiple formats (CSV, JSON, HTML)
- Storing reports in Azure Blob Storage for easy access and retention
- Providing device sync statistics broken down by operating system and user assignment

This automation helps organizations maintain better visibility into their device management health, identify potential compliance issues, and track devices that may require administrative attention or follow-up.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`
  - `User.Read.All`
- The Managed Identity must also have Contributor access to the Azure Storage Account
- The following PowerShell modules must be imported into your Azure Automation account:
  - `Az.Accounts`
  - `Az.Storage`

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| DaysSinceLastSync | Int | No | The number of days to use as a threshold for determining "stale" devices. Default is 7 days. |
| StorageAccountName | String | Yes | The name of the Azure Storage Account where the report will be stored. |
| StorageContainerName | String | Yes | The name of the Blob container in the Storage Account where the report will be stored. |
| ExcludedDeviceCategories | String[] | No | Array of device categories to exclude from the report. |
| BatchSize | Int | No | Number of devices to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying. Default is 5. |
| WhatIf | Switch | No | If specified, shows what would be done without creating or uploading any reports. |
| ReportFormat | String | No | Format for the generated report. Options are "CSV", "JSON", "HTML". Default is "CSV". |
| IncludeDetailedDeviceInfo | Switch | No | If specified, includes additional device information in the report. |

## Report Formats
The script supports three different report formats, each with its own advantages:

### CSV Format (Default)
- Best for data analysis and importing into other systems
- Compatible with Excel and other spreadsheet applications
- Easy to filter, sort, and manipulate data
- Most compact file size

### JSON Format
- Structured data format with rich metadata
- Suitable for programmatic consumption
- Includes summary statistics and configuration information
- Preserves all data types (dates, numbers, etc.)

### HTML Format
- Human-readable report with formatting and styling
- Color-coded status indicators for device sync age
- Visual summaries and organized sections
- Can be opened directly in web browsers

## Report Contents
The generated report includes the following information for each device:

### Basic Device Information
- Device Name
- Device ID
- Operating System Type and Version
- Last Sync DateTime
- Days Since Last Sync
- Device Category
- Serial Number
- Model and Manufacturer
- Enrollment DateTime
- Compliance State
- Owner Type

### User Information
- User Display Name
- User Email Address
- User Principal Name
- Department
- Job Title
- Office Location

### Additional Information (when IncludeDetailedDeviceInfo is specified)
- Phone Number
- WiFi MAC Address
- IMEI/MEID (for mobile devices)
- Compliance Grace Expiration
- Management Agent and State
- Encryption Status
- Supervision Status
- Jailbreak Status
- Azure AD Registration Status
- Enrollment Type
- Registration State

## Setting Up Managed Identity Permissions
You can use the standard `Add-GraphPermissions.ps1` script (available in other runbook folders in this repository) to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

You'll also need to assign storage permissions to your Managed Identity:
1. Navigate to your Azure Storage account
2. Select "Access Control (IAM)"
3. Add a role assignment with:
   - Role: "Storage Blob Data Contributor" (this is critical for allowing the managed identity to write to the storage container)
   - Assign access to: "Managed Identity"
   - Select your Automation Account's Managed Identity
4. Add a second role assignment with:
   - Role: "Reader and Data Access"
   - Assign access to: "Managed Identity"
   - Select your Automation Account's Managed Identity
   
> **Important**: The "Storage Blob Data Contributor" role is specifically required to allow the managed identity to create and write blobs to the storage container. Without this permission, the runbook will fail when attempting to upload the report.

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API and Azure Storage using the Automation Account's Managed Identity.
2. **Device Retrieval**: Gets all managed devices that haven't synced since the specified threshold date.
3. **Batch Processing**: Divides devices into batches of the specified size.
4. **User Lookup**: For each device, retrieves additional information about the primary user if one exists.
5. **Report Generation**: Creates a report in the specified format (CSV, JSON, or HTML).
6. **Storage Upload**: Uploads the report to Azure Blob Storage and returns the URL.

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Processes devices in configurable batches (default: 50 devices per batch)
- **Delay Between Batches**: Automatically pauses between batches (default: 10 seconds)
- **Throttling Detection**: Automatically detects when the Graph API returns throttling responses (HTTP 429)
- **Retry Logic**: Implements exponential backoff retry logic when throttled
- **Respect for Retry-After**: Honors the Retry-After header when provided by the Graph API

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| OutdatedDevices | Total number of devices that haven't synced since the threshold date |
| NoUserCount | Number of devices without a primary user assigned |
| SkippedCategoryCount | Number of devices skipped due to excluded categories |
| ErrorCount | Number of errors encountered during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| SyncThresholdDate | The date used as the sync threshold |
| ReportUrl | URL to the uploaded report in Azure Blob Storage |
| ReportFormat | Format of the generated report (CSV, JSON, HTML) |
| *OSName*Devices | Total number of outdated devices for each OS type |
| *OSName*WithUser | Number of outdated devices with primary users for each OS type |
| *OSName*NoUser | Number of outdated devices without primary users for each OS type |

## Scheduling Recommendations
- Schedule to run weekly for regular monitoring
- Consider running more frequently (e.g., daily) in large environments with strict compliance requirements
- Use different report formats for different audiences or purposes
- Save historical reports for trend analysis

## Integration with Other Solutions
This runbook complements other solutions in this repository:
- Use with **Alert-DeviceSyncReminder** for a complete device sync management workflow
- Review reports before running **Sync-IntuneDevices** to force sync all devices
- Analyze alongside **Report-IntuneDeviceCompliance** to correlate sync status with compliance issues

## Notes and Best Practices
- For first-time use, run with the `-WhatIf` parameter to validate your configuration
- Create different report formats for different use cases (CSV for data analysis, HTML for executive review)
- Monitor storage usage and implement lifecycle management for reports
- Consider creating a Power BI dashboard to visualize historical report data
- For very large environments, consider extending the script to support Azure Table Storage for enhanced querying capabilities

## HTML Report Features
When using the HTML format, the report includes:
- Color-coded status indicators for sync age (green, orange, red)
- Highlighted rows for devices without primary users
- Device counts by operating system
- Summary statistics and metadata
- Responsive design that works on desktop and mobile

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has all required permissions for both Graph API and Azure Storage
2. Verify the Az.Accounts and Az.Storage modules are imported into your Automation account
3. Ensure the specified storage container exists or that the Managed Identity has permissions to create it
4. Review the runbook logs for specific error messages
5. Test with the WhatIf parameter to validate logic without creating or uploading reports