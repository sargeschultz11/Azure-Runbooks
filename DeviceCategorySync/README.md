# Update-IntuneDeviceCategories.ps1

## Overview
This Azure Automation runbook script automatically updates the device categories of Windows, iOS, Android, and Linux devices in Microsoft Intune based on the primary user's department. It fetches all devices and updates the device category to match the department name of the assigned primary user, with built-in batching and throttling handling for large environments.

## Purpose
The primary purpose of this script is to ensure consistent device categorization in Intune by:
- Identifying devices (Windows, iOS, Android, and Linux) with missing or mismatched categories
- Retrieving the primary user's department information
- Setting the device category to match the user's department when available
- Processing devices in batches to avoid API throttling in large environments

This automation helps maintain better organization within the Intune portal and can be used for device targeting, reporting, and policy assignment. It is also useful for creating dynamic groups.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`
  - `DeviceManagementManagedDevices.ReadWrite.All`
  - `DeviceManagementServiceConfig.ReadWrite.All`
  - `DeviceManagementConfiguration.ReadWrite.All`
  - `User.Read.All`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account
- **IMPORTANT**: Device categories must be pre-created in Intune and must match **exactly** the department names in user account properties in Azure AD

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| WhatIf | Switch | No | If specified, shows what changes would occur without actually making any updates. |
| OSType | String | No | Specifies which operating systems to process. Valid values are "All", "Windows", "iOS", "Android", "Linux". Default is "All". |
| BatchSize | Int | No | Number of devices to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity.
2. **Device Category Retrieval**: Retrieves all device categories defined in Intune.
3. **Device Retrieval**: Gets all specified devices (Windows, iOS, Android, Linux, or any combination) from Intune.
4. **Batch Processing**: Divides devices into batches of the specified size.
5. **Processing Loop**: For each batch:
   - Processes each device in the batch
   - Checks if a device category is already assigned
   - Retrieves the primary user of the device
   - Gets the user's department information
   - If the department exists as a device category and differs from the current device category, updates the device's category
   - Waits for the specified delay period before processing the next batch

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Processes devices in configurable batches (default: 50 devices per batch)
- **Delay Between Batches**: Automatically pauses between batches (default: 10 seconds) to avoid overwhelming the Graph API
- **Throttling Detection**: Automatically detects when the Graph API returns throttling responses (HTTP 429)
- **Retry Logic**: Implements exponential backoff retry logic when throttled
- **Respect for Retry-After**: Honors the Retry-After header when provided by the Graph API
- **Server Error Handling**: Also handles 5xx server errors with retries

These features make the script suitable for large organizations with thousands of devices, as it can gracefully handle API rate limits.

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalDevices | Total number of devices processed |
| AlreadyCategorized | Number of devices with categories already matching departments |
| Updated | Number of devices that had their categories updated |
| Skipped | Number of devices skipped (no primary user, no department, or department not a category) |
| Errors | Number of devices that encountered errors during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| BatchesProcessed | Number of batches processed |
| *OSName*Devices | Total number of devices processed for each OS type (Windows, iOS, Android, Linux) |
| *OSName*Updated | Number of devices updated for each OS type |
| *OSName*Matched | Number of devices already properly categorized for each OS type |
| *OSName*Skipped | Number of devices skipped for each OS type |
| *OSName*Errors | Number of devices with errors for each OS type |

## Logging
The script utilizes verbose logging to provide detailed information about each step:
- All log entries include timestamps and log levels (INFO, WARNING, ERROR, WHATIF)
- Write-Verbose is used for standard logging in Azure Automation
- Specific error cases are captured and logged appropriately
- OS-specific statistics are maintained separately
- Batch processing status and progress are logged
- API throttling events are logged with retry information

## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API request errors are logged with details
- Throttling errors are handled with exponential backoff retries
- Device processing errors are isolated to prevent the entire script from failing
- Summary statistics include error counts for all device types (Windows, iOS, Android, and Linux)

## Notes
- **CRITICAL REQUIREMENT**: The script depends on exact matching between department names in Azure AD and device category names in Intune. If these don't match exactly, the categorization will not work.
- Before running this script, ensure that all departments used in your organization have corresponding device categories created in Intune with identical naming.
- For large environments (thousands of devices), consider adjusting the BatchSize and BatchDelaySeconds parameters to avoid throttling.
- Devices without primary users or where the user has no department are skipped
- The script counts and reports cases where department names don't exist as device categories
- For devices that already have the correct category assigned, no changes are made
- If department names in Azure AD don't match device categories in Intune exactly (including case, spacing, and special characters), the script will report these as skipped devices
- After making changes to Managed Identity permissions, it may take some time (up to an hour) for permissions to fully propagate through Azure's systems
