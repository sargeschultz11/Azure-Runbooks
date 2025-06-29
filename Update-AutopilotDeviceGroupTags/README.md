# Update-AutopilotDeviceGroupTags.ps1

## Overview
This Azure Automation runbook script synchronizes Windows Autopilot device group tags with their corresponding Intune device categories. It's designed as a companion to the Update-IntuneDeviceCategories.ps1 script, completing the device categorization workflow by ensuring that group tags used during Autopilot provisioning match the device categories used for management.

## Purpose
The primary purpose of this script is to ensure consistency between Autopilot group tags and Intune device categories by:
- Identifying Autopilot devices with group tags that don't match their corresponding Intune device categories
- **Pre-filtering devices** to only process those that actually need updates, dramatically improving performance
- Updating the Autopilot group tags to match the device categories
- Processing devices in batches to avoid API throttling in large environments
- Providing detailed logging and reporting with comprehensive statistics

This automation helps maintain a consistent device categorization approach across the entire device lifecycle, from initial provisioning through Autopilot to ongoing management in Intune. It ensures that group tags used for dynamic Azure AD group membership and Autopilot deployment profile assignment align with the device categories used for policy and app targeting.

## Key Performance Features (v1.2.1)
- **Smart Pre-filtering**: Only processes devices that actually need updates, skipping devices that already have correct group tags
- **Duplicate Device Handling**: Automatically selects the most recently synced Intune device when duplicates exist with the same serial number
- **Optimized API Usage**: Significantly reduced Graph API calls by filtering devices before batch processing
- **Enhanced Array Handling**: Properly handles single-object API responses to prevent type conversion issues
- **Improved Category Validation**: Better handling of 'Unknown' vs null/empty device categories

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`
  - `DeviceManagementManagedDevices.ReadWrite.All`
  - `DeviceManagementServiceConfig.ReadWrite.All`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account
- **IMPORTANT**: Device categories must be configured in Intune and assigned to devices (manually or using the Update-IntuneDeviceCategories.ps1 script from this repository)

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| WhatIf | Switch | No | If specified, shows what changes would occur without actually making any updates. |
| BatchSize | Int | No | Number of devices to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script (from the DeviceCategorySync folder) to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity.
2. **Device Category Retrieval**: Gets all device categories defined in Intune.
3. **Autopilot Device Retrieval**: Gets all Windows Autopilot devices.
4. **Intune Device Retrieval**: Gets all Intune devices with their categories.
5. **Device Lookup Creation**: Creates an optimized lookup dictionary of Intune devices by serial number with duplicate resolution.
6. **Pre-filtering**: Identifies only the devices that need updates, skipping:
   - Devices with no matching Intune device
   - Devices with 'Unknown' or no device category
   - Devices where group tag already matches device category
7. **Batch Processing**: Divides filtered devices into batches of the specified size.
8. **Processing Loop**: For each batch:
   - Updates the group tag to match the device category
   - Waits for the specified delay period before processing the next batch

## Performance Improvements
The script includes significant performance optimizations:
- **Reduced API Calls**: Only processes devices that actually need updates
- **Faster Device Matching**: Uses hashtable lookup instead of iterative matching
- **Intelligent Filtering**: Pre-filters devices to eliminate unnecessary processing
- **Duplicate Resolution**: Automatically handles duplicate Intune devices by selecting the most recent sync
- **Batch Optimization**: Only creates batches for devices that need updates

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Processes devices in configurable batches (default: 50 devices per batch)
- **Delay Between Batches**: Automatically pauses between batches (default: 10 seconds) to avoid overwhelming the Graph API
- **Throttling Detection**: Automatically detects when the Graph API returns throttling responses (HTTP 429)
- **Retry Logic**: Implements exponential backoff retry logic when throttled
- **Respect for Retry-After**: Honors the Retry-After header when provided by the Graph API

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalDevices | Total number of Autopilot devices processed |
| UpdatedCount | Number of devices that had their group tags updated |
| NoChangeCount | Number of devices where the group tag already matched the device category |
| NoCategoryCount | Number of devices skipped because the corresponding Intune device had no category assigned or was 'Unknown' |
| NoMatchCount | Number of devices skipped because no matching Intune device was found |
| ErrorCount | Number of devices that encountered errors during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| BatchesProcessed | Number of batches processed |

## When to Use This Script
This script is particularly useful:

1. **After Device Category Synchronization**: Run this script after the Update-IntuneDeviceCategories.ps1 script to ensure consistent categorization.
2. **During Bulk Device Operations**: After enrolling new devices or reassigning devices to different departments/categories.
3. **Before Autopilot Redeployment**: Ensure group tags are updated before devices go through Autopilot again.
4. **For Governance and Audit**: Periodically run to ensure consistency between provisioning configuration and management configuration.
5. **Large Environment Optimization**: Particularly beneficial in large environments where most devices don't need updates.

## Scheduling Recommendations
- Schedule this script to run after the Update-IntuneDeviceCategories.ps1 script (with enough delay to ensure the device category updates have been processed)
- Consider scheduling it to run weekly or monthly depending on your device enrollment and reassignment patterns
- Use the WhatIf parameter for the first few runs to validate behavior in your environment
- **Performance Note**: The script now runs significantly faster, making more frequent scheduling feasible

## Customization Options
- Adjust batch size and delay based on the size of your environment and API throttling limits
- Add exclude logic for specific device types or situations
- Extend to synchronize other attributes between Autopilot and Intune devices
- Modify the duplicate device resolution logic if different criteria are needed

## Notes and Best Practices
- **Prerequisite Workflow**: This script assumes device categories are already correctly assigned in Intune, either manually or using the Update-IntuneDeviceCategories.ps1 script
- **Serial Number Matching**: Devices are matched between Autopilot and Intune using serial numbers, which must be correctly reported by both services
- **Duplicate Handling**: When multiple Intune devices have the same serial number, the script automatically selects the one with the most recent sync time
- **First Run Guidance**: Use the WhatIf switch on your first run to see what changes would be made without actually updating any devices
- **Performance Optimization**: The script now pre-filters devices, making it much more efficient in environments where most devices don't need updates
- **API Throttling**: For very large environments (thousands of devices), the performance improvements reduce the likelihood of hitting API limits
- **Integration with existing workflows**: This script complements the Device Category Sync solution and should be part of your overall device lifecycle management strategy

## Version History
- **v1.2.1**: Major performance improvements with pre-filtering, optimized device lookup, and enhanced duplicate handling
- **v1.2**: Enhanced retry logic and throttling handling
- **v1.1**: Added batch processing and comprehensive logging
- **v1.0**: Initial release