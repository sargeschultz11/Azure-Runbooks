# Update-IntuneDeviceCategories.ps1

## Overview
This Azure Automation runbook script automatically updates the device categories of Windows and iOS devices in Microsoft Intune based on the primary user's department. It fetches all devices and updates the device category to match the department name of the assigned primary user.

## Purpose
The primary purpose of this script is to ensure consistent device categorization in Intune by:
- Identifying Windows and iOS devices with missing or mismatched categories
- Retrieving the primary user's department information
- Setting the device category to match the user's department when available

This automation helps maintain better organization within the Intune portal and can be used for device targeting, reporting, and policy assignment. It is also useful for creating dynamic groups.

## Prerequisites
- An Azure Automation account
- An Azure AD App Registration with the following:
  - Client ID
  - Client Secret
  - Proper Microsoft Graph API permissions:
    - `DeviceManagementManagedDevices.Read.All`
    - `DeviceManagementManagedDevices.ReadWrite.All`
    - `User.Read.All`
- The following variables defined in the Automation account:
  - `TenantId`: Your Azure AD tenant ID
  - `ClientId`: The App Registration's client ID
  - `ClientSecret`: The App Registration's client secret (stored as an encrypted variable)
- **IMPORTANT**: Device categories must be pre-created in Intune and must match **exactly** the department names in user account properties in Azure AD

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| TenantId | String | No | Your Azure AD tenant ID. If not provided, will be retrieved from Automation variables. |
| ClientId | String | No | The App Registration's client ID. If not provided, will be retrieved from Automation variables. |
| ClientSecret | String | No | The App Registration's client secret. If not provided, will be retrieved from Automation variables. |
| WhatIf | Switch | No | If specified, shows what changes would occur without actually making any updates. |
| OSType | String | No | Specifies which operating systems to process. Valid values are "All", "Windows", "iOS". Default is "All". |

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the provided client credentials.
2. **Device Category Retrieval**: Retrieves all device categories defined in Intune.
3. **Device Retrieval**: Gets all specified devices (Windows, iOS, or both) from Intune.
4. **Processing Loop**: For each device:
   - Checks if a device category is already assigned
   - Retrieves the primary user of the device
   - Gets the user's department information
   - If the department exists as a device category and differs from the current device category, updates the device's category

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
| WindowsDevices | Total number of Windows devices processed |
| WindowsUpdated | Number of Windows devices updated |
| WindowsMatched | Number of Windows devices already properly categorized |
| WindowsSkipped | Number of Windows devices skipped |
| WindowsErrors | Number of Windows devices with errors |
| iOSDevices | Total number of iOS devices processed |
| iOSUpdated | Number of iOS devices updated |
| iOSMatched | Number of iOS devices already properly categorized |
| iOSSkipped | Number of iOS devices skipped |
| iOSErrors | Number of iOS devices with errors |

## Logging
The script utilizes verbose logging to provide detailed information about each step:
- All log entries include timestamps and log levels (INFO, WARNING, ERROR, WHATIF)
- Write-Verbose is used for standard logging in Azure Automation
- Specific error cases are captured and logged appropriately
- OS-specific statistics are maintained separately


## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API request errors are logged with details
- Device processing errors are isolated to prevent the entire script from failing
- Summary statistics include error counts for both Windows and iOS devices

## Notes
- **CRITICAL REQUIREMENT**: The script depends on exact matching between department names in Azure AD and device category names in Intune. If these don't match exactly, the categorization will not work.
- Before running this script, ensure that all departments used in your organization have corresponding device categories created in Intune with identical naming.
- Devices without primary users or where the user has no department are skipped
- The script counts and reports cases where department names don't exist as device categories
- For devices that already have the correct category assigned, no changes are made
- If department names in Azure AD don't match device categories in Intune exactly (including case, spacing, and special characters), the script will report these as skipped devices

## Author Information
- **Author**: Ryan Schultz
- **Version**: 2.0
- **Creation Date**: 2025-04-02