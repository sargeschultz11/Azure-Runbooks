# Task-SetCompanyAttribute

## Overview
This Azure Automation runbook script sets the Company attribute for all member users in your Microsoft 365 tenant. It connects to Microsoft Graph API using a System-Assigned Managed Identity, retrieves all users based on specified filtering criteria, and updates the Company attribute to a consistent value across your organization.

## Purpose
The primary purpose of this solution is to ensure data consistency in user profiles by:
- Setting a uniform Company attribute across all user accounts
- Processing users in batches to handle large tenants efficiently
- Supporting filtering options to target specific subsets of users
- Implementing throttling protection for API requests
- Providing detailed logging and reporting of the update process

This automation helps organizations maintain better directory data quality, ensure consistency in user profile information, and simplify common administrative tasks related to user attribute management.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `User.ReadWrite.All`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| CompanyName | String | Yes | The company name to set for all users. |
| BatchSize | Int | No | Number of users to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |
| WhatIf | Switch | No | If specified, shows what changes would occur without actually making any updates. |
| ExcludeGuestUsers | Bool | No | If specified, excludes guest users from processing. Default is $true. |
| ExcludeServiceAccounts | Bool | No | If specified, excludes accounts with 'service' in the display name or UPN. Default is $true. |
| ProcessUnlicensedUsers | Bool | No | If specified, processes unlicensed users. Default is $true. |

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity.
2. **User Retrieval**: Gets all users based on the specified filtering criteria.
3. **Batch Processing**: Divides users into batches of the specified size.
4. **Processing Loop**: For each batch:
   - Processes each user in the batch
   - Checks if the company attribute needs updating
   - Updates the company attribute if necessary
   - Waits for the specified delay period before processing the next batch

## Throttling and Batching
The script includes built-in throttling detection and handling:
- **Batch Processing**: Processes users in configurable batches (default: 50 users per batch)
- **Delay Between Batches**: Automatically pauses between batches (default: 10 seconds) to avoid overwhelming the Graph API
- **Throttling Detection**: Automatically detects when the Graph API returns throttling responses (HTTP 429)
- **Retry Logic**: Implements exponential backoff retry logic when throttled
- **Respect for Retry-After**: Honors the Retry-After header when provided by the Graph API

## Filtering Options
The script provides several options to filter which users will be processed:

1. **ExcludeGuestUsers** (default: $true):
   - When enabled, only processes users with userType = 'Member'
   - When disabled, processes both member and guest users

2. **ExcludeServiceAccounts** (default: $true):
   - When enabled, skips any account with "service" in the displayName or userPrincipalName
   - Helps prevent modifying automated service accounts

3. **ProcessUnlicensedUsers** (default: $true):
   - When enabled, processes all users regardless of license status
   - When disabled, only processes users with at least one assigned license

These filtering options can be combined to precisely target the desired user accounts.

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalUsers | Total number of users processed |
| UpdatedCount | Number of users that had their company attribute updated |
| NoChangeCount | Number of users where the company attribute already matched |
| ErrorCount | Number of users that encountered errors during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| BatchesProcessed | Number of batches processed |
| CompanyName | The company name that was applied |

## Usage Examples

### Set Company Attribute for All Users
```powershell
$params = @{
    CompanyName = "Contoso Corporation"
}

Start-AzAutomationRunbook -Name "Set-CompanyAttribute" -Parameters $params
```

### Test Changes with WhatIf Mode
```powershell
$params = @{
    CompanyName = "Contoso Corporation"
    WhatIf = $true
}

Start-AzAutomationRunbook -Name "Set-CompanyAttribute" -Parameters $params
```

### Process All Users Including Guests
```powershell
$params = @{
    CompanyName = "Contoso Corporation"
    ExcludeGuestUsers = $false
}

Start-AzAutomationRunbook -Name "Set-CompanyAttribute" -Parameters $params
```

### Only Process Licensed Users
```powershell
$params = @{
    CompanyName = "Contoso Corporation"
    ProcessUnlicensedUsers = $false
}

Start-AzAutomationRunbook -Name "Set-CompanyAttribute" -Parameters $params
```

## Scheduling Recommendations
- Run with the `-WhatIf` parameter first to verify which users will be affected
- Schedule to run once for initial setup, then periodically (e.g., monthly) to catch new users
- Consider running after user provisioning processes to ensure new users get the correct company attribute

## Notes and Best Practices
- Always test with WhatIf mode first to ensure the expected users will be modified
- For large tenants, consider increasing the batch size and/or batch delay to avoid throttling
- The script only updates users where the company attribute is missing or different from the target value
- Consider combining this with other attribute standardization runbooks for comprehensive profile management
- Add the script to your user onboarding/provisioning process to ensure all new users get the attribute set correctly
- Microsoft 365 may take some time to propagate attribute changes to all services

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has the required `User.ReadWrite.All` permission
2. Verify the Az.Accounts module is imported into your Automation account
3. Review the runbook logs for specific error messages
4. Test with a smaller subset of users by adjusting the filtering parameters
5. Increase the batch delay if you encounter persistent throttling issues