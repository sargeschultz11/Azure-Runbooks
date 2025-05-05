# Report-OneDriveSharedItems

## Overview
This Azure Automation runbook scans a specified user's OneDrive for items that are being shared and generates a CSV report that is uploaded to an Azure Blob Storage container. This solution provides visibility into what files and folders users are sharing from their OneDrive accounts, which is essential for security monitoring and compliance reporting.

## Purpose
The primary purpose of this solution is to provide visibility and governance over shared OneDrive content by:
- Identifying all items (files and folders) that are being shared from a user's OneDrive
- Capturing sharing details including share type, permissions, and recipients
- Detecting potentially risky shares such as anonymous links or broad organizational access
- Creating structured CSV reports for analysis and record-keeping
- Automating the storage of reports in Azure Blob Storage for retention and auditing

This automation helps organizations maintain better security posture by ensuring appropriate data governance, detecting potential data leakage, and supporting compliance requirements around information sharing.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `Files.Read.All` (for reading OneDrive content and permissions)
  - `User.Read.All` (for looking up user information)
- The Managed Identity must also have the "Storage Blob Data Contributor" role on the target Storage Account
- The Az.Accounts and Az.Storage PowerShell modules must be imported into your Azure Automation account

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| UserPrincipalName | String | Yes | The user principal name (email address) of the OneDrive account to scan. |
| StorageAccountName | String | Yes | The name of the Azure Storage Account where the report will be stored. |
| StorageContainerName | String | Yes | The name of the Blob container in the Storage Account where the report will be stored. |
| IncludeAllFolders | Bool | No | If specified, scans all folders including subfolders. If not specified, only scans the root folder. Default is $true. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |
| WhatIf | Switch | No | If specified, shows what would be done but doesn't actually create or upload the report. |

## Report Contents
The generated CSV report includes the following information for each shared item:

- **Name**: The name of the shared file or folder
- **ItemType**: Whether the item is a file or folder
- **WebUrl**: The web URL to access the item
- **Path**: The item's path in OneDrive
- **Size**: The size of the item in bytes
- **CreatedDateTime**: When the item was created
- **LastModifiedDateTime**: When the item was last modified
- **SharedType**: The type of sharing (Direct, Link, Anonymous Link, Organization Link)
- **SharedWith**: The email addresses of users or groups the item is shared with
- **ShareLink**: The URL of any sharing links
- **Permissions**: The permissions granted (read, write, etc.)
- **ItemId**: The unique ID of the item in OneDrive
- **SharingId**: The SharePoint sharing ID associated with the item

## Share Types Tracked
The report identifies and categorizes different sharing methods:

1. **Direct Shares**: Shared directly with specific users or groups
2. **Organization Links**: Links that are accessible by anyone in your organization
3. **Anonymous Links**: Links that can be accessed without authentication
4. **Standard Links**: Links with specific access restrictions

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

You'll also need to assign storage permissions to your Managed Identity:
1. Navigate to your Azure Storage account
2. Select "Access Control (IAM)"
3. Add a role assignment with:
   - Role: "Storage Blob Data Contributor"
   - Assign access to: "Managed Identity"
   - Select your Automation Account's Managed Identity

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API and Azure Storage using the Automation Account's Managed Identity.
2. **User Identification**: Looks up the user ID from the provided UPN.
3. **OneDrive Access**: Gets the user's OneDrive drive ID.
4. **Content Scanning**: Recursively scans the OneDrive for shared items, examining sharing permissions.
5. **Report Generation**: Creates a CSV report with details of all shared items.
6. **Storage Upload**: Uploads the report to Azure Blob Storage and returns the URL.

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| UserPrincipalName | The UPN of the user whose OneDrive was scanned |
| TotalSharedItems | Total number of shared items found |
| AnonymousShares | Number of items shared with anonymous access |
| OrganizationShares | Number of items shared with organization-wide access |
| DirectShares | Number of items shared directly with specific users |
| LinkShares | Number of items shared via links with other permissions |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| ReportUrl | URL to the uploaded report in Azure Blob Storage |
| Timestamp | When the report was generated |

## Usage Examples

### Scan a Single User's OneDrive
```powershell
$params = @{
    UserPrincipalName = "user@contoso.com"
    StorageAccountName = "contosostorage"
    StorageContainerName = "onedrive-reports"
    IncludeAllFolders = $true
}

Start-AzAutomationRunbook -Name "Get-OneDriveSharedItemsReport" -Parameters $params
```

### Scan Multiple Users' OneDrives
You can create a parent runbook that calls this runbook for multiple users:

```powershell
$users = @("user1@contoso.com", "user2@contoso.com", "user3@contoso.com")
$storageAccount = "contosostorage"
$container = "onedrive-reports"

foreach ($user in $users) {
    $params = @{
        UserPrincipalName = $user
        StorageAccountName = $storageAccount
        StorageContainerName = $container
    }
    
    Start-AzAutomationRunbook -Name "Get-OneDriveSharedItemsReport" -Parameters $params
}
```

## Scheduling Recommendations
- Schedule daily or weekly scans for key users who handle sensitive information
- Run monthly scans for all users to maintain a comprehensive sharing inventory
- Consider running on-demand scans after organizational changes or security incidents
- For large organizations, stagger the scans across different times to avoid API throttling

## Integration with Other Solutions
This runbook can be effectively paired with:
- Security information and event management (SIEM) systems for alerting on risky shares
- Data Loss Prevention (DLP) solutions for comprehensive data governance
- Compliance reporting systems for audit and regulatory requirements
- Power BI dashboards for visualizing sharing patterns across the organization

## Error Handling and Throttling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API throttling is handled with exponential backoff
- Progress logging for troubleshooting
- Temporary file cleanup even when errors occur

## Notes and Best Practices
- When using with multiple users, consider the Graph API throttling limits
- The report captures a point-in-time snapshot - sharing permissions may change after the report is generated
- Consider reviewing anonymous or organization-wide shares regularly for security risks
- Set appropriate retention policies for the report storage container
- Use the WhatIf parameter first to validate functionality without creating reports