# Intune Device Sync Reminder Solution

## Overview
This Azure Automation runbook solution automatically identifies Intune-managed devices that haven't synced within a specified time threshold and sends email notifications to the primary users of those devices with instructions on how to perform a sync. The email template includes platform-specific instructions for Windows, iOS/iPadOS, Android, and macOS devices, giving users clear guidance on resolving the issue.

## Purpose
The primary purpose of this solution is to maintain a healthy and compliant device fleet by:
- Identifying devices that haven't synced with Intune within a configurable time period
- Notifying users of these devices with clear, action-oriented instructions
- Providing platform-specific guidance for all major operating systems
- Automating the notification process to reduce IT workload
- Implementing robust error handling, throttling management, and batching for large environments
- Sending nicely formatted HTML email messages from an internal mailbox, which provides a more professional experience compared to Intune's built-in notification settings that send from a Microsoft mailbox

This automation helps organizations maintain better security posture by ensuring devices regularly receive policy and security updates from Intune, reducing administrative overhead for IT staff, and empowering users to resolve sync issues themselves.

## Solution Components
The solution consists of the following components:

1. **Send-IntuneDeviceSyncReminders.ps1** - The main Azure Automation runbook script
2. **Add-GraphPermissions.ps1** - Helper script to configure necessary Graph API permissions
3. **email-template.html** - HTML email template for the notification emails to device users
4. **it-notification-template.html** - HTML email template for IT department notifications about devices without primary users

## Prerequisites
- An Azure Automation account with a System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`
  - `Directory.Read.All`
  - `User.Read.All`
  - `Mail.Send`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| DaysSinceLastSync | Int | No | The number of days to use as a threshold for determining "stale" devices that need to sync. Default is 7 days. |
| EmailSender | String | Yes | The email address that will appear as the sender of the notification emails. |
| ExcludedDeviceCategories | String[] | No | An array of device categories to exclude from the notification process. |
| MaxEmailsPerRun | Int | No | Maximum number of emails to send in a single runbook execution. Default is 100. |
| BatchSize | Int | No | Number of devices to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |
| WhatIf | Switch | No | If specified, shows what would be done but doesn't actually send emails. |
| TestEmailAddress | String | No | If specified, all emails will be sent to this address instead of the actual device users. Use this for testing purposes. |
| LogoUrl | String | No | URL to the company logo to use in the email template. Default is a placeholder. |
| SupportEmail | String | No | The email address users should contact for support. Default is "it@example.com". |
| SupportPhone | String | No | The phone number users should call for support. Default is "(555) 123-4567". |
| ITDepartmentEmail | String | No | The email address to send notifications about devices without primary users. If not specified, no IT notification email will be sent. |

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Replace the placeholder in the script with your Managed Identity's Object ID:
   ```powershell
   $AutomationMSI_ID = "<REPLACE_WITH_YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
   ```
4. Run the script from a PowerShell session with the required permissions:
   - AppRoleAssignment.ReadWrite.All
   - Application.Read.All

**IMPORTANT**: The `Add-GraphPermissions.ps1` script must be run on a local machine using an account with global admin permissions. It does not work when run as a runbook.

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity.
2. **Device Retrieval**: Gets all managed devices (Windows, iOS/iPadOS, Android, macOS/Linux) from Intune.
3. **Filter Outdated Devices**: Identifies devices that haven't synced since the specified threshold date.
4. **Batch Processing**: Divides devices into batches of the specified size for efficient processing.
5. **User Lookup**: For each outdated device, retrieves the primary user if one exists.
6. **Email Notification**: Sends platform-specific sync instruction emails to users.
7. **Output Generation**: Produces a summary object with detailed execution statistics.
8. **IT Department Notification**: If configured, sends a separate email to IT staff about devices without primary users.

## Email Template Features
The email sent to users includes:
- A professional HTML template with company logo
- Clear identification of the device that needs attention
- The last sync date for context
- Platform-specific instructions with separate tabs for:
  - iOS/iPadOS
  - Android
  - Windows
  - macOS
- Instructions for syncing the device and checking for updates
- Support contact information
- "Why this matters" explanation for users

The IT department notification email includes:
- Summary of devices requiring attention
- Detailed table with device name, OS, model, serial number, and last sync time
- Suggested actions for IT administrators
- Timestamp and context information

Unlike Intune's built-in notification system that sends emails from Microsoft addresses, this solution allows you to send beautiful, branded HTML emails from your organization's own email addresses. This creates a more professional, trustworthy communication that aligns with your organization's identity and increases the likelihood of user action.

## Throttling and Error Handling
The script includes robust throttling detection and error handling mechanisms:
- **Batch Processing**: Processes devices in configurable batches (default: 50 devices per batch)
- **Delay Between Batches**: Automatically pauses between batches (default: 10 seconds)
- **Throttling Detection**: Automatically detects when the Graph API returns throttling responses (HTTP 429)
- **Exponential Backoff**: Implements exponential backoff retry logic when throttled
- **Multiple Token Acquisition Methods**: Uses fallback methods to acquire authentication tokens if primary methods fail
- **Comprehensive Logging**: Detailed logging of all operations with appropriate severity levels

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalDevices | Total number of devices processed |
| OutdatedDevices | Number of devices that haven't synced since the threshold date |
| EmailsSent | Number of emails sent successfully |
| RecentlySyncedCount | Number of devices that have synced recently (within threshold) |
| SkippedNoUserCount | Number of devices skipped because they had no primary user |
| SkippedCategoryCount | Number of devices skipped due to excluded categories |
| SkippedMaxEmailsCount | Number of devices skipped due to reaching the maximum email limit |
| ErrorCount | Number of errors encountered during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |
| DurationMinutes | Total run time in minutes |
| SyncThresholdDate | The date used as the sync threshold |
| ITNotificationSent | Boolean indicating if an email was sent to IT about devices without primary users |
| DevicesWithoutUsersCount | Number of devices that have no primary user assigned |
| *OSName*Devices | Total number of devices processed for each OS type |
| *OSName*EmailsSent | Number of emails sent for each OS type |

## Scheduling and Testing
For optimal user experience and IT operations:

### Initial Testing
- First run the runbook with `WhatIf = $true` to validate logic without sending emails
- Use `TestEmailAddress` parameter to review email formatting by sending all notifications to a test account
- Review the runbook logs after test runs to ensure proper execution

### Production Scheduling
- Schedule the runbook to run weekly (e.g., every Monday morning)
- Consider running it during business hours so users can take immediate action
- Start with a shorter sync threshold (e.g., 14 days) initially, then adjust as needed
- Monitor the first few production runs closely to ensure proper operation

## Configuration Options

When setting up the runbook in Azure Automation, you'll need to provide parameter values. Here are some common configurations to consider:

### Common Parameters to Set
- **DaysSinceLastSync**: Set to your preferred threshold (e.g., 7-14 days)
- **EmailSender**: Your organization's notification email address
- **SupportEmail**: Your IT helpdesk email
- **SupportPhone**: Your IT helpdesk phone number
- **LogoUrl**: URL to your company logo for email branding

### Testing Options
- Use **TestEmailAddress** to direct all emails to a test account
- Enable **WhatIf** to simulate the process without sending actual emails

### For Large Environments
- Adjust **BatchSize** and **BatchDelaySeconds** to manage API throttling
- Set **MaxEmailsPerRun** to control the volume of notifications
- Use **ExcludedDeviceCategories** to skip certain device types
- Configure **ITDepartmentEmail** to receive reports about devices without users

### Monitoring Parameters
- **MaxRetries** and **InitialBackoffSeconds** control how the script handles API throttling
- Review the runbook job output to see detailed statistics about the execution

## Tips and Best Practices
- For large organizations, consider adjusting the BatchSize and BatchDelaySeconds parameters to balance throughput with API throttling limits
- The ExcludedDeviceCategories parameter is useful for devices that don't require regular syncing (e.g., kiosks, lab equipment)
- Implement a gradual rollout by starting with a small subset of devices and expanding over time
- Monitor the runbook logs for any issues with token acquisition or API throttling
- Use the ITDepartmentEmail parameter to ensure IT staff are aware of devices without primary users that need attention

### Excluding Device Groups
While the script currently supports excluding devices by category using the ExcludedDeviceCategories parameter, you might want to extend the functionality to exclude specific device groups. This would require modifying the script to:

1. Add a new parameter for excluded groups
2. Query the Graph API for each device's group memberships 
3. Skip processing devices that belong to excluded groups

This enhancement would be useful for:
- Excluding test or development devices
- Preventing notifications for special-purpose devices
- Creating exceptions for executive or sensitive devices
- Managing classroom or shared devices differently

## Customization
The email templates can be customized to match your organization's branding requirements:

1. **User notification template** (`email-template.html`):
   - `$LogoUrl` - URL to your company logo
   - `$Username` - The user's first name
   - `$DeviceName` - The name of the device that needs to sync
   - `$LastSyncTime` - The date and time of the last successful sync
   - `$SupportEmail` - Your support email address
   - `$SupportPhone` - Your support phone number

   The template includes HTML comments that allow you to easily enable or disable specific device sections:
   - `<!-- WINDOWS_INSTRUCTIONS_START -->` and `<!-- WINDOWS_INSTRUCTIONS_END -->` - Windows device instructions
   - `<!-- IOS_INSTRUCTIONS_START -->` and `<!-- IOS_INSTRUCTIONS_END -->` - iOS/iPadOS device instructions
   - `<!-- ANDROID_INSTRUCTIONS_START -->` and `<!-- ANDROID_INSTRUCTIONS_END -->` - Android device instructions
   - `<!-- MACOS_INSTRUCTIONS_START -->` and `<!-- MACOS_INSTRUCTIONS_END -->` - macOS device instructions
   - `<!-- UPDATES_SECTION_START -->` and `<!-- UPDATES_SECTION_END -->` - The entire updates section

   To disable a section, you can either:
   - Delete the content between the comment tags
   - Comment out the entire section with additional HTML comments by adding `<!--` to the beginning of a section & `-->` at the end

   To enable a section, you can:
   - Remove the `<!--` & `-->` at the beginning and end of a disabled section 

   This allows you to tailor the email template to include only instructions relevant to the device types in your environment.

2. **IT department notification template** (`it-notification-template.html`):
   - `$LogoUrl` - URL to your company logo
   - `$DeviceCount` - Number of devices without primary users
   - `$DaysSinceLastSync` - Threshold days set in the script
   - `$DeviceTableRows` - Dynamically generated HTML table rows
   - `$CurrentDate` - Current date/time when the email is sent

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has all required permissions
2. Verify the Az.Accounts module is imported into your Automation account
3. Review the runbook logs for specific error messages
4. Test with the WhatIf parameter to validate logic without sending emails
5. Use the TestEmailAddress parameter to send all notifications to a test account
6. If seeing issues with permission assignment, ensure you're running the Add-GraphPermissions.ps1 script from a local machine with global admin permissions

## Notes
- This solution uses a Managed Identity for authentication instead of App Registration credentials, enhancing security
- The script dynamically builds device OS statistics to provide insights into your device fleet
- Detailed logging enables troubleshooting and verification of successful operation
- The IT notification feature helps administrators identify and fix devices that don't have primary users assigned