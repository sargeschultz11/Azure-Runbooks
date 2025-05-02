# Alert-IntuneAppleTokenMonitor

## Overview
This Azure Automation runbook monitors the expiration dates of various Apple tokens and certificates in Microsoft Intune and sends proactive alerts through Microsoft Teams and/or email. It helps prevent service disruptions by ensuring that critical Apple integration components are renewed before they expire.

## Purpose
The primary purpose of this solution is to maintain uninterrupted Apple device management in Intune by:
- Monitoring expiration dates for Apple Push Notification certificates, VPP tokens, and DEP tokens
- Sending proactive notifications when tokens are approaching expiration
- Providing detailed information about each token to facilitate renewal
- Generating a comprehensive report of all tokens and their status
- Supporting multiple notification methods (Teams, email) for increased visibility

This automation helps organizations prevent service disruptions that can occur when Apple certificates and tokens expire unexpectedly. It reduces the administrative overhead of manually tracking expiration dates and provides a central monitoring solution for all Apple-related services in Intune.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementServiceConfig.Read.All`
  - `DeviceManagementConfiguration.Read.All`
  - `DeviceManagementApps.Read.All`
  - `Organization.Read.All`
  - `Mail.Send` (required if using email notifications)
- The Az.Accounts PowerShell module must be imported into your Azure Automation account
- A Microsoft Teams webhook for receiving notifications (if using Teams notifications)

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| WarningThresholdDays | Int | No | The number of days before expiration to start sending warning notifications. Default is 30 days. |
| TeamsWebhookUrl | String | No | Microsoft Teams webhook URL for sending notifications about token status. If not provided, Teams notifications will not be sent. |
| SendEmailNotification | Switch | No | When specified, email notifications will be sent. If specified, EmailSender and EmailRecipients must also be provided. |
| EmailSender | String | No | The email address that will be used as the sender for email notifications. Required if SendEmailNotification is specified. |
| EmailRecipients | String | No | A comma-separated list of email addresses that will receive the notifications. Required if SendEmailNotification is specified. |
| WhatIf | Switch | No | If specified, shows what would be done but doesn't actually send notifications. |

## Setting Up Teams Webhook
Microsoft is transitioning from Office 365 Connectors to the Workflows app in Teams. To create a webhook URL for Microsoft Teams notifications:

1. In Teams, navigate to the channel where you want to receive notifications
2. You'll need to use the Workflows app to create an incoming webhook
3. Use the "When a Teams webhook request is received" trigger in Workflows
4. Configure the workflow to post to your desired channel
5. Copy the generated HTTP POST URL from the workflow
6. Use this URL for the `TeamsWebhookUrl` parameter when running the runbook

For detailed step-by-step instructions on setting up webhooks with the Workflows app, refer to Microsoft's documentation: [Create incoming webhooks with Workflows for Microsoft Teams](https://support.microsoft.com/en-us/office/create-incoming-webhooks-with-workflows-for-microsoft-teams-8ae491c7-0394-4861-ba59-055e33f75498)

## Setting Up Email Notifications
To use email notifications:

1. Ensure your Managed Identity has the `Mail.Send` permission 
2. Specify a valid sender email address that exists in your Microsoft 365 tenant
3. Provide a comma-separated list of recipient email addresses 
4. Use the SendEmailNotification switch when running the runbook

The email notification includes:
- A formatted HTML table of all tokens requiring attention
- Color-coded status indicators (Warning, Critical, Expired)
- Renewal process instructions for each token type
- Organization branding based on your tenant information

## Setting Up Managed Identity Permissions
You can use the standard `Add-GraphPermissions.ps1` script (available in other runbook folders in this repository) to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

Make sure to use these permission IDs in the script:
```powershell
$GraphPermissionsList = @(
    @{Name = "DeviceManagementServiceConfig.Read.All"; Id = "06a5fe6d-c49d-46a7-b082-56b1b14103c7"},
    @{Name = "DeviceManagementConfiguration.Read.All"; Id = "dc377aa6-52d8-4e23-b271-2a7ae04cedf3"},
    @{Name = "DeviceManagementApps.Read.All"; Id = "7a6ee1e7-141e-4cec-ae74-d9db155731ff"},
    @{Name = "Organization.Read.All"; Id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"},
    @{Name = "Mail.Send"; Id = "b633e1c5-b582-4048-a93e-9f11b44c7e96"}
)
```

## Tokens Monitored
The script monitors the following Apple tokens and certificates in Intune:

1. **Apple Push Notification Certificate (APNs)**
   - Required for iOS and macOS device management
   - Typically expires annually
   - Renewal requires a CSR from Microsoft and processing on the Apple Push Certificates Portal

2. **Volume Purchase Program (VPP) Tokens**
   - Used for distributing and managing apps purchased through Apple Business Manager
   - Typically expire annually
   - Multiple tokens may exist for different locations or departments

3. **Device Enrollment Program (DEP) Tokens**
   - Used for Automated Device Enrollment through Apple Business Manager
   - Typically expire annually
   - Multiple tokens may exist for different locations or enrollment scenarios

## Notification Formats

### Teams Notification Format
The Teams notification includes:
- Token type (APNs, VPP, or DEP)
- Token name/identifier
- Expiration date
- Days remaining until expiration
- Status (OK, Warning, Critical, or Expired)

Notifications are sent using adaptive cards for better readability and visibility.

### Email Notification Format
The email notification includes:
- A professional HTML template with organization branding
- A table of all tokens requiring attention
- Color-coded status indicators
- Expiration dates and days remaining
- Renewal guidance specific to each token type
- Timestamp of report generation

## Status Thresholds
The script categorizes tokens into the following status levels:
- **OK**: More than 30 days (or custom threshold) until expiration
- **Warning**: Less than 30 days (or custom threshold) but more than 7 days until expiration
- **Critical**: 7 days or less until expiration
- **Expired**: Already expired

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalTokensChecked | Total number of tokens and certificates checked |
| HealthyTokens | Number of tokens with "OK" status |
| WarningTokens | Number of tokens with "Warning" status |
| CriticalTokens | Number of tokens with "Critical" or "Expired" status |
| ExpiringTokens | Number of tokens that require notification |
| ExecutionTimeMinutes | Total run time in minutes |
| TokenCollection | Array of all tokens with their details |
| ExpiringTokenDetails | Array of tokens that are expiring soon with their details |
| NotificationMethod | The methods used to send notifications (Teams, Email, or None) |
| EmailNotificationSent | Boolean indicating whether email notification was sent successfully |
| TeamsNotificationSent | Boolean indicating whether Teams notification was sent successfully |

## Scheduling Recommendations
- For most organizations, running this check weekly should be sufficient
- Consider running it more frequently (e.g., daily) when tokens are approaching expiration
- Schedule the runbook to run during business hours so IT staff can take immediate action when notifications are received

## Renewal Process Preparation
When a token is approaching expiration, you will need to:

1. **For APNs Certificate**:
   - Generate a new CSR in the Intune admin center
   - Use the CSR to renew the certificate in the Apple Push Certificates Portal
   - Upload the renewed certificate back to Intune

2. **For VPP Tokens**:
   - Download a new token from Apple Business Manager
   - Upload the new token to Intune
   - Ensure the new token is associated with the correct location/group

3. **For DEP Tokens**:
   - Generate a new server token in Apple Business Manager
   - Upload the new token to Intune
   - Verify that device assignments are preserved

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has all required permissions
2. Verify the Az.Accounts module is imported into your Automation account
3. Ensure the Teams webhook URL is valid and the associated connector is still active
4. If using email notifications, verify that the sender email address exists in your tenant
5. Review the runbook logs for specific error messages
6. Test with the WhatIf parameter to validate logic without sending notifications

## Notes
- All tokens and certificates are checked regardless of their expiration status
- The script uses the Microsoft Graph beta endpoint to access some Intune management features
- Notifications are only sent for tokens that are approaching expiration or have expired
- The script handles multiple tokens of each type (especially useful for VPP and DEP tokens)
- Organizations with separate Apple Business Manager instances for different regions may have multiple tokens of each type
- You can use both Teams and email notifications simultaneously for redundancy
- Email notifications require the Mail.Send permission which is considered a higher-privilege permission
