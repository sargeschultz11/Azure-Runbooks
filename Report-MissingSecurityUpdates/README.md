# Get-MissingSecurityUpdatesReport

## Overview
This Azure Automation runbook script monitors Windows devices in your environment that are missing multiple security updates, generating a comprehensive report that's automatically exported to Excel and uploaded to SharePoint. It leverages Log Analytics data to identify vulnerable devices, providing IT administrators with actionable information to maintain security compliance.

## Purpose
The primary purpose of this runbook is to:
- Identify Windows devices that are missing multiple security updates
- Collect detailed information about each device and its security status
- Generate a structured Excel report with comprehensive statistics
- Upload the report to a specified SharePoint location for easy access
- Send optional notifications via Microsoft Teams with a link to the report

This automation helps organizations maintain better security posture by enabling proactive identification and remediation of devices with missing security updates, reducing administrative overhead for IT security teams, and providing better visibility into the organization's patch compliance status.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `Sites.ReadWrite.All` (for SharePoint upload functionality)
- The Managed Identity must have access to the specified Log Analytics workspace
- The following PowerShell modules must be imported into your Azure Automation account:
  - `ImportExcel`
  - `Az.Accounts`
  - `Az.OperationalInsights`
- WUfB reporting to a Log Analytics workspace enabled 
- A SharePoint site ID and document library drive ID where the report will be uploaded

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| WorkspaceId | String | Yes | The Log Analytics Workspace ID to query for missing security updates data. |
| SharePointSiteId | String | Yes | The ID of the SharePoint site where the report will be uploaded. |
| SharePointDriveId | String | Yes | The ID of the document library drive where the report will be uploaded. |
| FolderPath | String | No | The folder path within the document library for upload. Default is root. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying. Default is 5. |
| TeamsWebhookUrl | String | No | Optional. Microsoft Teams webhook URL for sending notifications about the report. |

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script (common across all runbooks in this repository) to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

You'll also need to assign the appropriate permissions to access your Log Analytics workspace:
1. Navigate to your Log Analytics workspace in the Azure Portal
2. Go to Access Control (IAM)
3. Add a role assignment for your Automation Account's Managed Identity:
   - Role: Log Analytics Reader (or higher)
   - Assign access to: System assigned managed identity
   - Select your Automation Account

## Report Contents
The generated Excel report includes:

### "Missing Security Updates" Tab
A table with detailed information about each device missing security updates:
- Device Name
- Azure AD Device ID
- Alert ID
- Alert Generated Time
- Description of the security issue
- Recommendation for remediation

### "Summary" Tab
- Report metadata (generation date, system info)
- Total number of devices missing security updates
- Alert age distribution breakdown (showing how long devices have been missing updates)
  - Less than 24 hours
  - 1-3 days
  - 4-7 days
  - 8-14 days
  - 15-30 days
  - More than 30 days

## Log Analytics Query
The runbook uses the following Log Analytics query to retrieve devices missing multiple security updates:

```kql
UCDeviceAlert
| where AlertSubtype == "MultipleSecurityUpdatesMissing"
| where AlertStatus == "Active"
| summarize arg_max(TimeGenerated, *) by DeviceName
| project DeviceName, AzureADDeviceId, AlertId, TimeGenerated, Description, Recommendation
| order by TimeGenerated desc
```

This query looks for active alerts of type "MultipleSecurityUpdatesMissing" and retrieves the most recent alert for each device, ensuring the report contains only the latest information.

## Teams Notification
If a Teams webhook URL is provided, the runbook will send a notification with:
- A summary of the findings (number of devices missing updates)
- A direct link to the SharePoint report

The adaptive card used for Teams notifications includes:
- A clear, attention-grabbing header
- The number of devices missing security updates
- The report generation timestamp
- A prominent link to the full report

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph and Log Analytics using the System-Assigned Managed Identity.
2. **Data Retrieval**: Executes a KQL query against Log Analytics to find devices missing security updates.
3. **Data Processing**: Analyzes the alert data, including age distribution of alerts.
4. **Excel Report Generation**: Creates a detailed Excel report with device data and summary statistics.
5. **SharePoint Upload**: Uploads the report to the specified SharePoint location.
6. **Teams Notification**: Optionally sends a notification card to Teams with a link to the report.

## Scheduling Recommendations
- Run daily or weekly, depending on your security requirements and patching schedule
- Schedule the runbook to run after your regular patch maintenance window to identify devices that still have outstanding updates
- Consider running more frequently (e.g., daily) during critical security events or when addressing zero-day vulnerabilities

## Setup Instructions

### 1. Configure System-Assigned Managed Identity
1. In your Azure Automation account, navigate to Identity
2. Enable the System-assigned identity
3. Copy the Object ID of the Managed Identity for later use

### 2. Assign Microsoft Graph API Permissions
Run the included `Add-GraphPermissions.ps1` script to assign the necessary Graph API permissions.

### 3. Assign Log Analytics Permissions
Ensure your Managed Identity has appropriate access to your Log Analytics workspace.

### 4. Get SharePoint Site and Drive IDs
1. You'll need the SharePoint site ID and document library drive ID where reports will be uploaded
2. These can be obtained using Graph Explorer or PowerShell

### 5. Set Up Azure Automation Account
1. Create or use an existing Azure Automation account with System-Assigned Managed Identity enabled
2. Import the required modules:
   - ImportExcel module
   - Az.Accounts module
   - Az.OperationalInsights module

### 6. Import the Runbook
1. In the Automation account, go to Runbooks > Import a runbook
2. Upload the Get-MissingSecurityUpdatesReport.ps1 file
3. Set the runbook type to PowerShell

### 7. Schedule the Runbook
1. Navigate to the runbook > Schedules > Add a schedule
2. Create a new schedule or link to an existing one
3. Configure the required parameters:
   - WorkspaceId
   - SharePointSiteId
   - SharePointDriveId
   - (Optional) FolderPath
   - (Optional) TeamsWebhookUrl

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| DevicesCount | Total number of devices missing security updates |
| ReportName | Name of the generated report file |
| ExecutionTimeMinutes | Total execution time in minutes |
| Timestamp | Report generation timestamp |
| ReportUrl | SharePoint URL to the uploaded report |
| NotificationSent | Boolean indicating whether Teams notification was sent successfully |
| AgeDistribution | Array of objects containing alert age distribution statistics |

## Integration with Other Solutions
This runbook can be paired with other security automation tools:
- Use in conjunction with the Get-IntuneDeviceComplianceReport.ps1 runbook for comprehensive security compliance reporting
- Trigger remediation runbooks or workflows based on the findings
- Create Power Automate flows that use the report data to automatically assign tickets to IT staff

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has all required permissions
2. Verify all required modules are imported into your Automation account
3. Ensure the Log Analytics query returns results when run manually
4. Validate that your SharePoint site and drive IDs are correct
5. Check the runbook logs for specific error messages

## Error Handling
The runbook includes comprehensive error handling:
- Authentication failures are captured and reported
- API throttling is handled with exponential backoff
- Log Analytics query failures are properly managed
- SharePoint upload issues are handled with detailed error reporting
- All temporary files are cleaned up even when errors occur

## Advanced Customization
The runbook can be customized for specific needs:
- Modify the Log Analytics query to target different alert types
- Adjust the Excel report format to include additional information
- Customize the Teams notification card to include organization-specific branding
- Add additional data processing to group devices by department or location

## Best Practices
- Run the script with WhatIf/verbose logging initially to validate configuration
- Schedule the report to run outside of business hours to minimize API throttling
- Create a monitoring solution to ensure the runbook completes successfully
- Establish a process for IT staff to review and address the findings in the report