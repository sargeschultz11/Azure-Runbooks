# Sync-IntuneDevices

## Overview
This Azure Automation runbook script initiates a sync command for all Microsoft Intune managed devices. It connects to Microsoft Graph API using a System-Assigned Managed Identity, retrieves all managed devices, and triggers a sync operation for each device in configurable batches with throttling protection.

## Purpose
The primary purpose of this solution is to ensure device compliance and policy consistency by:
- Synchronizing all managed devices in your Intune environment on-demand
- Processing devices in batches to avoid API throttling in large environments
- Providing detailed logging and execution statistics
- Handling failures gracefully with retry logic and comprehensive reporting

This automation helps organizations maintain better device compliance by ensuring all devices receive the latest policies, configurations, and security settings without requiring manual intervention or user action. It's particularly useful after making bulk policy changes or when you need to quickly propagate security updates across your device fleet.

## Prerequisites
- An Azure Automation account with System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `DeviceManagementManagedDevices.ReadWrite.All`
  - `DeviceManagementManagedDevices.PrivilegedOperations.All`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| BatchSize | Int | No | Number of devices to process in each batch. Default is 50. |
| BatchDelaySeconds | Int | No | Number of seconds to wait between processing batches. Default is 10. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |

## Setting Up Managed Identity Permissions
You can use the included `Add-GraphPermissions.ps1` script to assign the necessary Microsoft Graph API permissions to your Automation Account's System-Assigned Managed Identity:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Get the Object ID of the Managed Identity from the Azure Portal
3. Update the `$AutomationMSI_ID` parameter in the script with your Managed Identity's Object ID
4. Run the script from a PowerShell session with suitable administrative permissions

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity.
2. **Device Retrieval**: Gets all managed devices from Intune.
3. **Batch Processing**: Divides devices into batches of the specified size.
4. **Processing Loop**: For each batch:
   - Processes each device in the batch by sending a sync command
   - Logs success or failure for each device
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
| SuccessCount | Number of devices successfully synced |
| ErrorCount | Number of devices that encountered errors during processing |
| DurationMinutes | Total run time in minutes |
| BatchesProcessed | Number of batches processed |

## When to Use This Runbook
This runbook is particularly useful in the following scenarios:

1. **After Policy Changes**: Run after deploying new policies or configurations to ensure all devices receive them promptly
2. **Security Incidents**: During security incidents when you need to ensure all devices quickly receive updated security policies
3. **Compliance Initiatives**: As part of compliance initiatives when you need to verify all devices have the latest settings
4. **Troubleshooting**: When devices are reported as out of sync or not receiving policies
5. **Scheduled Maintenance**: As part of a regular maintenance routine to keep all devices in sync

## Scheduling Recommendations
Consider scheduling this runbook to run:
- Weekly during off-hours to ensure regular synchronization
- After major policy deployments
- In response to compliance alerts or security incidents

For very large environments, schedule it during off-peak hours to minimize the impact of increased device check-ins.

## Notes and Best Practices
- For large environments (thousands of devices), consider adjusting the BatchSize and BatchDelaySeconds parameters to optimize throughput while avoiding throttling
- Running this script will cause all devices to check in with Intune, which may briefly increase server load
- The sync operation is queued for each device but actual synchronization depends on device connectivity and availability
- Mobile devices may not sync immediately if they are in power-saving mode or have poor connectivity
- Review the runbook logs after execution to identify any devices that consistently fail to sync, as this may indicate underlying issues
- The script logs detailed information about each device sync attempt for troubleshooting
- Consider pairing this runbook with compliance reporting solutions to verify the effectiveness of the sync operations

## Troubleshooting
If you encounter issues with the runbook:
1. Check that the Managed Identity has all required permissions
2. Verify the Az.Accounts module is imported into your Automation account
3. Check the runbook logs for specific error messages
4. For persistent sync failures on specific devices, verify device connectivity and ensure the device is still actively managed

## Security Considerations
- The `DeviceManagementManagedDevices.PrivilegedOperations.All` permission is a highly privileged permission - ensure your Automation Account is properly secured
- Consider implementing additional access controls for scheduling or triggering this runbook
- Review the execution logs regularly to monitor for unexpected behavior

## Integration with Other Solutions
This runbook can be effectively paired with:
- The Device Category Sync runbook from this repository
- Compliance reporting solutions
- Security incident response workflows