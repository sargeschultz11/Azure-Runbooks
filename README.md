# Azure-Runbooks

A collection of Azure Automation runbooks for Microsoft 365 and Intune management.

## Overview

This repository contains PowerShell scripts designed to be used as Azure Automation runbooks for automating various Microsoft 365 and Intune management tasks. These scripts help streamline administrative processes, maintain consistency across your environment, and reduce manual overhead.

## Repository Structure

The repository is organized into folders, with each folder containing a specific runbook solution:

```
Azure-Runbooks/
├── DeviceCategorySync/             # Sync device categories with user departments
├── Report-DiscoveredApps/          # Generate reports of discovered applications
├── Report-IntuneDeviceCompliance/  # Generate device compliance reports
├── Report-DevicesWithApp/          # Find devices with specific applications
├── Alert-DeviceSyncReminder/       # Send reminders for devices needing sync
├── Update-AutopilotDeviceGroupTags/ # Sync Autopilot group tags with Intune categories
└── [future runbooks]/              # More solutions will be added
```

Each runbook folder contains:
- The main PowerShell script (`.ps1`)
- A helper script for setting up permissions (`Add-GraphPermissions.ps1`)
- Detailed documentation (`README.md`)

## Authentication

All runbooks in this repository are designed to use Azure Automation's System-Assigned Managed Identity for authentication, which is the recommended approach for Azure Automation. Each folder includes an `Add-GraphPermissions.ps1` script that helps assign the necessary Microsoft Graph API permissions to your Automation Account's Managed Identity.

## Getting Started

Each runbook includes detailed documentation for implementation and usage. In general, to use these runbooks:

1. Import the script into your Azure Automation account
2. Enable System-Assigned Managed Identity on your Automation account
3. Use the included `Add-GraphPermissions.ps1` script to assign necessary Graph API permissions
4. Configure any required parameters specific to your environment
5. Create a schedule or link to a webhook for execution
6. Review logs and output after execution

## Requirements

- Azure Automation account
- Appropriate Microsoft Graph API permissions (varies by runbook)
- Required PowerShell modules (specified in each runbook's documentation)
- Pre-configured Azure Automation variables (if specified in a runbook)

## Available Solutions

### Device Management
- **Device Category Sync**: Automatically update Intune device categories based on the primary user's department.
- **Autopilot Group Tag Sync**: Synchronize Windows Autopilot device group tags with their corresponding Intune device categories.
- **Device Sync Reminder**: Identify devices that haven't synced in a specified period and send email notifications to their primary users.

### Reporting
- **Discovered Apps Report**: Generate comprehensive reports of all applications discovered across your managed devices.
- **Device Compliance Report**: Create detailed reports on device compliance status.
- **Devices with Specific App Report**: Identify all devices with a specific application installed.

## Contributing

Feel free to use these scripts as a starting point for your own automation needs. Contributions, improvements, and suggestions are welcome!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.