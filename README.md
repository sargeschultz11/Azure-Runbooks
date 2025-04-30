# Azure-Runbooks 

<img src="https://raw.githubusercontent.com/sargeschultz11/Azure-Runbooks/dev/sync-reminder/assets/repo_logo.png" alt="Logo"/>


[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Graph API](https://img.shields.io/badge/Microsoft_Graph-v1.0-blue.svg)](https://developer.microsoft.com/en-us/graph)
[![Azure](https://img.shields.io/badge/Azure_Automation-compatible-0089D6.svg)](https://azure.microsoft.com/en-us/products/automation)
[![GitHub release](https://img.shields.io/github/release/sargeschultz11/Azure-Runbooks.svg)](https://GitHub.com/sargeschultz11/Azure-Runbooks/releases/)
[![Maintenance](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/sargeschultz11/Azure-Runbooks)
[![Made with](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)

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
├── Alert-IntuneAppleTokenMonitor/  # Monitor Apple token expirations
├── Report-UserManagers/            # Generate reports of users and their managers
├── Report-MissingSecurityUpdates/  # Report on devices missing security updates
├── Sync-IntuneDevices/             # Force sync all managed Intune devices
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
- **Intune Device Sync**: Force synchronize all devices in your Intune environment with batch processing and throttling protection.

### Security and Compliance
- **Intune Apple Token Monitor**: Monitor expiration dates of Apple Push Notification certificates, VPP tokens, and DEP tokens in Microsoft Intune and send proactive alerts through Microsoft Teams.
- **Missing Security Updates Report**: Generate reports of Windows devices missing multiple security updates from Log Analytics data and upload them to SharePoint.

### Reporting
- **Discovered Apps Report**: Generate comprehensive reports of all applications discovered across your managed devices.
- **Device Compliance Report**: Create detailed reports on device compliance status.
- **Devices with Specific App Report**: Identify all devices with a specific application installed.
- **User Managers Report**: Generate a report of all licensed internal users along with their manager information.

## Branch Management

This repository follows a simplified Git workflow:

- The `main` branch contains stable, production-ready scripts
- Development branches are created for new features or significant modifications
- Once development work is merged into `main`, the development branches are typically deleted
- For users who have cloned this repository, note that development branches may disappear after their work is completed

If you're working with a specific development branch, consider creating your own fork to ensure your work isn't affected when branches are deleted.

## Discussions

I've enabled GitHub Discussions for this repository to foster collaboration and support among users. This is the best place to:

* Ask questions about implementing specific runbooks
* Share your success stories and implementations 
* Suggest new runbook ideas or improvements
* Discuss best practices for Azure Automation
* Get help with troubleshooting

Check out the [Discussions](https://github.com/sargeschultz11/Azure-Runbooks/discussions) tab to join the conversation. We encourage you to use Discussions for general questions and community interaction, while Issues should be used for reporting bugs or specific problems with the scripts.

## Contributing

Feel free to use these scripts as a starting point for your own automation needs. Contributions, improvements, and suggestions are welcome!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.