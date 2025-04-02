# Azure-Runbooks

A collection of Azure Automation runbooks for Microsoft 365 and Intune management.

## Overview

This repository contains PowerShell scripts designed to be used as Azure Automation runbooks for automating various Microsoft 365 and Intune management tasks. These scripts help streamline administrative processes, maintain consistency across your environment, and reduce manual overhead. This repo is a fresh project so it will be updated with more runbooks in the near future.

## Available Runbooks

### Device Category Sync

**[Update-IntuneDeviceCategories.ps1](DeviceCategorySync/Update-IntuneDeviceCategories.ps1)** - Automatically updates Intune device categories to match the primary user's department.

- Processes Windows, iOS, Android, and Linux devices
- Maps device categories to user departments
- Supports batch processing to avoid API throttling
- Includes detailed logging and error handling
- See the [DeviceCategorySync README](DeviceCategorySync/README.md) for detailed documentation

## Getting Started

Each runbook includes detailed documentation for implementation and usage. In general, to use these runbooks:

1. Import the script into your Azure Automation account
2. Configure necessary variables and credentials
3. Create a schedule or link to a webhook
4. Review logs and output after execution

## Requirements

- Azure Automation account
- Appropriate Microsoft Graph API permissions
- Necessary service principal or managed identity with required permissions
- Pre-configured Azure Automation variables as specified in each runbook

## Contributing

Feel free to use these scripts as a starting point for your own automation needs. Contributions, improvements, and suggestions are welcome!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Ryan Schultz