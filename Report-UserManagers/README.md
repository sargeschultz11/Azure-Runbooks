# Export-UserManagers

## Overview
This Azure Automation runbook script generates a report of **licensed internal users (non-guests)** in your Microsoft 365 tenant, along with their manager information. The report is exported as an Excel spreadsheet and automatically uploaded to a specified SharePoint document library.

## Purpose
The purpose of this runbook is to:
- Retrieve only internal, licensed Microsoft 365 users (excluding guest and unlicensed accounts)
- Collect user and manager information
- Generate a clean Excel report listing users and their assigned managers
- Upload the report to a specified SharePoint site and library
- Automate the process securely using Managed Identity authentication

## Prerequisites
- An Azure Automation account with **System-Assigned Managed Identity** enabled
- The Managed Identity must have the following Microsoft Graph API permissions:
  - `User.Read.All`
  - `Directory.Read.All`
  - `Sites.ReadWrite.All`
- The following modules must be imported into your Automation account:
  - `Az.Accounts`
  - `ImportExcel`
- You must know the SharePoint **Site ID** and **Drive ID** where the Excel file will be uploaded

## Parameters

| Parameter     | Type   | Required | Description                                                        |
|---------------|--------|----------|--------------------------------------------------------------------|
| SiteId        | String | Yes      | The SharePoint site ID where the report will be uploaded           |
| DriveId       | String | Yes      | The Drive ID of the SharePoint document library to upload into     |

## Setting Up Managed Identity Permissions
Use the included `Add-GraphPermissions.ps1` script to grant required permissions:

1. Enable System-Assigned Managed Identity for your Azure Automation account
2. Find the Object ID of the identity in Azure
3. Run the permission script with administrative rights:
   ```powershell
   .\Add-GraphPermissions.ps1 -AutomationMSI_ID "<YOUR_OBJECT_ID>"
   ```

## Report Contents

### Excel Sheet: Users and Managers
The exported Excel file includes the following columns:
- Display Name
- User Principal Name
- Email
- Manager Name
- Manager Email

Only users with assigned licenses and a `userType` of `Member` (i.e., non-guests) are included.

## Setup Instructions

### 1. Enable Managed Identity
- Navigate to your Automation Account → Identity → System-assigned → Enable

### 2. Import Required Modules
- Go to **Modules > Browse Gallery**
  - Import `Az.Accounts`
  - Import `ImportExcel`

### 3. Get SharePoint Site and Drive IDs
- Use Microsoft Graph Explorer or PowerShell to obtain:
  - `SiteId`: ID of the SharePoint site
  - `DriveId`: ID of the SharePoint document library

### 4. Import and Configure the Runbook
1. Go to **Runbooks > Import a Runbook**
2. Upload `Export-UserManagers.ps1`
3. Publish it

### 5. Schedule the Runbook
1. Go to the runbook > Schedules > Add schedule
2. Link it to your automation schedule and provide the required parameters

## Execution Flow
1. **Authenticate** using Managed Identity with Microsoft Graph
2. **Fetch all users** from Graph and filter only internal, licensed users
3. **Get each user’s manager** using `$expand=manager`
4. **Create an Excel file** containing all relevant data
5. **Upload the file** to your SharePoint library
6. **Clean up temp files** and output a success message

## Notes
- Guest users and unlicensed users are excluded from the report
- Manager lookups use `$expand=manager`, which may fail silently for users without managers
- No summary sheet is currently generated — just raw user+manager data
- The output filename includes a timestamp for uniqueness
- The script does not currently implement retry/backoff for throttling — consider adding if needed for large orgs

## Output Example
On successful run, you'll get:
- A timestamped Excel file uploaded to SharePoint
- A success message: `Report uploaded to SharePoint successfully.`
