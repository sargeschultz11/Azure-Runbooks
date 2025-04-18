# Requires -Modules "Microsoft.Graph.Applications"
<#
.SYNOPSIS
    Assigns Microsoft Graph API permissions to an Azure Automation Account's System-Assigned Managed Identity.
    
.DESCRIPTION
    This script assigns the necessary Microsoft Graph API permissions to allow the 
    Get-DevicesWithAppReport.ps1 runbook to authenticate using a System-Assigned Managed Identity 
    instead of an App Registration.
    
.NOTES
    Author:         Ryan Schultz
    Version:        1.0
    Creation Date:  April 2025
    
    Required permissions to run this script:
    - AppRoleAssignment.ReadWrite.All
    - Application.Read.All
    
.PARAMETER AutomationMSI_ID
    The Object ID of your Automation Account's System-Assigned Managed Identity.
    This can be found in the Azure Portal under your Automation Account > Identity > System assigned.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$AutomationMSI_ID = "<REPLACE_WITH_YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
)

# Microsoft Graph App ID (constant)
$GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

Write-Host "Starting Graph permission assignment process..." -ForegroundColor Cyan

try {
    Write-Host "Connecting to Microsoft Graph API..."
    Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph API" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph API: $_" -ForegroundColor Red
    Write-Host "Please ensure you have the required permissions and the Microsoft.Graph.Applications module is installed." -ForegroundColor Yellow
    exit 1
}

try {
    Write-Host "Retrieving Microsoft Graph Service Principal..."
    $GraphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"
    
    if ($null -eq $GraphServicePrincipal) {
        Write-Host "Could not find Microsoft Graph Service Principal. Exiting." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found Microsoft Graph Service Principal with ID: $($GraphServicePrincipal.Id)" -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving Microsoft Graph Service Principal: $_" -ForegroundColor Red
    exit 1
}

# Define the Graph permissions required for the App Devices Report runbook
# These IDs are standard across all tenants for Microsoft Graph
$GraphPermissionsList = @(
    @{Name = "DeviceManagementManagedDevices.Read.All"; Id = "2f51be20-0bb4-4fed-bf7b-db946066c75e"},
    @{Name = "DeviceManagementManagedDevices.ReadWrite.All"; Id = "243333ab-4d21-40cb-a475-36241daa0842"},
    @{Name = "DeviceManagementServiceConfig.ReadWrite.All"; Id = "5ac13192-7ace-4fcf-b828-1a26f28068ee"}
)

Write-Host "Assigning permissions to the Managed Identity ($AutomationMSI_ID)" -ForegroundColor Cyan

foreach ($permission in $GraphPermissionsList) {
    Write-Host "Processing permission: $($permission.Name)"
    
    $existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID | 
        Where-Object { $_.AppRoleId -eq $permission.Id }
        
    if (-not $existingAssignment) {
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID `
                -PrincipalId $AutomationMSI_ID `
                -ResourceId $GraphServicePrincipal.Id `
                -AppRoleId $permission.Id
                
            Write-Host "Permission $($permission.Name) assigned successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "Error assigning permission $($permission.Name): $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Permission $($permission.Name) already assigned" -ForegroundColor Yellow
    }
}

Write-Host "Permissions assignment completed" -ForegroundColor Green