# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
    Generates a report of all users and their managers and uploads to SharePoint using Managed Identity.

.DESCRIPTION
    Connects to Microsoft Graph using the Managed Identity of the Automation Account.
    Retrieves user and manager data, exports it to Excel, and uploads the report to a SharePoint document library.

.NOTES
    Requires the Az.Accounts module and Export-Excel module to be imported in your automation account.
    The Managed Identity must have the correct Graph API and SharePoint permissions assigned.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteId,

    [Parameter(Mandatory = $true)]
    [string]$DriveId
    
)

function Get-MsGraphToken {
    try {
        Write-Host "Authenticating with Managed Identity..."
        Connect-AzAccount -Identity | Out-Null

        $tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

        if ($tokenObj.Token -is [System.Security.SecureString]) {
            Write-Host "Token is SecureString, converting to plain text..."
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            )
        } else {
            Write-Host "Token is plain string, no conversion needed."
            $token = $tokenObj.Token
        }

        if (-not [string]::IsNullOrEmpty($token)) {
            Write-Host "Token acquired successfully."
            return $token
        } else {
            throw "Token was empty."
        }
    }
    catch {
        Write-Error "Failed to acquire Microsoft Graph token using Managed Identity: $_"
        throw
    }
}

$accessToken = Get-MsGraphToken

$users = @()
$filter = "userType eq 'Member' and assignedLicenses/$count ne 0"
$select = "id,displayName,givenName,surname,jobTitle,department,userPrincipalName,mail"
$expand = "manager"
$nextLink = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,givenName,surname,jobTitle,department,userPrincipalName,mail,userType,assignedLicenses&`$expand=manager"
$headers = @{ Authorization = "Bearer $accessToken" }

do {
    $response = Invoke-RestMethod -Uri $nextLink -Headers $headers -Method Get
    foreach ($user in $response.value) {
        # Skip if they're a guest or unlicensed
        if ($user.userType -ne 'Member' -or !$user.assignedLicenses.Count) {
            continue
        }

        $users += [PSCustomObject]@{
            FirstName         = $user.givenName
            LastName          = $user.surname
            DisplayName       = $user.displayName
            Title             = $user.jobTitle
            Department        = $user.department
            UserPrincipalName = $user.userPrincipalName
            Email             = $user.mail
            ManagerName       = $user.manager.displayName
            ManagerEmail      = $user.manager.mail
        }
    }

    $nextLink = $response.'@odata.nextLink'
} while ($nextLink)

$currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$excelPath = "$Env:TEMP\users_managers_report_$currentDate.xlsx"
$users | Export-Excel -Path $excelPath -AutoSize -WorksheetName "UsersAndManagers" -TableName "UsersManagersTable"
$fileContent = [System.IO.File]::ReadAllBytes($excelPath)
$fileName = "Users_Managers_Report_$currentDate.xlsx"
$uploadUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root:/$fileName`:/content"

Invoke-RestMethod -Uri $uploadUri -Headers @{ Authorization = "Bearer $accessToken" } -Method PUT -Body $fileContent -ContentType "application/octet-stream"
Remove-Item -Path $excelPath -Force
Write-Output "Report uploaded to SharePoint successfully."