<#
.SYNOPSIS
    Generates a report of all shared items in a specified user's OneDrive account and stores the report in Azure Blob Storage.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all shared items from the specified user's OneDrive, exports the data to a CSV file,
    and uploads the file to a specified Azure Storage Blob container.
    
.PARAMETER UserPrincipalName
    The user principal name (email address) of the OneDrive account to scan.
    
.PARAMETER StorageAccountName
    The name of the Azure Storage Account where the report will be stored.
    
.PARAMETER StorageContainerName
    The name of the Blob container in the Storage Account where the report will be stored.
    
.PARAMETER IncludeAllFolders
    Optional. If specified, scans all folders including subfolders. If not specified,
    only scans the root folder. Default is $true.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually create or upload the report.
    
.NOTES
    File Name: Get-OneDriveSharedItemsReport.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-05-05
    
    Required Graph API Permissions for Managed Identity:
    - Files.Read.All or Files.ReadWrite.All
    - User.Read.All
    
    Required Storage Permissions for Managed Identity:
    - Storage Blob Data Contributor role
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $true)]
    [string]$StorageAccountName,
    
    [Parameter(Mandatory = $true)]
    [string]$StorageContainerName,
    
    [Parameter(Mandatory = $false)]
    [bool]$IncludeAllFolders = $true,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { 
            Write-Error $Message
            Write-Verbose $LogMessage -Verbose
        }
        "WARNING" { 
            Write-Warning $Message 
            Write-Verbose $LogMessage -Verbose
        }
        "WHATIF" { 
            Write-Verbose "[WHATIF] $Message" -Verbose
        }
        default { 
            Write-Verbose $LogMessage -Verbose
        }
    }
}

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

function Invoke-MsGraphRequestWithRetry {
    param (
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $params = @{
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        Method      = $Method
        ContentType = $ContentType
    }
    
    if ($null -ne $Body -and $Method -ne "GET") {
        $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
    }
    
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            
            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $retryCount -lt $MaxRetries) {
                $retryAfter = $backoffSeconds
                if ($_.Exception.Response -ne $null -and $_.Exception.Response.Headers -ne $null) {
                    $retryAfterHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq "Retry-After" }
                    if ($retryAfterHeader) {
                        $retryAfter = [int]$retryAfterHeader.Value[0]
                    }
                }
                
                if ($statusCode -eq 429) {
                    Write-Log "Request throttled by Graph API (429). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                else {
                    Write-Log "Server error (5xx). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                
                Start-Sleep -Seconds $retryAfter
                
                $retryCount++
                $backoffSeconds = $backoffSeconds * 2
            }
            else {
                Write-Log "Graph API request failed with status code $statusCode`: $_" -Type "ERROR"
                throw $_
            }
        }
    }
}

function Get-UserIdFromUpn {
    param (
        [string]$Token,
        [string]$UserPrincipalName,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Looking up user ID for UPN: $UserPrincipalName"
        $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        if ($response -and $response.id) {
            Write-Log "Found user ID: $($response.id) for $UserPrincipalName"
            return $response.id
        } else {
            throw "User ID not found for $UserPrincipalName"
        }
    }
    catch {
        Write-Log "Error retrieving user ID: $_" -Type "ERROR"
        throw "Failed to retrieve user ID for $UserPrincipalName`: $_"
    }
}

function Get-UserOneDriveId {
    param (
        [string]$Token,
        [string]$UserId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Getting OneDrive drive ID for user ID: $UserId"
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        if ($response -and $response.id) {
            Write-Log "Found OneDrive drive ID: $($response.id)"
            return $response.id
        } else {
            throw "OneDrive drive ID not found for user with ID $UserId"
        }
    }
    catch {
        Write-Log "Error retrieving OneDrive drive ID: $_" -Type "ERROR"
        throw "Failed to retrieve OneDrive drive ID for user with ID $UserId`: $_"
    }
}

function Get-OneDriveSharedItems {
    param (
        [string]$Token,
        [string]$UserId,
        [string]$DriveId,
        [string]$FolderId = "root",
        [bool]$IncludeAllFolders = $true,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5,
        [System.Collections.ArrayList]$SharedItems = $null
    )
    
    if ($null -eq $SharedItems) {
        $SharedItems = New-Object System.Collections.ArrayList
    }
    
    try {
        Write-Log "Scanning items in folder: $FolderId"
        $itemsUri = "https://graph.microsoft.com/v1.0/users/$UserId/drives/$DriveId/items/$FolderId/children"
        
        $items = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $itemsUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $items += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of items..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $items += $response.value
        }
        
        Write-Log "Found $($items.Count) items in folder: $FolderId"
        
        foreach ($item in $items) {
            $itemPermissionsUri = "https://graph.microsoft.com/v1.0/users/$UserId/drives/$DriveId/items/$($item.id)/permissions"
            $permissionsResponse = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $itemPermissionsUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $permissions = $permissionsResponse.value
            
            if ($permissions -and ($permissions | Where-Object { -not $_.inheritedFrom })) {
                $isShared = $false
                $shareType = ""
                $sharedWith = ""
                $shareLink = ""
                $roles = ""
                
                foreach ($permission in $permissions) {
                    if ($permission.inheritedFrom) {
                        continue
                    }
                    
                    $isShared = $true
                    
                    $permRoles = $permission.roles -join ", "
                    if ($roles) {
                        $roles += "; $permRoles"
                    } else {
                        $roles = $permRoles
                    }
                    
                    if ($permission.link) {
                        if ($shareType) {
                            $shareType += "; Link"
                        } else {
                            $shareType = "Link"
                        }
                        
                        if ($permission.link.scope -eq "anonymous") {
                            $shareType = "Anonymous Link"
                        } elseif ($permission.link.scope -eq "organization") {
                            $shareType = "Organization Link"
                        }
                        
                        if ($shareLink) {
                            $shareLink += "; $($permission.link.webUrl)"
                        } else {
                            $shareLink = $permission.link.webUrl
                        }
                    }
                    
                    if ($permission.grantedToIdentities) {
                        foreach ($identity in $permission.grantedToIdentities) {
                            if ($shareType -notlike "*Direct*") {
                                if ($shareType) {
                                    $shareType += "; Direct"
                                } else {
                                    $shareType = "Direct"
                                }
                            }
                            
                            if ($identity.user) {
                                if ($sharedWith) {
                                    $sharedWith += "; $($identity.user.email)"
                                } else {
                                    $sharedWith = $identity.user.email
                                }
                            } elseif ($identity.group) {
                                if ($sharedWith) {
                                    $sharedWith += "; Group: $($identity.group.displayName)"
                                } else {
                                    $sharedWith = "Group: $($identity.group.displayName)"
                                }
                            }
                        }
                    }
                }
                
                if ($isShared) {
                    $parentPath = ""
                    if ($item.parentReference -and $item.parentReference.path) {
                        $parentPath = $item.parentReference.path -replace "^.+/root:", ""
                    }
                    
                    $sharingId = ""
                    $sharingInfoUri = "https://graph.microsoft.com/v1.0/users/$UserId/drives/$DriveId/items/$($item.id)?`$select=id,name,sharepointIds"
                    try {
                        $sharingInfoResponse = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $sharingInfoUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                        if ($sharingInfoResponse.sharepointIds -and $sharingInfoResponse.sharepointIds.siteItemUniqueId) {
                            $sharingId = $sharingInfoResponse.sharepointIds.siteItemUniqueId
                        }
                    }
                    catch {
                        Write-Log "Unable to retrieve sharepointIds for $($item.name): $_" -Type "WARNING"
                    }
                    
                    $sharedItem = [PSCustomObject]@{
                        Name = $item.name
                        ItemType = $item.folder ? "Folder" : "File"
                        WebUrl = $item.webUrl
                        Path = $parentPath
                        Size = $item.size
                        CreatedDateTime = $item.createdDateTime
                        LastModifiedDateTime = $item.lastModifiedDateTime
                        SharedType = $shareType
                        SharedWith = $sharedWith
                        ShareLink = $shareLink
                        Permissions = $roles
                        ItemId = $item.id
                        SharingId = $sharingId
                    }
                    
                    [void]$SharedItems.Add($sharedItem)
                    Write-Log "Found shared item: $($item.name), Shared as: $shareType, Shared with: $sharedWith"
                }
            }
            
            if ($IncludeAllFolders -and $item.folder) {
                Write-Log "Processing subfolder: $($item.name)"
                $SharedItems = Get-OneDriveSharedItems -Token $Token -UserId $UserId -DriveId $DriveId -FolderId $item.id -IncludeAllFolders $IncludeAllFolders -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds -SharedItems $SharedItems
            }
        }
        
        return $SharedItems
    }
    catch {
        Write-Log "Error retrieving shared items from folder $FolderId`: $_" -Type "ERROR"
        throw "Failed to retrieve shared items: $_"
    }
}

function Export-ToCsv {
    param (
        [System.Collections.ArrayList]$SharedItems,
        [string]$OutputPath
    )
    
    try {
        Write-Log "Exporting $($SharedItems.Count) shared items to CSV: $OutputPath"
        
        $SharedItems | Export-Csv -Path $OutputPath -NoTypeInformation
        
        if (Test-Path $OutputPath) {
            Write-Log "CSV file created successfully at: $OutputPath"
            return $true
        } else {
            throw "CSV file was not created at: $OutputPath"
        }
    }
    catch {
        Write-Log "Error creating CSV file: $_" -Type "ERROR"
        return $false
    }
}

function Upload-ToAzureBlob {
    param (
        [string]$StorageAccountName,
        [string]$ContainerName,
        [string]$FilePath,
        [string]$BlobName,
        [switch]$WhatIf
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would upload file $FilePath to blob container $ContainerName as $BlobName" -Type "WHATIF"
            return "WhatIf-BlobUrl"
        }
        
        Write-Log "Uploading file to Azure Storage..."
        
        $storageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
        $container = Get-AzStorageContainer -Name $ContainerName -Context $storageContext -ErrorAction SilentlyContinue
        
        if ($null -eq $container) {
            Write-Log "Container $ContainerName does not exist. Creating new container..." -Type "WARNING"
            New-AzStorageContainer -Name $ContainerName -Context $storageContext -Permission Off
        }
        
        $blobProperties = @{
            File      = $FilePath
            Container = $ContainerName
            Blob      = $BlobName
            Context   = $storageContext
            Properties = @{
                ContentType = "text/csv"
            }
        }
        
        $blob = Set-AzStorageBlobContent @blobProperties -Force
        
        if ($blob) {
            Write-Log "File uploaded successfully to $($blob.Name)"
            
            $blobUrl = $blob.ICloudBlob.Uri.AbsoluteUri
            Write-Log "Blob URL: $blobUrl"
            
            return $blobUrl
        }
        else {
            Write-Log "Failed to upload file to Azure Storage" -Type "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error uploading file to Azure Storage: $_" -Type "ERROR"
        return $false
    }
}

# Main script logic
try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO ACTUAL REPORTS WILL BE CREATED OR UPLOADED ===" -Type "WHATIF"
    }
    
    Write-Log "=== OneDrive Shared Items Report Process Started ==="
    Write-Log "User to scan: $UserPrincipalName"
    Write-Log "Storage Account: $StorageAccountName"
    Write-Log "Container Name: $StorageContainerName"
    Write-Log "Include All Folders: $IncludeAllFolders"
    
    $startTime = Get-Date
    $reportTimestamp = $startTime
    $token = Get-MsGraphToken
    $userId = Get-UserIdFromUpn -Token $token -UserPrincipalName $UserPrincipalName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $driveId = Get-UserOneDriveId -Token $token -UserId $userId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $sharedItems = Get-OneDriveSharedItems -Token $token -UserId $userId -DriveId $driveId -IncludeAllFolders $IncludeAllFolders -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $fileName = "OneDriveSharedItems_$($UserPrincipalName.Split('@')[0])_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $fileName)
    
    if ($sharedItems.Count -gt 0) {
        Write-Log "Found $($sharedItems.Count) shared items in OneDrive"
        
        $csvExported = Export-ToCsv -SharedItems $sharedItems -OutputPath $tempPath
        
        if ($csvExported) {
            $blobUrl = Upload-ToAzureBlob -StorageAccountName $StorageAccountName -ContainerName $StorageContainerName -FilePath $tempPath -BlobName $fileName -WhatIf:$WhatIf
            
            if (Test-Path $tempPath) {
                Remove-Item -Path $tempPath -Force
                Write-Log "Temporary file removed"
            }
        }
    } else {
        Write-Log "No shared items found in the OneDrive account" -Type "WARNING"
        
        $emptyItem = [PSCustomObject]@{
            Name = ""
            ItemType = ""
            WebUrl = ""
            Path = ""
            Size = ""
            CreatedDateTime = ""
            LastModifiedDateTime = ""
            SharedType = ""
            SharedWith = ""
            ShareLink = ""
            Permissions = ""
            ItemId = ""
            SharingId = ""
        }
        
        $emptyCollection = New-Object System.Collections.ArrayList
        [void]$emptyCollection.Add($emptyItem)
        $csvExported = Export-ToCsv -SharedItems $emptyCollection -OutputPath $tempPath
        
        if ($csvExported) {
            $blobUrl = Upload-ToAzureBlob -StorageAccountName $StorageAccountName -ContainerName $StorageContainerName -FilePath $tempPath -BlobName $fileName -WhatIf:$WhatIf
            
            if (Test-Path $tempPath) {
                Remove-Item -Path $tempPath -Force
                Write-Log "Temporary file removed"
            }
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Log "=== OneDrive Shared Items Report Process Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO REPORT WAS CREATED OR UPLOADED ===" -Type "WHATIF"
    }
    
    Write-Log "User: $UserPrincipalName"
    Write-Log "Shared Items: $($sharedItems.Count)"
    
    $shareTypeStats = @{}
    foreach ($item in $sharedItems) {
        if ($item.SharedType -match "Anonymous") {
            if (-not $shareTypeStats.ContainsKey("Anonymous")) {
                $shareTypeStats["Anonymous"] = 0
            }
            $shareTypeStats["Anonymous"]++
        }
        elseif ($item.SharedType -match "Organization") {
            if (-not $shareTypeStats.ContainsKey("Organization")) {
                $shareTypeStats["Organization"] = 0
            }
            $shareTypeStats["Organization"]++
        }
        elseif ($item.SharedType -match "Direct") {
            if (-not $shareTypeStats.ContainsKey("Direct")) {
                $shareTypeStats["Direct"] = 0
            }
            $shareTypeStats["Direct"]++
        }
        elseif ($item.SharedType -match "Link") {
            if (-not $shareTypeStats.ContainsKey("Link")) {
                $shareTypeStats["Link"] = 0
            }
            $shareTypeStats["Link"]++
        }
    }
    
    foreach ($shareType in $shareTypeStats.Keys) {
        Write-Log "Share Type - $shareType`: $($shareTypeStats[$shareType])"
    }
    
    $outputObject = [PSCustomObject][ordered]@{
        UserPrincipalName = $UserPrincipalName
        TotalSharedItems = $sharedItems.Count
        AnonymousShares = $shareTypeStats["Anonymous"] ? $shareTypeStats["Anonymous"] : 0
        OrganizationShares = $shareTypeStats["Organization"] ? $shareTypeStats["Organization"] : 0
        DirectShares = $shareTypeStats["Direct"] ? $shareTypeStats["Direct"] : 0
        LinkShares = $shareTypeStats["Link"] ? $shareTypeStats["Link"] : 0
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        ReportUrl = $blobUrl
        Timestamp = $reportTimestamp.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    return $outputObject
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
    
    if ($tempPath -and (Test-Path $tempPath)) {
        Remove-Item -Path $tempPath -Force
        Write-Log "Temporary file cleaned up during exit"
    }
}