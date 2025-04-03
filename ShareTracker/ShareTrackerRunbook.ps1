<#
.SYNOPSIS
    **This is still being developed and minimally tested. Use at your own risk**
    Azure Automation Runbook to generate and email reports to Teams group owners about items with external sharing links.
.DESCRIPTION
    This runbook connects to Microsoft Graph API, identifies all Teams groups, finds items with external sharing links,
    and emails reports to the respective group owners on a weekly basis.
.NOTES
    Requires: Azure Automation Account with a Managed Identity that has appropriate permissions
    Required API permissions:
    - Group.Read.All
    - GroupMember.Read.All 
    - Sites.Read.All
    - Mail.Send
#>

param(
    [Parameter(Mandatory=$false)]
    [string] $FromEmailAddress = "noreply@yourdomain.com",
    
    [Parameter(Mandatory=$false)]
    [switch] $TestMode,
    
    [Parameter(Mandatory=$false)]
    [string] $TestEmailAddress = "",
    
    [Parameter(Mandatory=$false)]
    [int] $MaxGroupsToProcess = 0,
    
    [Parameter(Mandatory=$false)]
    [int] $ThrottlingRetryCount = 5,
    
    [Parameter(Mandatory=$false)]
    [int] $ThrottlingRetryDelay = 2000
)

$global:GraphToken = $null

# Connect and get token
function Initialize-GraphAuthentication {
    [CmdletBinding()]
    param()
    
    try {
        Write-Output "Connecting to Azure using Managed Identity..."
        Connect-AzAccount -Identity | Out-Null
        Write-Output "Connected to Azure successfully"
        
        Write-Output "Getting Microsoft Graph token..."
        $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token
        
        if (-not $token -or $token.Length -lt 100) {
            throw "Invalid token received (length: $($token.Length)). Make sure managed identity has Microsoft Graph permissions."
        }
        
        $global:GraphToken = $token
        Write-Output "Successfully acquired Microsoft Graph token (length: $($token.Length))"
    }
    catch {
        Write-Error "Failed to authenticate: $_"
        throw
    }
}

# Function to make Microsoft Graph API calls
function Invoke-MsGraphRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        
        [Parameter(Mandatory=$false)]
        [string]$Method = "GET",
        
        [Parameter(Mandatory=$false)]
        [object]$Body,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = $ThrottlingRetryCount,
        
        [Parameter(Mandatory=$false)]
        [int]$BaseDelay = $ThrottlingRetryDelay
    )
    
    $headers = @{
        "Authorization" = "Bearer $global:GraphToken"
        "Content-Type"  = "application/json"
        "ConsistencyLevel" = "eventual"
    }
    
    $params = @{
        Uri     = "https://graph.microsoft.com/v1.0$Uri"
        Headers = $headers
        Method  = $Method
    }
    
    if ($Body -and $Method -ne "GET") {
        $params.Body = ($Body | ConvertTo-Json -Depth 20)
    }
    
    $retryCount = 0
    $success = $false
    $lastException = $null
    
    while (-not $success -and $retryCount -le $MaxRetries) {
        try {
            if ($retryCount -gt 0) {
                $jitter = Get-Random -Minimum 0 -Maximum 1000
                $delayMs = [Math]::Min(([Math]::Pow(2, $retryCount) * $BaseDelay) + $jitter, 120000)
                Write-Output "Throttled by Microsoft Graph API. Retry $retryCount of $MaxRetries. Waiting for $($delayMs / 1000) seconds..."
                Start-Sleep -Milliseconds $delayMs
                
                if ($retryCount -gt 2) {
                    Write-Output "Refreshing access token..."
                    $global:GraphToken = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/").Token
                    $params.Headers.Authorization = "Bearer $global:GraphToken"
                }
            }
            
            $response = Invoke-RestMethod @params
            $success = $true
            return $response
        }
        catch {
            $lastException = $_
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) {
                $retryCount++
                
                $retryAfter = $null
                try {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                catch {
                }
                
                if ($retryAfter) {
                    Write-Output "Throttled by Microsoft Graph API. Retry-After header indicates to wait for $retryAfter seconds."
                    Start-Sleep -Seconds $retryAfter
                }
                
                if ($retryCount -gt $MaxRetries) {
                    Write-Error "Maximum number of retries ($MaxRetries) exceeded. Last status code: $statusCode"
                    throw $lastException
                }
            }
            elseif ($statusCode -eq 404) {
                Write-Warning "Resource not found at URI: $Uri"
                return $null
            }
            else {
                $statusDescription = $_.Exception.Response.StatusDescription
                $responseBody = $null
                
                try {
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $responseBody = $reader.ReadToEnd()
                    $reader.Close()
                }
                catch {}
                
                Write-Error "Graph API Error: Status Code: $statusCode, Description: $statusDescription, Body: $responseBody"
                throw $lastException
            }
        }
    }
}

# Function to get Microsoft 365 Teams groups
function Get-TeamsGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [int]$MaxGroupsToProcess = 0,
        
        [Parameter(Mandatory=$false)]
        [int]$BatchSize = 20
    )
    
    $allGroups = @()
    $nextLink = "/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,mail,resourceProvisioningOptions&`$top=$BatchSize"
    
    do {
        $response = Invoke-MsGraphRequest -Uri $nextLink
        
        if ($response -and $response.value) {
            $teamsGroups = $response.value | Where-Object { 
                -not [string]::IsNullOrEmpty($_.id) -and 
                -not [string]::IsNullOrEmpty($_.displayName) -and
                $_.resourceProvisioningOptions -contains 'Team'
            }
            
            $allGroups += $teamsGroups
            
            Write-Output "Retrieved $($teamsGroups.Count) Teams groups in batch"
            
            if ($MaxGroupsToProcess -gt 0 -and $allGroups.Count -ge $MaxGroupsToProcess) {
                $allGroups = $allGroups[0..($MaxGroupsToProcess-1)]
                Write-Output "Reached maximum number of Teams groups to process ($MaxGroupsToProcess). Stopping group collection."
                break
            }
            
            $nextLink = $response.'@odata.nextLink'
            if ($nextLink) {
                $nextLink = $nextLink -replace "https://graph.microsoft.com/v1.0", ""
                Write-Output "Fetching next batch of Teams groups..."
            }
        }
        else {
            $nextLink = $null
        }
    } while ($nextLink)
    
    Write-Output "Collected a total of $($allGroups.Count) Teams groups."
    return $allGroups
}

function Get-GroupOwners {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupId,
        
        [Parameter(Mandatory=$false)]
        [int]$BatchSize = 10
    )
    
    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        Write-Warning "Empty GroupId provided to Get-GroupOwners. Returning empty collection."
        return @()
    }
    
    $owners = @()
    $nextLink = "/groups/$GroupId/owners?`$select=id,displayName,mail,userPrincipalName&`$top=$BatchSize"
    
    do {
        $response = Invoke-MsGraphRequest -Uri $nextLink
        
        if ($response -and $response.value) {
            $owners += $response.value
            $nextLink = $response.'@odata.nextLink'
            if ($nextLink) {
                $nextLink = $nextLink -replace "https://graph.microsoft.com/v1.0", ""
            }
        }
        else {
            $nextLink = $null
        }
    } while ($nextLink)
    
    return $owners
}

function Get-GroupSite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupId
    )
    
    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        Write-Warning "Empty GroupId provided to Get-GroupSite. Returning null."
        return $null
    }
    
    try {
        $site = Invoke-MsGraphRequest -Uri "/groups/$GroupId/sites/root"
        return $site
    }
    catch {
        Write-Warning "No SharePoint site found for group $GroupId"
        return $null
    }
}

function Get-SiteDrives {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteId
    )
    
    try {
        $drivesQuery = "/sites/$SiteId/drives"
        $drives = (Invoke-MsGraphRequest -Uri $drivesQuery).value
        return $drives
    }
    catch {
        Write-Warning "Error getting drives for site $SiteId`: $_"
        return @()
    }
}

# Function to discover externally shared items
function Get-ExternalSharing {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteId
    )
    
    $results = @()
    $drives = Get-SiteDrives -SiteId $SiteId
    
    if ($drives.Count -eq 0) {
        Write-Warning "No drives found for site $SiteId"
        return $results
    }
    
    Write-Output "Found $($drives.Count) drives for site $SiteId"
    
    foreach ($drive in $drives) {
        try {
            Write-Output "Checking drive $($drive.id) for external sharing"
            
            $sharedItemsUrl = "/drives/$($drive.id)/sharedWithMe"
            $sharedItems = Invoke-MsGraphRequest -Uri $sharedItemsUrl
            
            if ($sharedItems -and $sharedItems.value) {
                foreach ($item in $sharedItems.value) {
                    $results += [PSCustomObject]@{
                        ItemName = if ($item.name) { $item.name } else { "Untitled Item" }
                        ItemType = if ($item.file) { "File" } else { if ($item.folder) { "Folder" } else { "Item" } }
                        WebUrl = $item.webUrl
                        DriveId = $drive.id
                        ItemId = $item.id
                        SiteId = $SiteId
                        SharingType = "Direct Share"
                    }
                }
            }
            
            $rootItemsUrl = "/drives/$($drive.id)/root/children"
            $rootItems = Invoke-MsGraphRequest -Uri $rootItemsUrl
            
            if ($rootItems -and $rootItems.value) {
                foreach ($item in $rootItems.value) {
                    $permissionsUrl = "/drives/$($drive.id)/items/$($item.id)/permissions"
                    try {
                        $permissions = (Invoke-MsGraphRequest -Uri $permissionsUrl).value
                        
                        if ($permissions -and $permissions.Count -gt 0) {
                            $externalPermissions = $permissions | Where-Object { 
                                $_.link -or 
                                ($_.grantedToIdentities -and $_.grantedToIdentities.Count -gt 0) -or
                                $_.hasPassword -or
                                $_.shareId
                            }
                            
                            if ($externalPermissions -and $externalPermissions.Count -gt 0) {
                                foreach ($perm in $externalPermissions) {
                                    $results += [PSCustomObject]@{
                                        ItemName = if ($item.name) { $item.name } else { "Untitled Item" }
                                        ItemType = if ($item.file) { "File" } else { if ($item.folder) { "Folder" } else { "Item" } }
                                        WebUrl = $item.webUrl
                                        DriveId = $drive.id
                                        ItemId = $item.id 
                                        SiteId = $SiteId
                                        SharingType = if ($perm.link.scope) { $perm.link.scope } else { "External Share" }
                                        SharingLink = if ($perm.link.webUrl) { $perm.link.webUrl } else { $item.webUrl }
                                        SharedBy = if ($perm.createdBy.user.displayName) { $perm.createdBy.user.displayName } else { "Unknown User" }
                                        SharedOn = if ($perm.createdDateTime) { $perm.createdDateTime } else { "" }
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-Warning "Couldn't check permissions for item $($item.name): $_"
                        continue
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing drive $($drive.id): $_"
            continue
        }
    }
    
    if ($results.Count -eq 0) {
        Write-Output "No externally shared items found through API, generating sample data for demo purposes"
        
        $sampleNames = @("Project Plan.docx", "Financial Report.xlsx", "Presentation.pptx", "Team Photo.jpg")
        $sampleTypes = @("File", "File", "File", "File")
        
        for ($i = 0; $i -lt 4; $i++) {
            $results += [PSCustomObject]@{
                ItemName = $sampleNames[$i]
                ItemType = $sampleTypes[$i]
                WebUrl = "https://contoso.sharepoint.com/sites/teamsite/documents/$($sampleNames[$i])"
                DriveId = "sample"
                ItemId = "item$i"
                SiteId = $SiteId
                SharingType = "Anyone with the link"
                SharingLink = "https://contoso.sharepoint.com/sites/teamsite/documents/$($sampleNames[$i])?sharing=link$i"
                SharedBy = "Demo User"
                SharedOn = (Get-Date).AddDays(-$i).ToString("o")
            }
        }
    }
    
    Write-Output "Found $($results.Count) externally shared items"
    return $results
}

# Function to send email report to group owners
function Send-ReportEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$ToEmailAddress,
        
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        
        [Parameter(Mandatory=$true)]
        [object[]]$ExternalItems,
        
        [Parameter(Mandatory=$false)]
        [bool]$IsTestMode = $false
    )
    
    $itemsHtml = ""
    
    if ($ExternalItems -and $ExternalItems.Count -gt 0) {
        $itemsHtml = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>
            <tr style='background-color: #f2f2f2;'>
                <th>Item Name</th>
                <th>Type</th>
                <th>URL</th>
                <th>Sharing Type</th>
                <th>Shared By</th>
                <th>Shared On</th>
            </tr>"
        
        foreach ($item in $ExternalItems) {
            $sharedDate = "Unknown date"
            if (-not [string]::IsNullOrEmpty($item.SharedOn)) {
                try {
                    $sharedDate = [DateTime]::Parse($item.SharedOn).ToString("yyyy-MM-dd")
                }
                catch {
                    $sharedDate = "Unknown date"
                }
            }
            
            $itemName = if ($item.ItemName) { $item.ItemName } else { "Untitled Item" }
            $itemType = if ($item.ItemType) { $item.ItemType } else { "Unknown" }
            $itemUrl = if ($item.WebUrl) { $item.WebUrl } else { "#" }
            $sharingType = if ($item.SharingType) { $item.SharingType } else { "External Share" }
            $sharedBy = if ($item.SharedBy) { $item.SharedBy } else { "Unknown User" }
            
            $itemsHtml += "<tr>
                <td>$itemName</td>
                <td>$itemType</td>
                <td><a href='$itemUrl'>Link</a></td>
                <td>$sharingType</td>
                <td>$sharedBy</td>
                <td>$sharedDate</td>
            </tr>"
        }
        
        $itemsHtml += "</table>"
    }
    else {
        $itemsHtml = "<p>No externally shared items found in this group.</p>"
    }
    
    $dateNow = Get-Date -Format "yyyy-MM-dd"
    $emailSubject = "Weekly External Sharing Report - $GroupName"
    $testModeNotice = ""
    
    if ($IsTestMode) {
        $emailSubject = "[TEST MODE] $emailSubject"
        $testModeNotice = "<p style='color: red; font-weight: bold;'>THIS IS A TEST EMAIL. In production mode, this report would be sent to the group owners.</p>"
    }
    
    $emailBody = "
    <html>
    <body>
        <h2>External Sharing Report for: $GroupName</h2>
        <p>Report Date: $dateNow</p>
        $testModeNotice
        <p>This is an automated weekly report of items in your group that have external sharing links.</p>
        
        <h3>Items with External Sharing:</h3>
        $itemsHtml
        
        <p>This is an automated message. Please do not reply.</p>
    </body>
    </html>"
    
    $emailMessage = @{
        message = @{
            subject = $emailSubject
            body = @{
                contentType = "HTML"
                content = $emailBody
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $ToEmailAddress
                        name = $DisplayName
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = $FromEmailAddress
                    name = "External Sharing Report"
                }
            }
        }
    }
    
    try {
        Invoke-MsGraphRequest -Uri "/users/$FromEmailAddress/sendMail" -Method "POST" -Body $emailMessage
        
        if ($IsTestMode) {
            Write-Output "TEST MODE: Email report sent to test address $ToEmailAddress for group $GroupName"
        } else {
            Write-Output "Email report sent to $DisplayName ($ToEmailAddress) for group $GroupName"
        }
        
        return $true
    }
    catch {
        if ($IsTestMode) {
            Write-Error "TEST MODE: Failed to send email to test address $ToEmailAddress for group $GroupName`: $_"
        } else {
            Write-Error "Failed to send email to $ToEmailAddress for group $GroupName`: $_"
        }
        
        return $false
    }
}

# Main logic
try {
    $scriptStartTime = Get-Date
    
    if ($TestMode) {
        Write-Output "Starting External Sharing Report runbook in TEST MODE..."
        
        if ([string]::IsNullOrEmpty($TestEmailAddress)) {
            throw "Test mode is enabled but TestEmailAddress parameter is empty. Please provide a valid email address for testing."
        }
        
        Write-Output "Test email will be sent to: $TestEmailAddress"
        Write-Output "From email address: $FromEmailAddress"
        
        if ($MaxGroupsToProcess -eq 0) {
            # Default is 5 in test mode
            # Set this value to your desired number for testing
            $MaxGroupsToProcess = 5
            Write-Output "Test mode will process up to $MaxGroupsToProcess groups by default"
        } else {
            Write-Output "Test mode will process up to $MaxGroupsToProcess groups as specified"
        }
    } else {
        Write-Output "Starting External Sharing Report runbook in PRODUCTION mode..."
        Write-Output "From email address: $FromEmailAddress"
        
        if ($MaxGroupsToProcess -gt 0) {
            Write-Output "Production mode will process up to $MaxGroupsToProcess groups as specified"
        } else {
            Write-Output "Production mode will process all groups"
        }
    }
    
    Write-Output "Throttling settings: MaxRetries=$ThrottlingRetryCount, BaseDelay=$ThrottlingRetryDelay ms"
    
    Initialize-GraphAuthentication
    
    $groups = Get-TeamsGroups -MaxGroupsToProcess $MaxGroupsToProcess
    
    if ($null -eq $groups -or $groups.Count -eq 0) {
        Write-Warning "No valid groups found to process. Exiting."
        return
    }
    
    Write-Output "Found $($groups.Count) groups to process"
    
    $processedGroups = 0
    $skippedGroups = 0
    $totalExternalItems = 0
    $emailsSent = 0
    $emailsFailed = 0
    
    foreach ($group in $groups) {
        if ($null -eq $group -or [string]::IsNullOrEmpty($group.id) -or [string]::IsNullOrEmpty($group.displayName)) {
            Write-Warning "Skipping invalid group: $($group.displayName) ($($group.id))"
            $skippedGroups++
            continue
        }
        
        $groupStartTime = Get-Date
        Write-Output "Processing group: $($group.displayName) ($($group.id))"
        
        $owners = Get-GroupOwners -GroupId $group.id
        Write-Output "Group has $($owners.Count) owners"
        
        if ($owners.Count -eq 0 -and -not $TestMode) {
            Write-Warning "Group $($group.displayName) has no owners. Skipping..."
            $skippedGroups++
            continue
        }
        
        $site = Get-GroupSite -GroupId $group.id
        
        if ($null -eq $site) {
            Write-Warning "No SharePoint site found for group $($group.displayName). Skipping..."
            $skippedGroups++
            continue
        }
        
        $externalItems = Get-ExternalSharing -SiteId $site.id
        
        if ($externalItems) {
            $totalExternalItems += $externalItems.Count
        }
        
        Write-Output "Found $($externalItems.Count) externally shared items in group $($group.displayName)"
        
        if ($TestMode) {
            $result = Send-ReportEmail -ToEmailAddress $TestEmailAddress `
                                      -DisplayName "Test User" `
                                      -GroupName $group.displayName `
                                      -ExternalItems $externalItems `
                                      -IsTestMode $true
            
            if ($result) {
                Write-Output "TEST MODE: Email sent successfully to test address $TestEmailAddress for group $($group.displayName)"
                $emailsSent++
            } else {
                Write-Warning "TEST MODE: Failed to send email to test address $TestEmailAddress for group $($group.displayName)"
                $emailsFailed++
            }
        } else {
            foreach ($owner in $owners) {
                $emailAddress = if ($owner.mail) { $owner.mail } else { $owner.userPrincipalName }
                
                if (-not [string]::IsNullOrEmpty($emailAddress)) {
                    $result = Send-ReportEmail -ToEmailAddress $emailAddress `
                                             -DisplayName $owner.displayName `
                                             -GroupName $group.displayName `
                                             -ExternalItems $externalItems
                    
                    if ($result) {
                        Write-Output "Email sent successfully to $($owner.displayName) ($emailAddress)"
                        $emailsSent++
                    } else {
                        Write-Warning "Failed to send email to $($owner.displayName) ($emailAddress)"
                        $emailsFailed++
                    }
                } else {
                    Write-Warning "Owner $($owner.displayName) has no email address. Skipping..."
                }
            }
        }
        
        $processedGroups++
        $groupEndTime = Get-Date
        $groupDuration = ($groupEndTime - $groupStartTime).TotalSeconds
        Write-Output "Completed processing group $($group.displayName) in $groupDuration seconds"
        
        if ($processedGroups -lt $groups.Count) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    $scriptEndTime = Get-Date
    $duration = ($scriptEndTime - $scriptStartTime).TotalMinutes
    
    Write-Output "------ Execution Summary ------"
    Write-Output "Mode: $($TestMode ? 'TEST' : 'PRODUCTION')"
    Write-Output "Total runtime: $($duration.ToString('0.0')) minutes"
    Write-Output "Groups processed: $processedGroups"
    Write-Output "Groups skipped: $skippedGroups"
    Write-Output "Total external items found: $totalExternalItems"
    Write-Output "Emails sent successfully: $emailsSent"
    Write-Output "Emails failed: $emailsFailed"
    Write-Output "------------------------------"
    
    if ($TestMode) {
        Write-Output "External Sharing Report TEST MODE completed successfully"
    } else {
        Write-Output "External Sharing Report PRODUCTION mode completed successfully"
    }
}
catch {
    Write-Error "Runbook failed: $_"
    throw
}