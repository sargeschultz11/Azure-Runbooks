# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
    Sets the Company attribute for all member users in the Microsoft 365 tenant.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all member users, and updates the Company attribute to a specified value.
    It processes users in batches with built-in throttling protection to handle large tenants.
    
.PARAMETER CompanyName
    The company name to set for all users.
    
.PARAMETER BatchSize
    Optional. Number of users to process in each batch. Default is 50.
    
.PARAMETER BatchDelaySeconds
    Optional. Number of seconds to wait between processing batches. Default is 10.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER WhatIf
    Optional. If specified, shows what changes would occur without actually making any updates.
    
.PARAMETER ExcludeGuestUsers
    Optional. If specified, excludes guest users from processing. Default is $true.
    
.PARAMETER ExcludeServiceAccounts
    Optional. If specified, excludes accounts with 'service' in the display name or UPN. Default is $true.
    
.PARAMETER ProcessUnlicensedUsers
    Optional. If specified, processes unlicensed users. Default is $true.
    
.NOTES
    File Name: Set-CompanyAttribute.ps1
    Author: Ryan Schultz
    Version: 1.0
    
    Required Graph API Permissions for Managed Identity:
    - User.ReadWrite.All
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CompanyName,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchDelaySeconds = 10,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [bool]$ExcludeGuestUsers = $true,
    
    [Parameter(Mandatory = $false)]
    [bool]$ExcludeServiceAccounts = $true,
    
    [Parameter(Mandatory = $false)]
    [bool]$ProcessUnlicensedUsers = $true
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
        Write-Log "Authenticating with Managed Identity..."
        Connect-AzAccount -Identity | Out-Null

        $tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

        if ($tokenObj.Token -is [System.Security.SecureString]) {
            Write-Log "Token is SecureString, converting to plain text..."
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            )
        } else {
            Write-Log "Token is plain string, no conversion needed."
            $token = $tokenObj.Token
        }

        if (-not [string]::IsNullOrEmpty($token)) {
            Write-Log "Token acquired successfully."
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
        if ($Body -is [string]) {
            $params.Add("Body", $Body)
        } else {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
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

function Get-UsersToProcess {
    param (
        [string]$Token,
        [bool]$ExcludeGuestUsers = $true,
        [bool]$ExcludeServiceAccounts = $true,
        [bool]$ProcessUnlicensedUsers = $true,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Retrieving users to process..."
        
        $filter = @()
        
        # Add filters based on parameters
        if ($ExcludeGuestUsers) {
            $filter += "userType eq 'Member'"
        }
        
        # Build the final filter string
        $filterString = if ($filter.Count -gt 0) {
            "?`$filter=$($filter -join ' and ')"
        } else {
            ""
        }
        
        # Select properties
        $select = "id,displayName,userPrincipalName,companyName,userType,accountEnabled,assignedLicenses"
        $filterString += if ($filterString -eq "") { "?`$select=$select" } else { "&`$select=$select" }
        
        $uri = "https://graph.microsoft.com/v1.0/users$filterString"
        Write-Log "URI: $uri"
        $users = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $users += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of users..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $users += $response.value
        }
        
        Write-Log "Retrieved $($users.Count) users from Microsoft 365"
        
        $filteredUsers = $users | Where-Object {
            $include = $true
            
            if ($ExcludeServiceAccounts) {
                if ($_.displayName -match "service" -or $_.userPrincipalName -match "service") {
                    $include = $false
                }
            }
            
            if (-not $ProcessUnlicensedUsers) {
                if (-not $_.assignedLicenses -or $_.assignedLicenses.Count -eq 0) {
                    $include = $false
                }
            }
            
            $include
        }
        
        Write-Log "After additional filtering, $($filteredUsers.Count) users will be processed"
        return $filteredUsers
    }
    catch {
        Write-Log "Failed to retrieve users: $_" -Type "ERROR"
        throw "Failed to retrieve users: $_"
    }
}

function Update-UserCompanyAttribute {
    param (
        [string]$Token,
        [string]$UserId,
        [string]$UserDisplayName,
        [string]$UserPrincipalName,
        [string]$CurrentCompany,
        [string]$NewCompany,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        $needsUpdate = [string]::IsNullOrEmpty($CurrentCompany) -or $CurrentCompany -ne $NewCompany
        
        if (-not $needsUpdate) {
            Write-Log "User $UserDisplayName already has company set to '$CurrentCompany'. No update needed."
            return @{
                Success = $true
                Action = "NoActionNeeded"
            }
        }
        
        if ($WhatIf) {
            $action = [string]::IsNullOrEmpty($CurrentCompany) ? "Add" : "Update"
            Write-Log "Would $action company attribute for user $UserDisplayName ($UserPrincipalName) from '$CurrentCompany' to '$NewCompany'" -Type "WHATIF"
            return @{
                Success = $true
                Action = "WhatIf"
            }
        }
        else {
            Write-Log "Updating company attribute for user $UserDisplayName ($UserPrincipalName) from '$CurrentCompany' to '$NewCompany'"
            
            $uri = "https://graph.microsoft.com/v1.0/users/$UserId"
            
            $updateData = @{
                companyName = $NewCompany
            }
            
            Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "PATCH" -Body $updateData -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            Write-Log "Successfully updated company attribute for user $UserDisplayName"
            return @{
                Success = $true
                Action = "Updated"
            }
        }
    }
    catch {
        Write-Log "Failed to update company attribute for user $UserDisplayName`: $_" -Type "ERROR"
        return @{
            Success = $false
            Action = "Failed"
            Error = $_.ToString()
        }
    }
}

function Process-UserBatch {
    param (
        [string]$Token,
        [array]$Users,
        [string]$CompanyName,
        [switch]$WhatIf,
        [hashtable]$Stats,
        [int]$MaxRetries,
        [int]$InitialBackoffSeconds
    )
    
    $batchUpdateCount = 0
    $batchNoChangeCount = 0
    $batchErrorCount = 0
    
    foreach ($user in $Users) {
        try {
            $userId = $user.id
            $userDisplayName = $user.displayName
            $userPrincipalName = $user.userPrincipalName
            $currentCompany = $user.companyName
            
            Write-Log "Processing user: $userDisplayName (UPN: $userPrincipalName)"
            Write-Log "Current Company: '$currentCompany'"
            
            $updateResult = Update-UserCompanyAttribute -Token $Token -UserId $userId -UserDisplayName $userDisplayName -UserPrincipalName $userPrincipalName -CurrentCompany $currentCompany -NewCompany $CompanyName -WhatIf:$WhatIf -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($updateResult.Success) {
                if ($updateResult.Action -eq "Updated" -or $updateResult.Action -eq "WhatIf") {
                    $batchUpdateCount++
                    $Stats.UpdatedCount++
                }
                elseif ($updateResult.Action -eq "NoActionNeeded") {
                    $batchNoChangeCount++
                    $Stats.NoChangeCount++
                }
            }
            else {
                $batchErrorCount++
                $Stats.ErrorCount++
            }
        }
        catch {
            Write-Log "Error processing user $($user.displayName): $_" -Type "ERROR"
            $batchErrorCount++
            $Stats.ErrorCount++
        }
    }
    
    return @{
        UpdatedCount = $batchUpdateCount
        NoChangeCount = $batchNoChangeCount
        ErrorCount = $batchErrorCount
    }
}

# Main script execution
try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO CHANGES WILL BE MADE ===" -Type "WHATIF"
    }
    
    Write-Log "=== Company Attribute Update Process Started ==="
    Write-Log "Company name to set: $CompanyName"
    Write-Log "Batch Size: $BatchSize"
    Write-Log "Batch Delay: $BatchDelaySeconds seconds"
    Write-Log "Exclude Guest Users: $ExcludeGuestUsers"
    Write-Log "Exclude Service Accounts: $ExcludeServiceAccounts"
    Write-Log "Process Unlicensed Users: $ProcessUnlicensedUsers"
    
    $startTime = Get-Date
    
    $token = Get-MsGraphToken
    
    $users = Get-UsersToProcess -Token $token -ExcludeGuestUsers $ExcludeGuestUsers -ExcludeServiceAccounts $ExcludeServiceAccounts -ProcessUnlicensedUsers $ProcessUnlicensedUsers -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    $stats = @{
        TotalUsers = $users.Count
        UpdatedCount = 0
        NoChangeCount = 0
        ErrorCount = 0
    }
    
    $totalUsers = $users.Count
    $batches = [Math]::Ceiling($totalUsers / $BatchSize)
    Write-Log "Processing $totalUsers users in $batches batches of maximum $BatchSize users"
    
    for ($batchNum = 0; $batchNum -lt $batches; $batchNum++) {
        $start = $batchNum * $BatchSize
        $end = [Math]::Min(($batchNum + 1) * $BatchSize - 1, $totalUsers - 1)
        $currentBatchSize = $end - $start + 1
        
        Write-Log "Processing batch $($batchNum+1) of $batches (users $($start+1) to $($end+1) of $totalUsers)"
        
        $currentBatch = $users[$start..$end]
        
        $batchResult = Process-UserBatch -Token $token -Users $currentBatch -CompanyName $CompanyName -WhatIf:$WhatIf -Stats $stats -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Batch $($batchNum+1) results: $($batchResult.UpdatedCount) updated, $($batchResult.NoChangeCount) already correct, $($batchResult.ErrorCount) errors"
        
        if ($batchNum -lt $batches - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Company Attribute Update Process Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO CHANGES WERE MADE ===" -Type "WHATIF"
    }
    
    Write-Log "Overall Summary:"
    Write-Log "Total users processed: $totalUsers"
    Write-Log "Already had correct company: $($stats.NoChangeCount)"
    
    if ($WhatIf) {
        Write-Log "Would be updated: $($stats.UpdatedCount)" -Type "WHATIF"
    } else {
        Write-Log "Updated: $($stats.UpdatedCount)"
    }
    
    Write-Log "Errors: $($stats.ErrorCount)"
    
    $outputObject = [PSCustomObject][ordered]@{
        TotalUsers = $stats.TotalUsers
        UpdatedCount = $stats.UpdatedCount
        NoChangeCount = $stats.NoChangeCount
        ErrorCount = $stats.ErrorCount
        WhatIfMode = $WhatIf
        DurationMinutes = $duration.TotalMinutes
        BatchesProcessed = $batches
        CompanyName = $CompanyName
    }
    
    return $outputObject
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
}