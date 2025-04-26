<#
.SYNOPSIS
    Moves a Microsoft Teams channel and its content to another team using Microsoft Graph API.
    
.DESCRIPTION
    This Azure Automation runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    exports content from a source channel, creates a new channel in the target team in migration mode, 
    imports the content with preserved metadata, and completes the migration process.
    
    Note: This process creates a NEW channel in the target team with content from the source channel.
    It does NOT delete the source channel or modify it in any way.
    
.PARAMETER SourceTeamId
    The ID of the source team containing the channel to be moved.
    
.PARAMETER SourceChannelId
    The ID of the source channel to be moved.
    
.PARAMETER TargetTeamId
    The ID of the target team where the channel should be moved to.
    
.PARAMETER NewChannelName
    Optional. The name for the new channel. If not specified, the original channel name will be used.
    
.PARAMETER NewChannelDescription
    Optional. The description for the new channel. If not specified, the original channel description will be used.
    
.PARAMETER DateRangeStart
    Optional. The start date for message export (format: yyyy-MM-dd). If not specified, all messages will be included.
    
.PARAMETER DateRangeEnd
    Optional. The end date for message export (format: yyyy-MM-dd). If not specified, all messages will be included.
    
.PARAMETER BatchSize
    Optional. The number of messages to process in each batch. Default is 20 to avoid throttling.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.NOTES
    File Name: Move-TeamsChannel.ps1
    Author: Ryan Schultz
    Version: 1.0
    
    Required Microsoft Graph API permissions for the System-Assigned Managed Identity:
    - TeamworkMigration.All (Application)
    - Team.ReadBasic.All (Application)
    - Channel.ReadBasic.All (Application)
    - ChannelMessage.Read.All (Application)
    - Group.ReadWrite.All (Application)
#>

# Requires -Modules "Az.Accounts"

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SourceTeamId,
    
    [Parameter(Mandatory = $true)]
    [string]$SourceChannelId,
    
    [Parameter(Mandatory = $true)]
    [string]$TargetTeamId,
    
    [Parameter(Mandatory = $false)]
    [string]$NewChannelName = "",
    
    [Parameter(Mandatory = $false)]
    [string]$NewChannelDescription = "",
    
    [Parameter(Mandatory = $false)]
    [string]$DateRangeStart = "",
    
    [Parameter(Mandatory = $false)]
    [string]$DateRangeEnd = "",
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 20,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5
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
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            )
        } else {
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
        Write-Log "Failed to acquire Microsoft Graph token using Managed Identity: $_" -Type "ERROR"
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
        [int]$InitialBackoffSeconds = 5,
        [switch]$Raw = $false
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
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 20))
        }
    }
    
    while ($true) {
        try {
            if ($Raw) {
                return Invoke-WebRequest @params
            } else {
                return Invoke-RestMethod @params
            }
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
                $errorDetails = ""
                if ($_.ErrorDetails) {
                    $errorDetails = $_.ErrorDetails.Message
                } elseif ($_.Exception.Response) {
                    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    $errorDetails = $reader.ReadToEnd()
                }
                
                Write-Log "Graph API request failed with status code $statusCode. Error: $_. Details: $errorDetails" -Type "ERROR"
                throw $_
            }
        }
    }
}

function Get-TeamChannelDetails {
    param (
        [string]$Token,
        [string]$TeamId,
        [string]$ChannelId
    )
    
    try {
        Write-Log "Retrieving channel details for Team ID: $TeamId, Channel ID: $ChannelId"
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId"
        
        $channelDetails = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Retrieved channel details: '$($channelDetails.displayName)'"
        return $channelDetails
    }
    catch {
        Write-Log "Failed to retrieve channel details: $_" -Type "ERROR"
        throw "Failed to retrieve channel details: $_"
    }
}

function Get-ChannelMessages {
    param (
        [string]$Token,
        [string]$TeamId,
        [string]$ChannelId,
        [string]$DateRangeStart = "",
        [string]$DateRangeEnd = ""
    )
    
    try {
        Write-Log "Retrieving messages from channel '$ChannelId' in team '$TeamId'"
        
        $filter = ""
        if ($DateRangeStart -ne "" -and $DateRangeEnd -ne "") {
            $filter = "?`$filter=lastModifiedDateTime ge $($DateRangeStart)T00:00:00Z and lastModifiedDateTime le $($DateRangeEnd)T23:59:59Z"
        } elseif ($DateRangeStart -ne "") {
            $filter = "?`$filter=lastModifiedDateTime ge $($DateRangeStart)T00:00:00Z"
        } elseif ($DateRangeEnd -ne "") {
            $filter = "?`$filter=lastModifiedDateTime le $($DateRangeEnd)T23:59:59Z"
        }
        
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages$filter"
        
        $messages = @()
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        $messages += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of messages..."
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $response.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            $messages += $response.value
        }
        
        Write-Log "Retrieved $($messages.Count) messages from channel"
        
        $allMessages = @()
        foreach ($message in $messages) {
            $allMessages += $message
            
            $repliesUri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages/$($message.id)/replies"
            $repliesResponse = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $repliesUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($repliesResponse.value.Count -gt 0) {
                foreach ($reply in $repliesResponse.value) {
                    $reply | Add-Member -MemberType NoteProperty -Name "isReply" -Value $true
                    $reply | Add-Member -MemberType NoteProperty -Name "parentMessageId" -Value $message.id
                    $allMessages += $reply
                }
                
                while ($null -ne $repliesResponse.'@odata.nextLink') {
                    $repliesResponse = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $repliesResponse.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                    
                    foreach ($reply in $repliesResponse.value) {
                        $reply | Add-Member -MemberType NoteProperty -Name "isReply" -Value $true
                        $reply | Add-Member -MemberType NoteProperty -Name "parentMessageId" -Value $message.id
                        $allMessages += $reply
                    }
                }
            }
        }
        
        Write-Log "Retrieved a total of $($allMessages.Count) messages (including replies)"
        
        return $allMessages | Sort-Object -Property createdDateTime
    }
    catch {
        Write-Log "Failed to retrieve channel messages: $_" -Type "ERROR"
        throw "Failed to retrieve channel messages: $_"
    }
}

function Create-ChannelInMigrationMode {
    param (
        [string]$Token,
        [string]$TeamId,
        [string]$ChannelName,
        [string]$ChannelDescription,
        [string]$MembershipType = "standard",
        [datetime]$CreatedDateTime = (Get-Date).AddYears(-1)  # Default to 1 year ago
    )
    
    try {
        Write-Log "Creating channel '$ChannelName' in migration mode in team '$TeamId'"
        
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels"
        
        $body = @{
            "@microsoft.graph.channelCreationMode" = "migration"
            "displayName" = $ChannelName
            "description" = $ChannelDescription
            "membershipType" = $MembershipType
            "createdDateTime" = $CreatedDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        }
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $body -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Successfully created channel in migration mode. Channel ID: $($response.id)"
        return $response
    }
    catch {
        Write-Log "Failed to create channel in migration mode: $_" -Type "ERROR"
        throw "Failed to create channel in migration mode: $_"
    }
}

function Import-MessageToChannel {
    param (
        [string]$Token,
        [string]$TeamId,
        [string]$ChannelId,
        [PSCustomObject]$Message,
        [string]$ReplyToId = $null
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages"
        
        if ($ReplyToId) {
            $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages/$ReplyToId/replies"
        }
        
        $messageBody = @{
            "body" = @{
                "content" = $Message.body.content
                "contentType" = $Message.body.contentType
            }
            "createdDateTime" = $Message.createdDateTime
        }
        
        if ($Message.from -and $Message.from.user) {
            if ($Message.from.user.userIdentityType -eq "aadUser") {
                $messageBody["from"] = @{
                    "user" = @{
                        "id" = $Message.from.user.id
                        "displayName" = $Message.from.user.displayName
                        "userIdentityType" = "aadUser"
                    }
                }
            } else {
                $messageBody["from"] = @{
                    "user" = @{
                        "id" = $Message.from.user.id
                        "displayName" = $Message.from.user.displayName
                    }
                }
            }
        }
        
        if ($Message.subject) {
            $messageBody["subject"] = $Message.subject
        }
        
        if ($Message.importance) {
            $messageBody["importance"] = $Message.importance
        }
        
        if ($Message.attachments -and $Message.attachments.Count -gt 0) {
            $messageBody["attachments"] = $Message.attachments
        }
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $messageBody -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        return $response
    }
    catch {
        Write-Log "Failed to import message: $_" -Type "ERROR"
        throw "Failed to import message: $_"
    }
}

function Complete-ChannelMigration {
    param (
        [string]$Token,
        [string]$TeamId,
        [string]$ChannelId
    )
    
    try {
        Write-Log "Completing migration for channel ID: $ChannelId in team ID: $TeamId"
        
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/completeMigration"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body "{}" -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Migration completed successfully"
        return $true
    }
    catch {
        Write-Log "Failed to complete channel migration: $_" -Type "ERROR"
        throw "Failed to complete channel migration: $_"
    }
}

function Complete-TeamMigration {
    param (
        [string]$Token,
        [string]$TeamId
    )
    
    try {
        Write-Log "Completing migration for team ID: $TeamId"
        
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/completeMigration"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body "{}" -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Team migration completed successfully"
        return $true
    }
    catch {
        Write-Log "Failed to complete team migration: $_" -Type "ERROR"
        throw "Failed to complete team migration: $_"
    }
}

# Main script execution starts here
try {
    Write-Log "=== Teams Channel Migration Process Started ==="
    
    $startTime = Get-Date
    $token = Get-MsGraphToken
    
    $sourceChannelDetails = Get-TeamChannelDetails -Token $token -TeamId $SourceTeamId -ChannelId $SourceChannelId
    
    if ([string]::IsNullOrEmpty($NewChannelName)) {
        $NewChannelName = $sourceChannelDetails.displayName
    }
    
    if ([string]::IsNullOrEmpty($NewChannelDescription)) {
        $NewChannelDescription = $sourceChannelDetails.description
    }
    
    Write-Log "Source Channel: '$($sourceChannelDetails.displayName)' ($SourceChannelId)"
    Write-Log "Target Team: $TargetTeamId"
    Write-Log "New Channel Name: '$NewChannelName'"
    
    # Get the oldest message date to use for channel creation timestamp
    $allMessages = Get-ChannelMessages -Token $token -TeamId $SourceTeamId -ChannelId $SourceChannelId -DateRangeStart $DateRangeStart -DateRangeEnd $DateRangeEnd
    
    if ($allMessages.Count -eq 0) {
        Write-Log "No messages found in the source channel for the specified date range." -Type "WARNING"
        $oldestMessageDate = (Get-Date).AddYears(-1)  # Default to 1 year ago
    } else {
        $oldestMessageDate = ($allMessages | Sort-Object -Property createdDateTime | Select-Object -First 1).createdDateTime
        Write-Log "Oldest message date: $oldestMessageDate"
    }
    
    $newChannel = Create-ChannelInMigrationMode -Token $token -TeamId $TargetTeamId -ChannelName $NewChannelName -ChannelDescription $NewChannelDescription -MembershipType $sourceChannelDetails.membershipType -CreatedDateTime $oldestMessageDate
    
    $parentMessages = $allMessages | Where-Object { -not ($_.PSObject.Properties.Name -contains "isReply") -or -not $_.isReply }
    $replyMessages = $allMessages | Where-Object { $_.PSObject.Properties.Name -contains "isReply" -and $_.isReply }
    
    Write-Log "Processing $($parentMessages.Count) parent messages and $($replyMessages.Count) reply messages in batches of $BatchSize"
    
    $messageIdMapping = @{}
    
    for ($i = 0; $i -lt $parentMessages.Count; $i += $BatchSize) {
        $batchEnd = [Math]::Min($i + $BatchSize - 1, $parentMessages.Count - 1)
        $currentBatch = $parentMessages[$i..$batchEnd]
        
        Write-Log "Processing parent messages batch $($i / $BatchSize + 1) of $([Math]::Ceiling($parentMessages.Count / $BatchSize)) (messages $($i+1) to $($batchEnd+1) of $($parentMessages.Count))"
        
        foreach ($message in $currentBatch) {
            try {
                $newMessage = Import-MessageToChannel -Token $token -TeamId $TargetTeamId -ChannelId $newChannel.id -Message $message
                $messageIdMapping[$message.id] = $newMessage.id
                Write-Log "Imported message: $($message.id) -> $($newMessage.id)"
            }
            catch {
                Write-Log "Failed to import message $($message.id): $_" -Type "ERROR"
            }
            
            Start-Sleep -Milliseconds 200
        }
        
        if ($batchEnd -lt $parentMessages.Count - 1) {
            Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch..."
            Start-Sleep -Seconds $BatchDelaySeconds
        }
    }
    
    Write-Log "Processing reply messages..."
    $replyGroups = $replyMessages | Group-Object -Property parentMessageId
    
    foreach ($group in $replyGroups) {
        $parentId = $group.Name
        
        if ($messageIdMapping.ContainsKey($parentId)) {
            $newParentId = $messageIdMapping[$parentId]
            
            Write-Log "Processing $($group.Group.Count) replies for parent message $parentId -> $newParentId"
            
            for ($i = 0; $i -lt $group.Group.Count; $i += $BatchSize) {
                $batchEnd = [Math]::Min($i + $BatchSize - 1, $group.Group.Count - 1)
                $currentBatch = $group.Group[$i..$batchEnd]
                
                foreach ($reply in $currentBatch) {
                    try {
                        $newReply = Import-MessageToChannel -Token $token -TeamId $TargetTeamId -ChannelId $newChannel.id -Message $reply -ReplyToId $newParentId
                        Write-Log "Imported reply: $($reply.id) -> $($newReply.id)"
                    }
                    catch {
                        Write-Log "Failed to import reply $($reply.id): $_" -Type "ERROR"
                    }
                    
                    Start-Sleep -Milliseconds 200
                }
                
                if ($batchEnd -lt $group.Group.Count - 1) {
                    Write-Log "Waiting $BatchDelaySeconds seconds before processing next batch of replies..."
                    Start-Sleep -Seconds $BatchDelaySeconds
                }
            }
        }
        else {
            Write-Log "Parent message ID $parentId not found in mapping. Cannot import replies." -Type "WARNING"
        }
    }
    
    Complete-ChannelMigration -Token $token -TeamId $TargetTeamId -ChannelId $newChannel.id
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Teams Channel Migration Process Completed ==="
    Write-Log "Migration completed in $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Source Channel: $($sourceChannelDetails.displayName) (Team: $SourceTeamId)"
    Write-Log "Target Channel: $($newChannel.displayName) (Team: $TargetTeamId, Channel: $($newChannel.id))"
    Write-Log "Messages Processed: $($allMessages.Count) ($($parentMessages.Count) parent messages, $($replyMessages.Count) replies)"
    
    $result = [PSCustomObject]@{
        SourceTeamId = $SourceTeamId
        SourceChannelId = $SourceChannelId
        SourceChannelName = $sourceChannelDetails.displayName
        TargetTeamId = $TargetTeamId
        TargetChannelId = $newChannel.id
        TargetChannelName = $newChannel.displayName
        MessagesProcessed = $allMessages.Count
        ParentMessagesProcessed = $parentMessages.Count
        ReplyMessagesProcessed = $replyMessages.Count
        DurationMinutes = $duration.TotalMinutes
        NewChannelUrl = $newChannel.webUrl
    }
    
    return $result
}
catch {
    Write-Log "Channel migration failed: $_" -Type "ERROR"
    throw $_
}