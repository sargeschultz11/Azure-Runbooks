<#
.SYNOPSIS
    Monitors Intune Apple token and certificate expiration dates and sends notifications.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves expiration information for various Intune tokens (Apple Push Notification service certificates,
    VPP tokens, DEP tokens), and sends Teams notifications for tokens approaching expiration.
    
.PARAMETER WarningThresholdDays
    The number of days before expiration to start sending warning notifications.
    Default is 30 days.
    
.PARAMETER TeamsWebhookUrl
    Microsoft Teams webhook URL for sending notifications about token status.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually send notifications.
    
.NOTES
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-17
    
    Required Graph API Permissions for Managed Identity:
    - DeviceManagementServiceConfig.Read.All
    - DeviceManagementConfiguration.Read.All
    - DeviceManagementApps.Read.All
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$WarningThresholdDays = 30,
    
    [Parameter(Mandatory = $true)]
    [string]$TeamsWebhookUrl,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

Write-Output "=== Intune Apple Token Monitor Started ==="
Write-Output "Warning threshold: $WarningThresholdDays days"
$startTime = Get-Date
$notificationThreshold = (Get-Date).AddDays($WarningThresholdDays)
$tokenCollection = @()
$expiringTokens = @()

# Connect to Microsoft Graph using Managed Identity
try {
    Write-Output "Connecting to Microsoft Graph using Managed Identity..."
    Connect-AzAccount -Identity | Out-Null
    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
    
    if ([string]::IsNullOrEmpty($token)) {
        throw "Failed to acquire token - token is empty"
    }
    Write-Output "Successfully connected to Microsoft Graph"
}
catch {
    Write-Output "Failed to connect to Microsoft Graph: $_"
    throw "Authentication failed: $_"
}

# Get organization information
try {
    Write-Output "Retrieving organization information..."
    $uri = "https://graph.microsoft.com/v1.0/organization"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $orgResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $orgDomain = ($orgResponse.value | 
                 Select-Object -ExpandProperty verifiedDomains | 
                 Where-Object { $_.isInitial } | 
                 Select-Object -ExpandProperty name)
    Write-Output "Organization domain: $orgDomain"
}
catch {
    Write-Output "Failed to retrieve organization information: $_"
    $orgDomain = "your organization"
}

# Check Apple Push Notification Certificate
try {
    Write-Output "Checking Apple Push Notification Certificate..."
    $uri = "https://graph.microsoft.com/beta/deviceManagement/applePushNotificationCertificate"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $applePushCert = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    if ($applePushCert) {
        Write-Output "Found Apple Push Notification Certificate for $($applePushCert.appleIdentifier)"
        Write-Output "Certificate expires: $($applePushCert.expirationDateTime)"
        $expirationDate = [datetime]$applePushCert.expirationDateTime
        $daysLeft = ($expirationDate - (Get-Date)).Days
        Write-Output "Days until expiration: $daysLeft"
        $status = if ($daysLeft -le 0) { "Expired" } 
                 elseif ($daysLeft -le 7) { "Critical" } 
                 elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                 else { "OK" }
        $tokenInfo = [PSCustomObject]@{
            TokenType = "Apple Push Notification Certificate"
            Name = $applePushCert.appleIdentifier
            ExpirationDate = $expirationDate
            DaysUntilExpiration = $daysLeft
            Status = $status
        }
        $tokenCollection += $tokenInfo
        if ($notificationThreshold -ge $expirationDate) {
            Write-Output "Certificate will expire soon! Adding to notification list."
            $expiringTokens += $tokenInfo
        }
    }
    else {
        Write-Output "No Apple Push Notification Certificate found."
    }
}
catch {
    Write-Output "Error checking Apple Push Notification Certificate: $_"
}

# Check Apple VPP Tokens
try {
    Write-Output "Checking Apple VPP Tokens..."
    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/vppTokens"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    $vppResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $vppTokens = $vppResponse.value
    if ($vppTokens -and $vppTokens.Count -gt 0) {
        Write-Output "Found $($vppTokens.Count) VPP tokens"
        foreach ($vppToken in $vppTokens) {
            Write-Output "Processing VPP token: $($vppToken.organizationName) - $($vppToken.appleId)"
            Write-Output "Token expires: $($vppToken.expirationDateTime)"
            $expirationDate = [datetime]$vppToken.expirationDateTime
            $daysLeft = ($expirationDate - (Get-Date)).Days
            Write-Output "Days until expiration: $daysLeft"
            $status = if ($daysLeft -le 0) { "Expired" } 
                     elseif ($daysLeft -le 7) { "Critical" } 
                     elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                     else { "OK" }
            $tokenInfo = [PSCustomObject]@{
                TokenType = "Apple VPP Token"
                Name = "$($vppToken.organizationName): $($vppToken.appleId)"
                ExpirationDate = $expirationDate
                DaysUntilExpiration = $daysLeft
                Status = $status
            }
            $tokenCollection += $tokenInfo
            if ($notificationThreshold -ge $expirationDate) {
                Write-Output "VPP token will expire soon! Adding to notification list."
                $expiringTokens += $tokenInfo
            }
        }
    }
    else {
        Write-Output "No Apple VPP tokens found."
    }
}
catch {
    Write-Output "Error checking Apple VPP tokens: $_"
}

# Check Apple DEP Tokens
try {
    Write-Output "Checking Apple DEP Tokens..."
    
    $uri = "https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings"
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    
    $depResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $depTokens = $depResponse.value
    if ($depTokens -and $depTokens.Count -gt 0) {
        Write-Output "Found $($depTokens.Count) DEP tokens"
        foreach ($depToken in $depTokens) {
            Write-Output "Processing DEP token: $($depToken.tokenName) - $($depToken.appleIdentifier)"
            Write-Output "Token expires: $($depToken.tokenExpirationDateTime)"
            if (-not [string]::IsNullOrEmpty($depToken.tokenExpirationDateTime)) {
                $expirationDate = [datetime]$depToken.tokenExpirationDateTime
                $daysLeft = ($expirationDate - (Get-Date)).Days
                Write-Output "Days until expiration: $daysLeft"
                
                $status = if ($daysLeft -le 0) { "Expired" } 
                         elseif ($daysLeft -le 7) { "Critical" } 
                         elseif ($daysLeft -le $WarningThresholdDays) { "Warning" } 
                         else { "OK" }
                
                $tokenInfo = [PSCustomObject]@{
                    TokenType = "Apple DEP Token"
                    Name = "$($depToken.tokenName): $($depToken.appleIdentifier)"
                    ExpirationDate = $expirationDate
                    DaysUntilExpiration = $daysLeft
                    Status = $status
                }
                $tokenCollection += $tokenInfo
                if ($notificationThreshold -ge $expirationDate) {
                    Write-Output "DEP token will expire soon! Adding to notification list."
                    $expiringTokens += $tokenInfo
                }
            }
            else {
                Write-Output "DEP token $($depToken.tokenName) has no expiration date."
            }
        }
    }
    else {
        Write-Output "No Apple DEP tokens found."
    }
}
catch {
    Write-Output "Error checking Apple DEP tokens: $_"
}

# Send Teams notification if there are expiring tokens
if ($expiringTokens.Count -gt 0) {
    try {
        Write-Output "Sending Teams notification for $($expiringTokens.Count) expiring tokens..."
        if ($WhatIf) {
            Write-Output "WHATIF: Would send Teams notification about expiring tokens"
        }
        else {
            $attachments = @()
            foreach ($token in $expiringTokens) {
                $cardAttachment = @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    content = @{
                        type = "AdaptiveCard"
                        version = "1.0"
                        body = @(
                            @{
                                type = "TextBlock"
                                text = "Intune Apple Token Expiration Alert"
                                size = "Large"
                                weight = "Bolder"
                            },
                            @{
                                type = "TextBlock"
                                text = "Token requires immediate attention:"
                                wrap = $true
                            },
                            @{
                                type = "FactSet"
                                facts = @(
                                    @{
                                        title = "Token Type"
                                        value = $token.TokenType
                                    },
                                    @{
                                        title = "Token Name"
                                        value = $token.Name
                                    },
                                    @{
                                        title = "Expiration Date"
                                        value = $token.ExpirationDate.ToString("yyyy-MM-dd")
                                    },
                                    @{
                                        title = "Days Remaining"
                                        value = "$($token.DaysUntilExpiration)"
                                    },
                                    @{
                                        title = "Status"
                                        value = $token.Status
                                    }
                                )
                            },
                            @{
                                type = "TextBlock"
                                text = "Please take action to renew this token before it expires."
                                wrap = $true
                            }
                        )
                    }
                }
                
                $attachments += $cardAttachment
            }
            
            $message = @{
                attachments = $attachments
            }
            $jsonBody = ConvertTo-Json -InputObject $message -Depth 10
            Write-Output "Sending JSON payload:"
            Write-Output $jsonBody
            $params = @{
                Uri = $TeamsWebhookUrl
                Method = "POST"
                Body = $jsonBody
                ContentType = "application/json"
                UseBasicParsing = $true
            }
            
            $teamsRequest = Invoke-WebRequest @params
            Write-Output "Teams notification sent successfully"
            Write-Output "Response: $($teamsRequest.StatusCode) $($teamsRequest.StatusDescription)"
        }
    }
    catch {
        Write-Output "Error sending Teams notification: $_"
        
        if ($null -ne $_.Exception) {
            Write-Output "Exception type: $($_.Exception.GetType().FullName)"
            
            if ($null -ne $_.Exception.Response) {
                Write-Output "Status code: $($_.Exception.Response.StatusCode)"
                
                try {
                    $responseStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    Write-Output "Response body: $responseBody"
                }
                catch {
                    Write-Output "Could not read response body: $_"
                }
            }
            
            if ($_.Exception.InnerException) {
                Write-Output "Inner exception: $($_.Exception.InnerException.Message)"
            }
        }
    }
}
else {
    Write-Output "No expiring tokens found. No notifications sent."
}

# Generate summary
$healthyTokens = $tokenCollection | Where-Object { $_.Status -eq "OK" }
$warningTokens = $tokenCollection | Where-Object { $_.Status -eq "Warning" }
$criticalTokens = $tokenCollection | Where-Object { $_.Status -eq "Critical" -or $_.Status -eq "Expired" }

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Output "=== Intune Apple Token Monitor Summary ==="
Write-Output "Total tokens checked: $($tokenCollection.Count)"
Write-Output "Healthy tokens: $($healthyTokens.Count)"
Write-Output "Warning tokens: $($warningTokens.Count)"
Write-Output "Critical/expired tokens: $($criticalTokens.Count)"
Write-Output "Execution time: $($duration.TotalMinutes.ToString("0.00")) minutes"

$result = [PSCustomObject]@{
    TotalTokensChecked = $tokenCollection.Count
    HealthyTokens = $healthyTokens.Count
    WarningTokens = $warningTokens.Count
    CriticalTokens = $criticalTokens.Count
    ExpiringTokens = $expiringTokens.Count
    ExecutionTimeMinutes = $duration.TotalMinutes
    TokenCollection = $tokenCollection
    ExpiringTokenDetails = $expiringTokens
}

return $result