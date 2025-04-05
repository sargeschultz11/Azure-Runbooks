<#
.SYNOPSIS
    Generates a report of discovered apps from Microsoft Intune and uploads it to a SharePoint document library.

.DESCRIPTION
    This Azure Runbook script connects to the Microsoft Graph API using client credentials (App Registration),
    retrieves all discovered apps from Intune with their installation counts, exports the data
    to an Excel file, and uploads the file to a specified SharePoint document library.
    It supports batch processing and implements throttling detection with retry logic.

.PARAMETER TenantId
    The Azure AD tenant ID. If not provided, will be retrieved from Automation variables.

.PARAMETER ClientId
    The App Registration's client ID. If not provided, will be retrieved from Automation variables.

.PARAMETER ClientSecret
    The App Registration's client secret. If not provided, will be retrieved from Automation variables.
    
.PARAMETER SharePointSiteId
    The ID of the SharePoint site where the report will be uploaded.

.PARAMETER SharePointDriveId
    The ID of the document library drive where the report will be uploaded.

.PARAMETER FolderPath
    Optional. The folder path within the document library where the report will be uploaded.
    If not specified, the file will be uploaded to the root of the document library.

.PARAMETER BatchSize
    Optional. Number of records to process in each batch. Default is 100.

.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.

.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.

.NOTES
    File Name: Get-IntuneDiscoveredAppsReport.ps1
    Author: Ryan Schultz
    Version: 1.0
    Created: 2025-04-04

    Requires -Modules ImportExcel

#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$UseManagedIdentity = $true,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointDriveId,
    
    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "",
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 100,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl
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
        Write-Log "Acquiring Microsoft Graph token using Managed Identity..."
        
        Connect-AzAccount -Identity | Out-Null
        
        $azAccountsModule = Get-Module -Name Az.Accounts -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        Write-Log "Using Az.Accounts module version: $($azAccountsModule.Version)"
        
        $token = $null
        
        if ($azAccountsModule.Version.Major -ge 2) {
            try {
                $tokenResult = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
                
                if ($null -ne $tokenResult -and $null -ne $tokenResult.Token) {
                    $token = $tokenResult.Token
                    Write-Log "Successfully acquired token using standard approach"
                }
            }
            catch {
                Write-Log "Failed to get token using standard approach: $($_.Exception.Message)" -Type "WARNING"
            }
        }
        
        if ([string]::IsNullOrEmpty($token)) {
            Write-Log "Attempting alternative token acquisition approach"
            
            if (-not (Get-Module -Name Az.Accounts)) {
                Import-Module Az.Accounts -ErrorAction Stop
            }
            
            $context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
            $tokenCache = $context.TokenCache
            $cachedTokens = $tokenCache.ReadItems() | Where-Object { $_.Resource -eq "https://graph.microsoft.com" }
            
            if ($cachedTokens -and $cachedTokens.Count -gt 0) {
                $latestToken = $cachedTokens | Sort-Object ExpiresOn -Descending | Select-Object -First 1
                $token = $latestToken.AccessToken
                Write-Log "Successfully acquired token from token cache"
            }
        }
        
        if ([string]::IsNullOrEmpty($token)) {
            Write-Log "Attempting final fallback token acquisition approach" -Type "WARNING"
            
            $armToken = Get-AzAccessToken
            
            if ($null -ne $armToken -and $null -ne $armToken.Token) {
                $graphToken = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$((Get-AzContext).Tenant.Id)/oauth2/v2.0/token" -Body @{
                    grant_type    = "client_credentials"
                    client_id     = $env:IDENTITY_CLIENT_ID 
                    scope         = "https://graph.microsoft.com/.default"
                    client_secret = $env:IDENTITY_CLIENT_SECRET 
                } -ContentType "application/x-www-form-urlencoded"
                
                if ($graphToken -and $graphToken.access_token) {
                    $token = $graphToken.access_token
                    Write-Log "Successfully acquired token using token exchange approach"
                }
            }
        }
        
        if ([string]::IsNullOrEmpty($token)) {
            throw "Failed to acquire valid token from managed identity after trying multiple approaches"
        }
        
        if ($token -notmatch '\..*\.') {
            $tokenPreview = if ($token.Length -gt 20) { $token.Substring(0, 20) + "..." } else { $token }
            throw "Acquired token does not appear to be a valid JWT. Token preview: $tokenPreview"
        }
        
        Write-Log "Successfully acquired Microsoft Graph API token via Managed Identity"
        return $token
    }
    catch {
        Write-Log "Failed to acquire Microsoft Graph token using Managed Identity: $($_.Exception.Message)" -Type "ERROR"
        throw "Authentication failed: $($_.Exception.Message)"
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

function Get-IntuneDiscoveredApps {
    param (
        [string]$Token,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5,
        [int]$BatchSize = 100
    )
    
    try {
        Write-Log "Retrieving discovered apps from Intune..."
        
        $discoveredApps = @()
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?`$top=$BatchSize"
        
        $count = 0
        $batchCount = 0
        
        do {
            $batchCount++
            Write-Log "Retrieving batch $batchCount of discovered apps..."
            
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            if ($response.value.Count -gt 0) {
                $discoveredApps += $response.value
                $count += $response.value.Count
                Write-Log "Retrieved $($response.value.Count) apps in this batch, total count: $count"
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($null -ne $uri)
        
        Write-Log "Retrieved a total of $($discoveredApps.Count) discovered apps"
        
        return $discoveredApps
    }
    catch {
        Write-Log "Failed to retrieve discovered apps: $_" -Type "ERROR"
        throw "Failed to retrieve discovered apps: $_"
    }
}

function Export-DataToExcel {
    param (
        [array]$Data,
        [string]$FilePath
    )
    
    try {
        Write-Log "Exporting data to Excel file: $FilePath"
        
        if (-not (Get-Module -Name ImportExcel)) {
            Import-Module ImportExcel -ErrorAction Stop
        }
        
        $reportInfo = [PSCustomObject]@{
            'Report Generated' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            'Generated By'     = $env:COMPUTERNAME
            'Number of Apps'   = $Data.Count
        }
        
        $excelParams = @{
            Path          = $FilePath
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            AutoFilter    = $true
            WorksheetName = "Discovered Apps"
            TableName     = "DiscoveredApps"
            PassThru      = $true
        }
        
        $excel = $Data | Select-Object @{Name = 'Application Name'; Expression = { $_.displayName } }, 
        @{Name = 'Publisher'; Expression = { $_.publisher } }, 
        @{Name = 'Version'; Expression = { $_.version } }, 
        @{Name = 'Device Count'; Expression = { $_.deviceCount } }, 
        @{Name = 'Platform'; Expression = { $_.platform } },
        @{Name = 'Size in Bytes'; Expression = { $_.sizeInByte } }, 
        @{Name = 'App ID'; Expression = { $_.id } } | 
        Export-Excel @excelParams
        
        $summarySheet = $excel.Workbook.Worksheets.Add("Summary")
        $summarySheet.Cells["A1"].Value = "Report Summary"
        $summarySheet.Cells["A1:B1"].Merge = $true
        $summarySheet.Cells["A1:B1"].Style.Font.Bold = $true
        $summarySheet.Cells["A1:B1"].Style.Font.Size = 14
        
        $row = 3
        $summarySheet.Cells["A$row"].Value = "Report Generated"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Report Generated'
        
        $row++
        $summarySheet.Cells["A$row"].Value = "Generated By"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Generated By'
        
        $row++
        $summarySheet.Cells["A$row"].Value = "Number of Apps"
        $summarySheet.Cells["B$row"].Value = $reportInfo.'Number of Apps'
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Top Publishers"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $publisherSummary = $Data | Group-Object -Property publisher | 
        Sort-Object -Property Count -Descending | 
        Select-Object -First 10
        
        $summarySheet.Cells["A$row"].Value = "Publisher"
        $summarySheet.Cells["B$row"].Value = "App Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($publisher in $publisherSummary) {
            $summarySheet.Cells["A$row"].Value = if ([string]::IsNullOrEmpty($publisher.Name)) { "(Unknown)" } else { $publisher.Name }
            $summarySheet.Cells["B$row"].Value = $publisher.Count
            $row++
        }
        
        $row += 2
        $summarySheet.Cells["A$row"].Value = "Platform Summary"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $row++
        
        $platformSummary = $Data | Where-Object { ![string]::IsNullOrEmpty($_.platform) } | 
        Group-Object -Property platform | 
        Sort-Object -Property Count -Descending
        
        $summarySheet.Cells["A$row"].Value = "Platform"
        $summarySheet.Cells["B$row"].Value = "App Count"
        $summarySheet.Cells["A$row"].Style.Font.Bold = $true
        $summarySheet.Cells["B$row"].Style.Font.Bold = $true
        $row++
        
        foreach ($platform in $platformSummary) {
            $summarySheet.Cells["A$row"].Value = $platform.Name
            $summarySheet.Cells["B$row"].Value = $platform.Count
            $row++
        }
        
        $summarySheet.Column(1).AutoFit()
        $summarySheet.Column(2).AutoFit()
        
        try {
            $excel.Workbook.Worksheets[0].View.TabSelected = $false
            $summarySheet.View.TabSelected = $true
            $excel.Workbook.View.ActiveTab = 1
        }
        catch {
            Write-Log "Could not set the active sheet, but this is not critical for report generation" -Type "WARNING"
        }
        
        $excel.Save()
        $excel.Dispose()
        
        Write-Log "Excel file created successfully at: $FilePath"
    }
    catch {
        Write-Log "Failed to export data to Excel: $_" -Type "ERROR"
        throw "Failed to export data to Excel: $_"
    }
}

function Upload-FileToSharePoint {
    param (
        [string]$Token,
        [string]$SiteId,
        [string]$DriveId,
        [string]$FolderPath,
        [string]$FilePath,
        [string]$FileName,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Uploading file to SharePoint..."
        
        $fileContent = [System.IO.File]::ReadAllBytes($FilePath)
        $fileSize = $fileContent.Length
        
        $uploadPath = if ([string]::IsNullOrEmpty($FolderPath)) {
            $FileName
        }
        else {
            "$FolderPath/$FileName"
        }
        
        $uploadUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root:/$uploadPath`:/content"
        
        Write-Log "Uploading file to: $uploadUri"
        Write-Log "File size: $fileSize bytes"
        
        $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uploadUri -Method "PUT" -Body $fileContent -ContentType "application/octet-stream" -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "File uploaded successfully. WebUrl: $($response.webUrl)"
        return $response
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Type "ERROR"
        throw "Failed to upload file to SharePoint: $_"
    }
}

function Send-TeamsNotification {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebhookUrl,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ReportData
    )
    
    try {
        Write-Log "Sending notification to Microsoft Teams..."
        
        $executionTime = [math]::Round($ReportData.ExecutionTimeMinutes, 2)
        
        $adaptiveCard = @{
            type        = "message"
            attachments = @(
                @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    contentUrl  = $null
                    content     = @{
                        "$schema" = "http://adaptivecards.io/schemas/adaptive-card.json"
                        type      = "AdaptiveCard"
                        version   = "1.2"
                        msTeams   = @{
                            width = "full"
                        }
                        body      = @(
                            @{
                                type   = "TextBlock"
                                size   = "Large"
                                weight = "Bolder"
                                text   = "Intune Discovered Apps Report"
                                wrap   = $true
                                color  = "Default"
                            },
                            @{
                                type     = "TextBlock"
                                spacing  = "None"
                                text     = "Report generated on $($ReportData.Timestamp)"
                                wrap     = $true
                                isSubtle = $true
                            },
                            @{
                                type  = "FactSet"
                                facts = @(
                                    @{
                                        title = "Report Name:"
                                        value = $ReportData.ReportName
                                    },
                                    @{
                                        title = "Applications Discovered:"
                                        value = $ReportData.AppsCount.ToString()
                                    },
                                    @{
                                        title = "Execution Time:"
                                        value = "$executionTime minutes"
                                    }
                                )
                            }
                        )
                        actions   = @(
                            @{
                                type  = "Action.OpenUrl"
                                title = "View Report"
                                url   = $ReportData.ReportUrl
                            }
                        )
                    }
                }
            )
        }
        
        $body = ConvertTo-Json -InputObject $adaptiveCard -Depth 20
        
        $params = @{
            Uri         = $WebhookUrl
            Method      = "POST"
            Body        = $body
            ContentType = "application/json"
        }
        
        $response = Invoke-RestMethod @params
        
        Write-Log "Teams notification sent successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send Teams notification: $_" -Type "WARNING"
        return $false
    }
}

try {
    $startTime = Get-Date
    Write-Log "=== Intune Discovered Apps Report Generation Started ==="
    
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "ImportExcel module not found. Installing..." -Type "WARNING"
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
    
    $token = Get-MsGraphToken
    
    $discoveredApps = Get-IntuneDiscoveredApps -Token $token -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds -BatchSize $BatchSize
    
    if ($discoveredApps.Count -eq 0) {
        Write-Log "No discovered apps found in Intune" -Type "WARNING"
        return
    }
    
    $currentDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $reportName = "Intune_Discovered_Apps_Report_$currentDate.xlsx"
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $reportName)
    
    Export-DataToExcel -Data $discoveredApps -FilePath $tempPath
    
    if (Test-Path -Path $tempPath) {
        $fileInfo = Get-Item -Path $tempPath
        Write-Log "Excel file created with size: $($fileInfo.Length) bytes"
        
        if ($fileInfo.Length -lt 1000) {
            Write-Log "Warning: Excel file appears to be very small, which might indicate a formatting issue" -Type "WARNING"
        }
    }
    
    $uploadResult = Upload-FileToSharePoint -Token $token -SiteId $SharePointSiteId -DriveId $SharePointDriveId -FolderPath $FolderPath -FilePath $tempPath -FileName $reportName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    
    Remove-Item -Path $tempPath -Force
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "=== Intune Discovered Apps Report Generation Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Report URL: $($uploadResult.webUrl)"
    
    $result = [PSCustomObject]@{
        ReportName           = $reportName
        AppsCount            = $discoveredApps.Count
        ReportUrl            = $uploadResult.webUrl
        ExecutionTimeMinutes = $duration.TotalMinutes
        Timestamp            = $currentDate
    }
    
    if (-not [string]::IsNullOrEmpty($TeamsWebhookUrl)) {
        $notificationSent = Send-TeamsNotification -WebhookUrl $TeamsWebhookUrl -ReportData $result
        if ($notificationSent) {
            $result | Add-Member -MemberType NoteProperty -Name "NotificationSent" -Value $true
        }
        else {
            $result | Add-Member -MemberType NoteProperty -Name "NotificationSent" -Value $false
        }
    }
    
    return $result
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    
    if ($tempPath -and (Test-Path -Path $tempPath)) {
        Remove-Item -Path $tempPath -Force
    }
    
    throw $_
}
finally {
    Write-Log "Script execution completed"
}