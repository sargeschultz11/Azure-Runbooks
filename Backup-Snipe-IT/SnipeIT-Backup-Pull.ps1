<# 
.SYNOPSIS
  Azure Automation runbook: list Snipe-IT backups (rows[]), download newest, upload to Azure Blob via SAS.

.REQUIREMENTS
  PowerShell 7 runbook
  Encrypted Automation Variables:
    - SnipeItApiToken          : Snipe-IT API Bearer token
    - SnipeItContainerSasUrl   : FULL container SAS URL (Write+Create)
#>

param(
  [Parameter(Mandatory=$true)][string]$SnipeItBaseUrl,   # e.g. https://snipe-it.your-snipe-domain.net
  [string]$SnipeItTokenVar = "SnipeItApiToken",
  [string]$ContainerSasVar = "SnipeItContainerSasUrl",
  [string]$BlobPrefix      = "snipeit-backup"
)

# --- Secrets
try {
  $apiToken        = Get-AutomationVariable -Name $SnipeItTokenVar
  $containerSasUrl = Get-AutomationVariable -Name $ContainerSasVar
} catch {
  throw "Failed to read Automation Variables. Ensure '$SnipeItTokenVar' and '$ContainerSasVar' exist and are Encrypted=On. $_"
}
if ([string]::IsNullOrWhiteSpace($apiToken))        { throw "Missing API token in '$SnipeItTokenVar'." }
if ([string]::IsNullOrWhiteSpace($containerSasUrl)) { throw "Missing container SAS in '$ContainerSasVar'." }

# --- API endpoints
$apiRoot     = "$SnipeItBaseUrl/api/v1"
$listUrl     = "$apiRoot/settings/backups"
$downloadUrl = "$apiRoot/settings/backups/download"   # append /{file}

# --- HTTP headers
$jsonHeaders = @{ "Authorization" = "Bearer $apiToken"; "Accept" = "application/json" }
$zipHeaders  = @{ "Authorization" = "Bearer $apiToken"; "Accept"  = "application/zip"  }

# --- 1) List backups
Write-Output "Listing backups from $listUrl"
try {
  $listResp = Invoke-RestMethod -Uri $listUrl -Headers $jsonHeaders -Method GET -TimeoutSec 600
} catch {
  throw "Failed to list backups. $_"
}

# Defensive visibility without leaking secrets
$topLevelKeys = ($listResp | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -join ','
Write-Output "Response keys: $topLevelKeys"

# --- Normalize to an array using documented shape { total, rows: [] }
$items = $null
if ($listResp -and $listResp.PSObject.Properties.Name -contains 'rows') {
  $items = @($listResp.rows)
} elseif ($listResp.PSObject.Properties.Name -contains 'data') {
  $items = @($listResp.data)
} elseif ($listResp.PSObject.Properties.Name -contains 'files') {
  $items = @($listResp.files)
} elseif ($listResp -is [System.Collections.IEnumerable]) {
  $items = @($listResp)
}

if (-not $items -or $items.Count -eq 0) { throw "No backups found via API listing." }

# --- 2) Choose newest item
# Prefer numeric epoch 'modified_value' if present, else RFC date in 'modified_display',
# else derive timestamp from filename.
$projected = foreach ($it in $items) {
  # Pull filename from common fields
  $fn = $it.filename
  if (-not $fn) { $fn = $it.file ?? $it.name ?? $it.basename ?? $it.path }

  # Build sort key
  $sortKey = [DateTime]::MinValue

  if ($it.PSObject.Properties.Name -contains 'modified_value' -and $it.modified_value) {
    # Snipe-IT returns epoch seconds
    $epoch = [double]$it.modified_value
    $sortKey = [DateTimeOffset]::FromUnixTimeSeconds([int64]$epoch).UtcDateTime
  }
  elseif ($it.PSObject.Properties.Name -contains 'modified_display' -and $it.modified_display) {
    [DateTime]$dtOut = $null
    if ([DateTime]::TryParse($it.modified_display, [ref]$dtOut)) {
      $sortKey = $dtOut.ToUniversalTime()
    }
  }
  elseif ($fn -and ($fn -match '(\d{4})[^\d]?(\d{2})[^\d]?(\d{2})[^\d]?[_-]?(\d{2})(\d{2})(\d{2})')) {
    $sortKey = Get-Date -Date ("{0}-{1}-{2}T{3}:{4}:{5}Z" -f $Matches[1],$Matches[2],$Matches[3],$Matches[4],$Matches[5],$Matches[6])
  }

  [PSCustomObject]@{ Filename = $fn; SortKey = $sortKey }
}

$best = $projected | Sort-Object SortKey -Descending | Select-Object -First 1
$latestFile = $best.Filename
if ([string]::IsNullOrWhiteSpace($latestFile)) { throw "Backups were listed but no filename field was found." }
Write-Output "Newest backup detected: $latestFile (UTC sort key: $($best.SortKey.ToString('s'))Z)"

# --- 3) Download newest ZIP
$stamp = (Get-Date).ToUniversalTime().ToString("yyyyMMdd_HHmmss")
$tmp   = Join-Path ([System.IO.Path]::GetTempPath()) ("snipeit_{0}.zip" -f $stamp)
$dlUrl = "$downloadUrl/$([Uri]::EscapeDataString($latestFile))"

Write-Output "Downloading $latestFile"
try {
  Invoke-WebRequest -Uri $dlUrl -Headers $zipHeaders -Method GET -TimeoutSec 1800 -OutFile $tmp | Out-Null
} catch {
  throw "Download failed for '$latestFile'. $_"
}
if (-not (Test-Path $tmp)) { throw "Backup ZIP not found at $tmp after download." }
$size = (Get-Item $tmp).Length
Write-Output ("Downloaded {0:N0} bytes" -f $size)

# --- 4) Build blob URL from container SAS (do not log SAS)
try {
  $uri           = [System.Uri]$containerSasUrl
  if (-not $uri.Query) { throw "Container SAS URL lacks a query string." }
  $containerBase = $uri.GetLeftPart([System.UriPartial]::Path).TrimEnd('/')
  $sasQuery      = $uri.Query.TrimStart('?')
  $blobFile      = "$BlobPrefix-$stamp-$latestFile"
  $blobUrl       = "{0}/{1}?{2}" -f $containerBase, [Uri]::EscapeDataString($blobFile), $sasQuery
} catch {
  throw "Invalid Container SAS URL in '$ContainerSasVar'. $_"
}

# --- 5) Upload via HTTPS PUT
Write-Output "Uploading to Azure Blob (sanitized)…"
$headers = @{ "x-ms-blob-type" = "BlockBlob"; "x-ms-version" = "2021-12-02" }
try {
  Invoke-WebRequest -Uri $blobUrl -Method PUT -Headers $headers -InFile $tmp -TimeoutSec 1800 | Out-Null
  Write-Output "Upload complete."
} catch {
  throw "Blob upload failed. $_"
} finally {
  Remove-Item -Path $tmp -ErrorAction SilentlyContinue
}

Write-Output "Success. Blob: $blobFile"