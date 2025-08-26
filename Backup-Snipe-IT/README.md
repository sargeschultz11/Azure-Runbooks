# Snipe-IT to Azure Blob Backup

Daily offsite backup pipeline for Snipe-IT. The runbook:

1. Lists available backups from the Snipe-IT API (`/settings/backups`).
2. Selects the most recent backup (by `modified_value` if present).
3. Downloads it from `/settings/backups/download/{file}`.
4. Uploads it to an Azure Blob container via SAS.

Runs in Azure Automation (PowerShell 7).

---

## Prerequisites

1. **Snipe-IT host**

   * Ensure scheduled backups are generated (`php artisan snipeit:backup` or Laravel scheduler).
   * Create an **API token** with `settings/backups` read access.

2. **Azure Storage**

   * Target container exists.
   * SAS with **Write** and **Create** permissions (and optionally Add).
   * SAS should be **short-lived** and rotated regularly.

3. **Azure Automation (PowerShell 7)**

   * Import `SnipeIT-Backup-Pull.ps1` as a runbook (PowerShell 7.2).
   * Create encrypted Automation Variables:

     * `SnipeItApiToken` → Snipe-IT API token
     * `SnipeItContainerSasUrl` → full container SAS URL

---

## Parameters

| Name              | Type   | Required | Description                                                                  |
| ----------------- | ------ | -------: | ---------------------------------------------------------------------------- |
| `SnipeItBaseUrl`  | string |      Yes | Base URL without trailing slash. Example: `https://snipe-it.scswiderski.net` |
| `SnipeItTokenVar` | string |          | Name of Automation variable storing API token. Default `SnipeItApiToken`.    |
| `ContainerSasVar` | string |          | Name of Automation variable storing SAS. Default `SnipeItContainerSasUrl`.   |
| `BlobPrefix`      | string |          | Blob name prefix. Default `snipeit-backup`.                                  |

**Blob naming:**
`${BlobPrefix}-YYYYMMDD_HHMMSS-{filename}.zip`

---

## Quick start

1. **Add Automation Variables**

   * `SnipeItApiToken` (String, Encrypted=On)
   * `SnipeItContainerSasUrl` (String, Encrypted=On)

2. **Import runbook**

   * Type: PowerShell 7.2
   * Upload `SnipeIT-Backup-Pull.ps1`.

3. **Test run**

   * Parameters:

     * `SnipeItBaseUrl`: `https://snipe-it.your-snipe-domain.net`
   * Do **not** pass SAS or token manually; they come from variables.

4. **Schedule**

   * Run daily after the Snipe-IT host’s backup cron finishes.

---

## Networking

* Outbound HTTPS to:

  * `https://<snipeit>/api/v1/settings/backups`
  * `https://<snipeit>/api/v1/settings/backups/download/{file}`
  * `https://<storage>.blob.core.windows.net/<container>/<blob>?<sas>`


---

## Error handling

* `No backups found via API listing.` → verify backups exist in UI.
* `401/403 from Snipe-IT` → token invalid/insufficient scope.
* `Blob upload 403` → SAS missing `w`/`c`.
* `Timeouts` → increase `-TimeoutSec`, check file size/network.

---

## Operational checklist

* [ ] Confirm Snipe-IT backups exist in UI.
* [ ] SAS not expired.
* [ ] Container lifecycle rules enforce retention.
* [ ] Runbook succeeded and uploaded file of expected size.

---
