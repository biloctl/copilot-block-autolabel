# How-To: No-Copilot Auto-Labeling for OneDrive — Full Rebuild Guide

A complete, self-contained guide to rebuild this solution from scratch. Captures **what we
built, why we built it that way, the dead ends we ruled out, and the full script.**

---

## 1. What this solution does

Applies a **No-Copilot** sensitivity label to every file inside a
`CONFIDENTIAL` folder at the root of each user's OneDrive, across all members of one or more
licensing security groups. A labeled file, paired with the DLP policy below, is excluded from
Microsoft 365 Copilot.

- A scheduled **Azure Automation PowerShell runbook** does the labeling.
- It authenticates as an **Entra app registration** (client secret) and calls **Microsoft
  Graph** to read group membership, find each OneDrive, ensure the folder exists, and label
  the files.
- A **DLP-for-Copilot policy** keyed on this label performs the Copilot exclusion once the
  label is applied. You must create this policy — applying the label alone does not block
  Copilot. The runbook labels; the DLP policy enforces.

---

## 2. Why this design (and what we rejected)

We started by trying to do this with **Purview auto-labeling + a metadata column** and walked
it back. Recording the dead ends so we don't repeat them:

- **Metadata column + "Document property is" auto-labeling condition** — the original plan
  (custom column with a folder default, read by a policy). Rejected: OneDrive doesn't expose
  per-folder default column values or the search-schema UI; sensitivity auto-labeling policies
  **can't be scoped to a folder** (only to the whole OneDrive account); the crawled-property →
  managed-property mapping is fragile; and the condition is documented to match files but then
  fail to label them.
- **Dedicated SharePoint library with a default label** — rejected because OneDrive's
  automatic per-user access boundary would have to be recreated by breaking permission
  inheritance and managing unique permissions on thousands of folders. Worse problem than the
  one we're solving.
- **Power Automate** — rejected for scale. The OneDrive trigger/connector is per-user;
  covering all users centrally would mean building app-level Graph change notifications anyway,
  with weaker secret handling and limited run history.

**Why the runbook won:** we were going to script across every OneDrive no matter what (even
the column approach needed per-OneDrive provisioning). So instead of provisioning a fragile
metadata proxy for "this folder," the runbook just **evaluates the folder path directly in
code** — deterministic, auditable, no crawl/index dependency, no flaky condition.

**Why app registration, not managed identity:** we initially built this on a system-assigned
managed identity (no secrets to manage). But **managed identities are NOT supported for
metered Graph APIs**, and `assignSensitivityLabel` is metered. So we pivoted to an app
registration authenticating with a client secret.

**Why config lives in Automation Variables, not in the script:** the script reads every
environment-specific value (group IDs, label GUID, folder name, tenant/client/secret) from
Automation Variables via a `Get-Config` helper. This means scope or label changes are a
variable edit, not a code edit + republish; secrets never live in source; and the same script
works across environments. A missing required variable throws a clear error instead of running
with a wrong/empty value.

---

## 3. Reference values (ours — substitute your own)

| Item | Value |
|---|---|
| App registration (client) ID | `<YOUR_APP_CLIENT_ID>` |
| Subscription ID | `<YOUR_SUBSCRIPTION_ID>` |
| Resource group | `<YOUR_RESOURCE_GROUP>` |
| Automation account | `<YOUR_AUTOMATION_ACCOUNT>` |
| Runbook | `<YOUR_RUNBOOK_NAME>` |
| Metered billing resource | `GraphMeteredBilling` |
| Label immutable GUID | `<YOUR_LABEL_IMMUTABLE_GUID>` |
| SharePoint admin URL | `https://<YOUR_TENANT>-admin.sharepoint.com` |
| Folder name | `CONFIDENTIAL` (lookup is case-insensitive; casing is cosmetic) |
| Runtime | PowerShell 7.2 (custom Runtime Environment) |

---

## 4. Build steps

### Step 1 — App registration (Entra ID)
Create a confidential-client app registration; add a **client secret** (record the value and
its **expiry date** — see Open Items).

- Portal: **Entra ID → App registrations → New registration** → then **Certificates & secrets
  → New client secret**.
- CLI equivalent:
  ```bash
  az ad app create --display-name "Graph-AutoLabel-App"
  az ad sp create --id <appId>
  az ad app credential reset --id <appId> --append --years 1
  ```

### Step 2 — Graph application permissions (admin-consented)
Grant two **application** permissions and consent them:
- `GroupMember.Read.All` — read licensing group members
- `Files.ReadWrite.All` — read files + assign/extract labels across all OneDrives
  *(broad/tenant-wide — flag for security review; `Sites.Selected` is the least-privilege
  alternative but requires granting per-OneDrive, which doesn't scale)*

```powershell
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All","Application.Read.All"

# The app registration's SERVICE PRINCIPAL object ID (NOT the app/client ID)
$spId  = (Get-MgServicePrincipal -Filter "appId eq '<YOUR_APP_CLIENT_ID>'").Id

# Microsoft Graph's service principal in your tenant
$graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

foreach ($perm in 'GroupMember.Read.All','Files.ReadWrite.All') {
    $role = $graph.AppRoles | Where-Object { $_.Value -eq $perm -and $_.AllowedMemberTypes -contains 'Application' }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spId `
        -PrincipalId $spId -ResourceId $graph.Id -AppRoleId $role.Id
}
```
> Common mistake: using the **app/client ID** where the **service principal object ID** is
> required. The filter above resolves the right one. Verify on the app's **API permissions**
> blade — both show "Granted for <tenant>".

### Step 3 — Enable metered API billing
`assignSensitivityLabel` is **metered/protected** (~$0.00185/call, **100,000/day** tenant cap).
The app's client ID must be linked to a subscription or every call returns **HTTP 402**.
> `extractSensitivityLabels` (reading labels) is **FREE**. The Automation account already being
> in a subscription does **NOT** satisfy this — it's a separate app-to-subscription link.

```bash
az account set --subscription <YOUR_SUBSCRIPTION_ID>

az graph-services account create `
  --resource-group <YOUR_RESOURCE_GROUP> `
  --resource-name GraphMeteredBilling `
  --subscription <YOUR_SUBSCRIPTION_ID> `
  --app-id <YOUR_APP_CLIENT_ID> `
  --location global

# Verify (empty array = you're querying the WRONG active subscription)
az resource list --resource-type Microsoft.GraphServices/accounts
```
Success JSON shows `provisioningState: Succeeded`, matching `appId`, and a `billingPlanId`.
Permissions to run: **Contributor** on the subscription + **Application Owner/Administrator**
on the app. `--location global` is required (metered Graph is commercial-cloud only).

### Step 4 — Automation runtime + Graph modules
The **system-generated** PS 7.2 Runtime Environment is **locked** (can't add modules). Create a
**custom** one and attach the Graph modules.

- Portal: **Automation account → Runtime Environments → Create** (PowerShell **7.2**) →
  **Add from gallery**: `Microsoft.Graph.Authentication` first, then `Microsoft.Graph.Groups`
  (Groups depends on Authentication). Imports take several minutes — wait for **Available**.
- Then point the runbook at this custom environment.
- Sanity check inside the runbook: `Get-Module -ListAvailable Microsoft.Graph.Authentication`

### Step 5 — Create the Automation Variables (the externalized config)
Portal: **Automation account → Shared Resources → Variables → Add** (String type).

| Variable | Value | Notes |
|---|---|---|
| `Confidential_GroupIds` | comma-separated group object IDs | test group(s) first; all groups later |
| `Confidential_LabelId` | `<YOUR_LABEL_IMMUTABLE_GUID>` | the label's **immutable** GUID |
| `Confidential_FolderName` | `CONFIDENTIAL` | |
| `Confidential_TenantId` | directory (tenant) ID | |
| `Confidential_ClientId` | `<YOUR_APP_CLIENT_ID>` | |
| `Confidential_ClientSecret` | the secret value | **ENCRYPTED** |

Get the label's **immutable** GUID (Security & Compliance PowerShell — NOT the portal URL GUID):
```powershell
Connect-IPPSSession
Get-Label -Identity "<YOUR_LABEL_NAME>" | Select-Object DisplayName, ImmutableId
```
CLI to create a variable:
```powershell
New-AzAutomationVariable -AutomationAccountName "<YOUR_AUTOMATION_ACCOUNT>" `
  -ResourceGroupName "<YOUR_RESOURCE_GROUP>" `
  -Name "Confidential_ClientSecret" -Value "<secret>" -Encrypted $true
```

### Step 6 — Create the runbook
Portal: **Automation account → Runbooks → Create** → PowerShell, runtime **7.2**, assigned to
your custom Runtime Environment. Paste the full script (Section 6 below). **Save**, then
**Publish**.

### Step 7 — Test in stages (do NOT skip)
1. Point `Confidential_GroupIds` at a **test group containing only your account** (or a small
   group of ~10 testers).
2. Leave `$WhatIfMode = $true`. Run from the **Test pane**. Expect it to connect, expand the
   group, resolve OneDrive(s), check/"would create" the folder, and log "would label" per file
   — **0 metered calls**.
3. Put a couple of Office files in your `CONFIDENTIAL` folder; re-run WhatIf until **Files
   seen** is non-zero and you see "would label" lines.
4. Flip `$WhatIfMode = $false`, run live. Expect **Metered calls** and **Newly labeled** to
   match the file count, 0 errors. Re-run once more — should show **Already labeled** = total,
   **Metered calls = 0** (proves the check-first / no-double-spend behavior).
5. Confirm Copilot is actually blocked on a labeled file (ask Copilot to summarize it from the
   user's account; it should refuse/fail). Allow propagation time.

### Step 8 — Schedule
Daily at **3 AM EST**, recurring, no expiration. Overnight chosen because **open/locked files
fail to label silently** — 3 AM minimizes that.
- Portal: **Runbook → Schedules → Link to schedule**. The schedule runs the **published**
  runbook, which **must be in APPLY mode** (`$WhatIfMode=$false`, `$VerifyOnly=$false`).
- CLI:
  ```powershell
  New-AzAutomationSchedule -AutomationAccountName "<YOUR_AUTOMATION_ACCOUNT>" `
    -ResourceGroupName "<YOUR_RESOURCE_GROUP>" -Name "Daily" `
    -StartTime (Get-Date "03:00:00").AddDays(1) -DayInterval 1 -TimeZone "America/New_York"
  Register-AzAutomationScheduledRunbook -AutomationAccountName "<YOUR_AUTOMATION_ACCOUNT>" `
    -ResourceGroupName "<YOUR_RESOURCE_GROUP>" -RunbookName "<YOUR_RUNBOOK_NAME>" -ScheduleName "Daily"
  ```
- Go live: change `Confidential_GroupIds` to all group IDs. The first full run is the **backfill** —
  watch metered calls vs the 100k/day cap (overflow defers automatically).

---

## 5. Operating it

- **Apply** runs nightly on the schedule. New files get labeled within ~24h.
- **Verify** (coverage check): set `$VerifyOnly = $true` and run from the Test pane whenever you
  want. Read-only, free; lists every in-scope file **MISSING** the label by user + filename.
  Set it back to `$false` afterward (the published/scheduled copy must stay in apply mode).
- Only act on files that stay MISSING across multiple checks. Transient misses are usually just
  open/locked files that self-resolve on the next run.

---

## 6. The full runbook script

> `Apply-SensitivityLabel.ps1`. Config is read from Automation Variables (Step 5). Mode is
> controlled by `$WhatIfMode` / `$VerifyOnly` near the top.

```powershell
<#
.SYNOPSIS
    Applies a sensitivity label to files in a designated folder within the OneDrive of
    every member of one or more licensing security groups.

.DESCRIPTION
    Azure Automation PowerShell runbook authenticating as an Entra app registration
    (client secret). For each in-scope user it:
        1. Resolves the user's OneDrive (skips gracefully if not provisioned)
        2. Ensures the target folder exists (creates it if missing)
        3. Recursively lists files in that folder
        4. Checks each file's current label (extractSensitivityLabels - NOT metered)
        5. Assigns the label only if not already correct (assignSensitivityLabel - METERED)
        6. Honors throttling (429/503) and a daily call cap

.PREREQUISITES
    - App registration (confidential client) with client secret stored as an ENCRYPTED
      Automation variable. (Managed identity is NOT supported for metered Graph APIs.)
    - Graph application permissions: GroupMember.Read.All, Files.ReadWrite.All (admin consent).
    - Metered billing: app ID linked to a subscription via 'az graph-services account create'
      or assignSensitivityLabel returns HTTP 402. (extractSensitivityLabels is free.)
    - Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Groups (PS 7.2 runtime env).
    - LABEL CAVEAT: assumes a classification label applicable app-only. An encryption-bearing
      label may not be assignable app-only.
#>

#region ---------- Configuration (read from Automation Variables) ----------

# Reads an Automation variable, falling back to a default if it isn't set.
function Get-Config {
    param(
        [Parameter(Mandatory)][string] $Name,
        [string] $Default
    )
    try {
        $val = Get-AutomationVariable -Name $Name -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
        return $val
    }
    catch {
        if ($PSBoundParameters.ContainsKey('Default')) { return $Default }
        throw "Required Automation variable '$Name' is not set and has no default."
    }
}

# Licensing security group object IDs - comma/semicolon-separated list in one variable.
$GroupIdsRaw  = Get-Config -Name 'Confidential_GroupIds'
$GroupIds     = @($GroupIdsRaw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# Immutable GUID of the sensitivity label this runbook applies.
$LabelId      = Get-Config -Name 'Confidential_LabelId'

# Folder (relative to OneDrive root) that triggers labeling.
$FolderName   = Get-Config -Name 'Confidential_FolderName'

$DailyCallCap = 90000      # buffer below the tenant-wide 100,000/day metered cap
$MaxRetries   = 5          # per-call retry attempts on throttling/transient errors

# --- App registration (confidential client) auth ---
$TenantId     = Get-Config -Name 'Confidential_TenantId'        # directory (tenant) ID
$ClientId     = Get-Config -Name 'Confidential_ClientId'        # app registration application (client) ID
$ClientSecret = Get-Config -Name 'Confidential_ClientSecret'    # store as an ENCRYPTED Automation variable

# --- Mode controls ---
$WhatIfMode   = $true      # TRUE = log what WOULD be labeled, never calls the metered assign API
$VerifyOnly   = $false     # TRUE = audit only: report files MISSING the label (free, no assign)

#endregion

#region ---------- Helpers ----------

$script:CallsMade = 0

# Logs to the INFORMATION stream, which (unlike Write-Output) is NOT captured when a
# function's return value is assigned to a variable.
function Write-Log {
    param([string] $Message)
    Write-Information $Message -InformationAction Continue
}

function Connect-Graph {
    Write-Output "Connecting to Microsoft Graph as app registration $ClientId ..."
    $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $cred   = New-Object System.Management.Automation.PSCredential ($ClientId, $secure)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
    $ctx = Get-MgContext
    Write-Output "Connected. AppId=$($ctx.ClientId) Scopes=$($ctx.Scopes -join ',')"
}

# Robustly extracts an HTTP status code from a Graph SDK error.
# Invoke-MgGraphRequest throws HttpRequestException (.StatusCode), not .Exception.Response.StatusCode.
function Get-GraphErrorStatus {
    param($ErrorRecord)
    $ex = $ErrorRecord.Exception
    if ($ex.PSObject.Properties.Name -contains 'StatusCode' -and $ex.StatusCode) {
        return [int]$ex.StatusCode
    }
    if ($ex.Response -and $ex.Response.StatusCode) {
        return [int]$ex.Response.StatusCode
    }
    switch -Regex ($ex.Message) {
        'NotFound'                          { return 404 }
        'TooManyRequests'                   { return 429 }
        'PaymentRequired'                   { return 402 }
        'ServiceUnavailable|GatewayTimeout' { return 503 }
        default                             { return $null }
    }
}

# Wrapper around Invoke-MgGraphRequest with throttle-aware retry/backoff.
function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        [object] $Body
    )
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject' }
            if ($null -ne $Body) {
                $params.Body        = ($Body | ConvertTo-Json -Depth 6)
                $params.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @params
        }
        catch {
            $status = Get-GraphErrorStatus $_
            if ($status -in 429, 503, 504) {
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                $delay = if ($retryAfter) { [int]$retryAfter } else { [math]::Pow(2, $attempt) }
                Write-Warning "  Throttled/transient ($status). Attempt $attempt/$MaxRetries. Waiting $delay s."
                Start-Sleep -Seconds $delay
                continue
            }
            if ($status -eq 402) {
                throw "HTTP 402 from $Uri - metered API billing is not enabled for this app. Associate the app with an Azure subscription. Aborting."
            }
            throw
        }
    }
    throw "Exceeded $MaxRetries retries for $Method $Uri"
}

# Returns the user's default drive id, or $null if OneDrive isn't provisioned yet.
function Get-UserDriveId {
    param([Parameter(Mandatory)][string] $UserId)
    try {
        $drive = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/drive"
        return $drive.id
    }
    catch {
        $status = Get-GraphErrorStatus $_
        if ($status -eq 404) {
            Write-Warning "  No OneDrive provisioned for $UserId - skipping (will be picked up on a later run)."
            return $null
        }
        throw
    }
}

# Ensures the target folder exists in the drive root; creates it if missing.
function Confirm-Folder {
    param(
        [Parameter(Mandatory)][string] $DriveId,
        [Parameter(Mandatory)][string] $Folder
    )
    try {
        $null = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$Folder"
        return $true
    }
    catch {
        if ((Get-GraphErrorStatus $_) -ne 404) { throw }
    }
    if ($WhatIfMode) {
        Write-Log "  [WhatIf] Would create folder '$Folder'."
        return $false
    }
    Write-Log "  Creating folder '$Folder'."
    $body = @{ name = $Folder; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' }
    $null = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children" -Body $body
    return $true
}

# Recursively returns all FILE driveItems under the given folder path.
function Get-FolderFiles {
    param(
        [Parameter(Mandatory)][string] $DriveId,
        [Parameter(Mandatory)][string] $Path
    )
    $files = @()
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($Path):/children?`$top=200"
    do {
        $page = Invoke-GraphWithRetry -Method GET -Uri $uri
        foreach ($item in $page.value) {
            if ($item.folder) {
                $files += Get-FolderFiles -DriveId $DriveId -Path "$Path/$($item.name)"
            }
            elseif ($item.file) {
                $files += $item
            }
        }
        $uri = $page.'@odata.nextLink'
    } while ($uri)
    return $files
}

# Reads the sensitivity label IDs currently assigned to a file (NOT metered).
function Get-FileLabelIds {
    param(
        [Parameter(Mandatory)][string] $DriveId,
        [Parameter(Mandatory)][string] $ItemId
    )
    try {
        $result = Invoke-GraphWithRetry -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/extractSensitivityLabels"
        return @($result.labels | Select-Object -ExpandProperty sensitivityLabelId -ErrorAction SilentlyContinue)
    }
    catch {
        Write-Warning "    Could not read label on $ItemId : $($_.Exception.Message)"
        return $null
    }
}

# True if the file already carries the target label.
function Test-AlreadyLabeled {
    param(
        [Parameter(Mandatory)][string] $DriveId,
        [Parameter(Mandatory)][string] $ItemId
    )
    $ids = Get-FileLabelIds -DriveId $DriveId -ItemId $ItemId
    if ($null -eq $ids) { return $false }   # unreadable -> let the assign step try
    return ($ids -contains $LabelId)
}

# Assigns the target label (METERED). Honors WhatIf and the daily cap.
function Set-NoCopilotLabel {
    param(
        [Parameter(Mandatory)][string] $DriveId,
        [Parameter(Mandatory)][string] $ItemId,
        [Parameter(Mandatory)][string] $ItemName
    )
    if ($WhatIfMode) {
        Write-Log "    [WhatIf] Would label: $ItemName"
        return
    }
    if ($script:CallsMade -ge $DailyCallCap) {
        Write-Warning "    Daily call cap ($DailyCallCap) reached - deferring '$ItemName' to next run."
        return
    }
    # assignSensitivityLabel is async - a 202 means accepted, not confirmed-applied.
    $body = @{ sensitivityLabelId = $LabelId; assignmentMethod = 'standard' }
    $null = Invoke-GraphWithRetry -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/assignSensitivityLabel" `
        -Body $body
    $script:CallsMade++
    Write-Log "    Labeled (accepted): $ItemName"
}

# Resolves the distinct set of users to process across one or more groups (deduped).
function Get-ScopedUsers {
    if (-not $GroupIds -or $GroupIds.Count -eq 0) {
        throw "No groups in scope. Set the 'Confidential_GroupIds' variable (comma-separated object IDs)."
    }
    Write-Log "Resolving members across $($GroupIds.Count) group(s)..."
    $seen = @{}
    foreach ($gid in $GroupIds) {
        try {
            # Cast to user: returns only users, as typed objects with Id/UserPrincipalName populated.
            $members = Get-MgGroupTransitiveMemberAsUser -GroupId $gid -All `
                        -Property 'id,userPrincipalName' -ErrorAction Stop
            $count = 0
            foreach ($u in $members) {
                $id  = if ($u.Id) { $u.Id } else { $u.AdditionalProperties['id'] }
                $upn = $u.UserPrincipalName
                if (-not $upn -and $u.AdditionalProperties) {
                    $upn = $u.AdditionalProperties['userPrincipalName']
                    if (-not $upn) { $upn = $u.AdditionalProperties['userprincipalname'] }
                }
                if (-not $upn) { $upn = $id }   # never log a blank line
                if ($id -and -not $seen.ContainsKey($id)) {
                    $seen[$id] = [pscustomobject]@{ Id = $id; UserPrincipalName = $upn }
                    $count++
                }
            }
            Write-Log "  Group $gid : $count user(s)."
        }
        catch {
            Write-Warning "  Group $gid : FAILED to read members - $($_.Exception.Message)"
        }
    }
    $users = @($seen.Values)
    Write-Log "Total distinct user(s) in scope: $($users.Count)."
    return $users
}

#endregion

#region ---------- Main ----------

$summary = [pscustomobject]@{
    UsersProcessed = 0
    UsersSkipped   = 0
    FilesSeen      = 0
    FilesLabeled   = 0
    FilesAlready   = 0
    FilesVerified  = 0
    FilesMissing   = 0
    MissingList    = @()
    Errors         = 0
}

try {
    Connect-Graph
    $users = Get-ScopedUsers

    foreach ($user in $users) {
        $upn = $user.UserPrincipalName
        Write-Output "----- $upn -----"
        try {
            $driveId = Get-UserDriveId -UserId $user.Id
            if (-not $driveId) { $summary.UsersSkipped++; continue }

            $folderExists = Confirm-Folder -DriveId $driveId -Folder $FolderName
            if (-not $folderExists) { $summary.UsersProcessed++; continue }

            $files = Get-FolderFiles -DriveId $driveId -Path $FolderName
            Write-Output "  $($files.Count) file(s) in '$FolderName'."

            foreach ($file in $files) {
                $summary.FilesSeen++
                try {
                    if ($VerifyOnly) {
                        $ids = Get-FileLabelIds -DriveId $driveId -ItemId $file.id
                        if ($null -ne $ids -and ($ids -contains $LabelId)) {
                            $summary.FilesVerified++
                        }
                        else {
                            $summary.FilesMissing++
                            $summary.MissingList += "$upn :: $($file.name)"
                            Write-Warning "    MISSING label: $($file.name)"
                        }
                        continue
                    }
                    if (Test-AlreadyLabeled -DriveId $driveId -ItemId $file.id) {
                        $summary.FilesAlready++
                        continue
                    }
                    Set-NoCopilotLabel -DriveId $driveId -ItemId $file.id -ItemName $file.name
                    if (-not $WhatIfMode) { $summary.FilesLabeled++ }
                }
                catch {
                    $summary.Errors++
                    Write-Warning "    ERROR on file '$($file.name)': $($_.Exception.Message)"
                }
            }
            $summary.UsersProcessed++
        }
        catch {
            $summary.Errors++
            Write-Warning "  ERROR processing $upn : $($_.Exception.Message)"
        }
    }
}
finally {
    $mode = if ($VerifyOnly) { 'VERIFY-ONLY' } elseif ($WhatIfMode) { 'WHATIF' } else { 'APPLY' }
    Write-Output "================ RUN SUMMARY ================"
    Write-Output ("Mode            : {0}" -f $mode)
    Write-Output ("Users processed : {0}" -f $summary.UsersProcessed)
    Write-Output ("Users skipped   : {0}" -f $summary.UsersSkipped)
    Write-Output ("Files seen      : {0}" -f $summary.FilesSeen)
    if ($VerifyOnly) {
        Write-Output ("Verified labeled: {0}" -f $summary.FilesVerified)
        Write-Output ("MISSING label   : {0}" -f $summary.FilesMissing)
    }
    else {
        Write-Output ("Already labeled : {0}" -f $summary.FilesAlready)
        Write-Output ("Newly labeled   : {0}" -f $summary.FilesLabeled)
        Write-Output ("Metered calls   : {0}" -f $script:CallsMade)
    }
    Write-Output ("Errors          : {0}" -f $summary.Errors)
    if ($summary.MissingList.Count -gt 0) {
        Write-Output "--- Files MISSING the label ---"
        foreach ($m in $summary.MissingList) { Write-Output "  $m" }
    }
    Write-Output "============================================="
}

#endregion
```

---

## 7. Key lessons / gotchas (these cost us time)

- **Managed identity won't work** for the metered call — use an app registration. This is why
  the whole auth model is app-registration + client secret.
- **"Accepted" ≠ "applied."** `assignSensitivityLabel` returns 202 and labels asynchronously; a
  file can be accepted but never labeled, **silently**. The 202's Location header returns the
  true operation state, and failures are logged in the Purview audit log. **Verify mode is the
  practical safety net.**
- **Open/locked files fail silently.** A file open for editing can't be written to, so the label
  never lands and no clean error appears. This was the root cause of a "mystery" unlabeled docx
  (it was open in Word). Drove the **3 AM schedule** and the **"make changes in browser, not the
  synced client"** guidance for users.
- **PDFs** — gated by a tenant setting (`EnableSensitivityLabelForPDF`, currently OFF, pending a
  change request) and inconsistent even when on. Office formats (.docx/.pptx/.xlsx) work cleanly.
- **Unsupported file types** (.jpg, .csv, .txt, .cer, etc.) can never be labeled and will always
  show MISSING in verify — platform limitation, not a bug.
- **Label overwrites** — the API overwrites any existing label; the check-first step avoids
  redundant metered calls and protects already-correctly-labeled files.
- **SDK shape gotchas (specific bugs we hit and fixed):**
  - `Get-MgGroupTransitiveMember` returned empty IDs/UPNs → switched to
    `Get-MgGroupTransitiveMemberAsUser` with a property fallback.
  - `Invoke-MgGraphRequest` errors don't expose `.Exception.Response.StatusCode.value__` →
    `Get-GraphErrorStatus` reads `HttpRequestException.StatusCode` with a message fallback.
  - `Write-Output` for logging **inside a function whose return is captured** pollutes the
    return value (log strings get iterated as if they were users) → all in-function logging
    goes through `Write-Log` (Information stream).
- **Runtime Environments** — the system-generated PS7.2 environment is locked; you must create a
  custom one to add the Graph modules, and point the runbook at it. Modules are per-runtime, so a
  module imported under the wrong runtime version yields "cmdlet not recognized."
- **PDF tenant check** (SharePoint Online Management Shell):
  ```powershell
  Connect-SPOService -Url https://<YOUR_TENANT>-admin.sharepoint.com
  Get-SPOTenant | Select-Object EnableAIPIntegration, EnableSensitivityLabelForPDF
  # to enable (change request): Set-SPOTenant -EnableSensitivityLabelForPDF $true
  ```

---

## 8. Folder behavior (for user communications)

- Must be named **`CONFIDENTIAL`** at the **OneDrive root** (lookup is case-insensitive; casing
  is cosmetic).
- **Renamed** → a new `CONFIDENTIAL` folder is created on the next run.
- **Existing** folder → used as-is, its contents labeled (no duplicate created).
- **Nested folders inside** it → labeled **recursively**.
- **Moved inside another folder** → no longer found at root; a new root-level one is created.

---

## 9. Open items / ongoing maintenance

- [ ] **PDF tenant setting** — pending change-control approval (`Set-SPOTenant -EnableSensitivityLabelForPDF $true`).
- [ ] **Client secret expiry** — set a calendar reminder to rotate before it lapses. This is the
      one silent failure mode that kills the runbook months out with no obvious cause.
- [ ] Decide whether to **lock down label removal** in the label publish policy (prevents a user
      stripping the label to re-expose a file to Copilot).
- [ ] **Periodic verify-mode check** — act only on files that persist as MISSING across runs.
- [ ] First **all-groups** run is the backfill — watch metered calls vs the 100k/day cap.
- [ ] If you ever move to an **encryption-bearing label**, retest app-only assignment — it may
      not be assignable app-only and could break the unattended model.
