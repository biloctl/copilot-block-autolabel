<#
.SYNOPSIS
    Applies a "NoCopilot" sensitivity label to files in a designated folder
    within the OneDrive of every member of a licensing security group.

.DESCRIPTION
    Designed to run as an Azure Automation PowerShell runbook authenticating as an
    Entra app registration (client secret). For each in-scope user it:
        1. Resolves the user's OneDrive (skips gracefully if not provisioned)
        2. Ensures the target folder exists (creates it if missing)
        3. Recursively lists files in that folder
        4. Checks each file's current label (extractSensitivityLabels - NOT metered)
        5. Assigns the label only if it isn't already correct (assignSensitivityLabel - METERED)
        6. Honors throttling (429/503) and a daily call cap

    The runbook ONLY applies the label. Enforcement (keeping Copilot from
    using the file) is done separately by a DLP-for-Copilot policy keyed on
    this same label. Treat "labeled" and "enforced" as two independent things.

.PREREQUISITES
    AUTH MODEL: an Entra app registration (confidential client) authenticating with a
    client secret. (Managed identities are NOT supported for metered Graph APIs, which
    is why this uses an app registration.) Store the secret as an ENCRYPTED Automation
    variable - never in code.

    Graph APPLICATION permissions granted to the app registration (admin consent required):
        - GroupMember.Read.All     (read the licensing groups' members)
        - Files.ReadWrite.All      (read files, assign/extract labels across OneDrives)
      (Sites.ReadWrite.All can substitute for Files.ReadWrite.All depending on your tenant.)

    METERED API BILLING:
        assignSensitivityLabel is a metered/protected Graph API. The app registration's
        app ID MUST be associated with an Azure subscription for billing (via
        'az graph-services account create') or every call returns HTTP 402 regardless
        of permissions. Enable this BEFORE running for real.
        (extractSensitivityLabels is NOT metered, so the label-check step is free.)

    LABEL ENCRYPTION CAVEAT:
        If the label applies encryption (IRM), app-only assignment may return
        "not supported" and require delegated user context. This runbook assumes a
        CLASSIFICATION-ONLY label, with the DLP-for-Copilot policy doing enforcement.
        Confirm before rollout.

    Modules to import into the Automation account:
        Microsoft.Graph.Authentication, Microsoft.Graph.Groups

.NOTES
    This is a sketch / starting point. Validate end-to-end on ONE OneDrive
    (a test group containing only your account + WhatIfMode) before widening scope.
#>

#region ---------- Configuration (prefer Automation Variables over hardcoding) ----------

# Reads an Automation variable, falling back to a default if it isn't set.
# Lets the runbook work out-of-the-box while keeping values changeable without a code edit.
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

# Licensing security group object IDs - comma-separated list in one variable.
# Test: point this at a test group containing only your account.
# Go live: change the variable to all 40 group IDs. No code change either way.
# No default - the runbook must not guess its scope.
$GroupIdsRaw  = Get-Config -Name 'Confidential_GroupIds'
$GroupIds     = @($GroupIdsRaw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# Immutable GUID of the sensitivity label this runbook applies.
# Set the 'Confidential_LabelId' variable. CONFIRM your DLP-for-Copilot policy keys on the SAME GUID.
$LabelId      = Get-Config -Name 'Confidential_LabelId'

# Folder (relative to OneDrive root) that triggers labeling. Set the 'Confidential_FolderName' variable.
$FolderName   = Get-Config -Name 'Confidential_FolderName'

$DailyCallCap = 90000                                               # buffer below the tenant-wide 100,000/day metered cap
$MaxRetries   = 5                                                   # per-call retry attempts on throttling/transient errors

# --- App registration (confidential client) auth ---
$TenantId     = Get-Config -Name 'Confidential_TenantId'        # directory (tenant) ID
$ClientId     = Get-Config -Name 'Confidential_ClientId'        # app registration application (client) ID
$ClientSecret = Get-Config -Name 'Confidential_ClientSecret'    # store as an ENCRYPTED Automation variable

# --- Test control ---
# Test by pointing Confidential_GroupIds at a test group containing only your account.
$WhatIfMode   = $true      # TRUE = log what WOULD be labeled, never calls the metered assign API

# --- Verification mode ---
# TRUE = audit only: read each file's ACTUAL label (free, non-metered) and report any
# file in scope that is MISSING the target label. Assigns nothing, spends nothing.
# Run this AFTER an apply run (give labels time to settle) as your coverage check -
# it catches silent "accepted but not applied" misses that the apply run can't detect.
$VerifyOnly   = $false

#endregion

#region ---------- Helpers ----------

$script:CallsMade = 0

# Logs to the INFORMATION stream, which (unlike Write-Output) is NOT captured when a
# function's return value is assigned to a variable. Use this for all informational
# logging inside functions whose output is captured.
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
# Invoke-MgGraphRequest throws HttpRequestException (.StatusCode), not the
# .Exception.Response.StatusCode shape, so check several locations + the message.
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
        'NotFound'                      { return 404 }
        'TooManyRequests'               { return 429 }
        'PaymentRequired'               { return 402 }
        'ServiceUnavailable|GatewayTimeout' { return 503 }
        default                         { return $null }
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
            # 429 = throttled, 503/504 = transient. Honor Retry-After if present.
            if ($status -in 429, 503, 504) {
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                $delay = if ($retryAfter) { [int]$retryAfter } else { [math]::Pow(2, $attempt) }
                Write-Warning "  Throttled/transient ($status). Attempt $attempt/$MaxRetries. Waiting $delay s."
                Start-Sleep -Seconds $delay
                continue
            }
            # 402 = metered billing not enabled. Fatal — stop the whole run, it won't fix itself.
            if ($status -eq 402) {
                throw "HTTP 402 from $Uri - metered API billing is not enabled for this app. Associate the app with an Azure subscription. Aborting."
            }
            throw   # anything else: bubble up to the per-file handler
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
        return $false   # nothing to label in a folder that doesn't exist yet
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
        [Parameter(Mandatory)][string] $Path   # drive-relative, e.g. "NoCopilot" or "NoCopilot/sub"
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

# Reads the sensitivity label IDs currently assigned to a file.
# extractSensitivityLabels is NOT metered. Returns @() if none, or $null if unreadable.
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

    # NOTE: assignSensitivityLabel is async - a 202 means accepted, not confirmed-applied.
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
    # Dedupe by user Id: the same person can sit in several of the groups.
    $seen = @{}

    foreach ($gid in $GroupIds) {
        try {
            # Cast to user: returns only users, as typed objects with Id/UserPrincipalName populated.
            # Transitive expands any nesting; for assigned groups it returns the direct users.
            $members = Get-MgGroupTransitiveMemberAsUser -GroupId $gid -All `
                        -Property 'id,userPrincipalName' -ErrorAction Stop
            $count = 0
            foreach ($u in $members) {
                # Read direct properties, falling back to AdditionalProperties across SDK shapes.
                $id  = if ($u.Id) { $u.Id } else { $u.AdditionalProperties['id'] }
                # UPN can land on the typed property OR in AdditionalProperties under varying casing.
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
            # One bad/inaccessible group shouldn't sink the whole run, but make it loud.
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
                        # Audit: report whether the target label is actually present. No assign.
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
