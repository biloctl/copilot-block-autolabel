# OneDrive Sensitivity Auto-Labeling Runbook

An Azure Automation PowerShell runbook that applies a Microsoft Purview **sensitivity label**
to every file inside a designated folder (e.g. `Confidential`) at the root of each user's
OneDrive, across the members of one or more security groups. Applying the label is only half of
the solution: to actually exclude those files from Microsoft 365 Copilot you must also create a
**DLP-for-Copilot policy** keyed on the same label. The runbook labels; the DLP policy enforces.
Treat "labeled" and "enforced" as two independent things.

It authenticates as an Entra **app registration** (client secret) and calls **Microsoft Graph**
to resolve group membership, find each OneDrive, ensure the folder exists, and label the files.
It is **idempotent** (skips already-labeled files via a free read), **throttle-aware**, and has
a **verify mode** that audits which in-scope files are missing the label.

## Why a runbook instead of Purview auto-labeling?
Sensitivity auto-labeling policies can't be scoped to a specific folder (only to the whole
OneDrive account), and OneDrive doesn't expose the metadata/column machinery you'd need to fake
it. A runbook evaluates the folder path directly — deterministic, auditable, no crawl/index
dependency. See `HOWTO-public.md` for the full rationale and the approaches that were rejected.

## Files
- `Apply-SensitivityLabel.ps1` — the runbook.
- `HOWTO-public.md` — full rebuild guide: prerequisites, Azure/Graph setup, commands, testing,
  scheduling, and a hard-won "gotchas" section.

## Quick start
1. Create an Entra **app registration** with a client secret.
2. Grant it Graph **application** permissions `GroupMember.Read.All` and `Files.ReadWrite.All`
   (admin consent).
3. Enable **metered API billing** for the app (`assignSensitivityLabel` is metered) by linking
   its app ID to an Azure subscription via `az graph-services account create`.
4. In an Azure Automation account (PowerShell 7.2 custom runtime with
   `Microsoft.Graph.Authentication` and `Microsoft.Graph.Groups`), create these **Automation
   Variables**:

   | Variable | Description |
   |---|---|
   | `Confidential_GroupIds` | comma-separated security-group object IDs |
   | `Confidential_LabelId` | the sensitivity label's **immutable** GUID |
   | `Confidential_FolderName` | folder name at the OneDrive root (e.g. `Confidential`) |
   | `Confidential_TenantId` | directory (tenant) ID |
   | `Confidential_ClientId` | app registration application (client) ID |
   | `Confidential_ClientSecret` | client secret (**store encrypted**) |

5. Import the runbook, test with `$WhatIfMode = $true` against a small group, then schedule it.
6. Create a **DLP-for-Copilot policy** keyed on the same label GUID so labeled files are actually
   excluded from Copilot. Labeling alone does not enforce anything.

> **Test first.** Point `Confidential_GroupIds` at a group containing only your own account and
> run with `$WhatIfMode = $true` (logs only, no API spend) before widening scope.

## Modes
- `$WhatIfMode = $true` — log what *would* be labeled; makes no metered calls.
- `$VerifyOnly = $true` — read-only audit; reports files **missing** the label (free).
- both `$false` — apply mode.

## Notes & limitations
- **Managed identity is not supported** for metered Graph APIs — hence the app-registration +
  client-secret model.
- `assignSensitivityLabel` is **async**: a `202` means *accepted*, not *applied*. Use verify
  mode to confirm coverage.
- **Open/locked files fail silently** — schedule runs off-hours.
- Only **Office formats** label cleanly out of the box; PDF requires a tenant setting and is
  inconsistent; other types (images, CSV, TXT, etc.) can't be labeled at all.
- The client secret **expires** — rotate it before it lapses or the runbook stops silently.

## Security
No secrets or tenant identifiers are committed in this repo. All environment-specific values are
supplied at runtime via Automation Variables. Do not commit secrets — see `.gitignore`.

## License
   MIT
