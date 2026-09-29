# Environmental Compliance Obligations Register – QTMP Ormeau Maintenance Facility

This project replaces the DTI-HSEQ-TP002.1 Environmental Compliance Obligations Register (an Excel workbook that forms part of the project EMP) with:

- **SharePoint Lists** as the data store
- a **Power Apps canvas app** as the front end
- **Power Automate** flows for notifications, recurring occurrences and the monthly TP002 export

It uses Microsoft 365 standard connectors only.

## Status

| Phase | Deliverable | State |
|---|---|---|
| 1 | Data audit and mapping – [`docs/data-audit.md`](docs/data-audit.md) | Approved 29/09/2026 (all decisions accepted) |
| 2 | List schema and provisioning – [`docs/list-schema.md`](docs/list-schema.md), [`provision/`](provision/), [`docs/manual-provisioning.md`](docs/manual-provisioning.md) | **For review** |
| 3 | Data transform and load (`import/`) | Not started |
| 4 | Power Apps canvas app (`app/`, `docs/app-build-guide.md`) | Not started |
| 5 | Power Automate flows (`flows/`, `export/`) | Not started |

## Repository layout

```
source/     Original TP002.1 workbook (input to the audit and transform)
docs/       Audit, schema, build guides
provision/  list-schema.json (source of truth), Provision-Lists.ps1, doc generator, tests
import/     Workbook layout, audit, transform and load scripts
app/        Power Apps YAML, one file per screen (Phase 4)
flows/      Step-by-step Power Automate flow definitions (Phase 5)
export/     Fallback TP002 workbook exporter (Phase 5)
```

## Order of operations

### 1. Audit (done)
`pip install -r requirements.txt`, then `python import/audit_workbook.py` regenerates `docs/data-audit-inventory.md`.

### 2. Create the SharePoint lists
Choose a site (see `docs/list-schema.md` §5), then either:

**A. Script** (PowerShell 7.2+ and PnP.PowerShell 2.x or later):
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
# One-off, by an Entra admin: register an app for interactive PnP sign-in
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP - QTMP Env Register" -Tenant <tenant>.onmicrosoft.com -SharePointDelegatePermissions "AllSites.Manage"
# Preview, then run (safe to re-run)
./provision/Provision-Lists.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -ClientId <app-id> -SetRegionalSettings -WhatIf
./provision/Provision-Lists.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -ClientId <app-id> -SetRegionalSettings
```
**B. Browser**: follow `docs/manual-provisioning.md`.

Then work through the verification checklist at the end of `docs/manual-provisioning.md`.

### 3–5. *(added as each phase is approved)*

## Changing the schema
1. Edit `provision/list-schema.json`.
2. `python provision/build_schema_docs.py` refreshes the generated doc sections.
3. `python -m pytest provision` runs the schema checks, and `pwsh provision/tests/Test-Provisioning.ps1` runs the script against a mock SharePoint.
4. Re-run `Provision-Lists.ps1`. It only applies the differences.
