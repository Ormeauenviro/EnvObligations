# Environmental Compliance Obligations Register – QTMP Ormeau Maintenance Facility

This project replaces the DTI-HSEQ-TP002.1 Environmental Compliance Obligations Register (an Excel workbook that forms part of the project EMP) with:

- **SharePoint Lists** as the data store
- a **Power Apps canvas app** as the front end
- **Power Automate** flows for notifications, recurring occurrences and the monthly TP002 export

It uses Microsoft 365 standard connectors only.

## Status

| Phase | Deliverable | State |
|---|---|---|
| 1 | Data audit and mapping – [`docs/data-audit.md`](docs/data-audit.md) | **For review** |
| 2 | List schema and provisioning (`docs/list-schema.md`, `provision/`) | Not started |
| 3 | Data transform and load (`import/`) | Not started |
| 4 | Power Apps canvas app (`app/`, `docs/app-build-guide.md`) | Not started |
| 5 | Power Automate flows (`flows/`, `export/`) | Not started |

## Repository layout

```
source/     Original TP002.1 workbook (input to the audit and transform)
docs/       Audit, schema, build guides
provision/  PnP.PowerShell list provisioning (Phase 2)
import/     Workbook layout, audit, transform and load scripts
app/        Power Apps YAML, one file per screen (Phase 4)
flows/      Step-by-step Power Automate flow definitions (Phase 5)
export/     Fallback TP002 workbook exporter (Phase 5)
```

## Order of operations

1. `pip install -r requirements.txt`, then `python import/audit_workbook.py` to regenerate `docs/data-audit-inventory.md`.
2. *(Phase 2 onwards – added as each phase is approved.)*
