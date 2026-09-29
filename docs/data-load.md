# Phase 3 – Data transform and load

This phase turns the TP002.1 workbook into clean data for each list, then loads it into the SharePoint lists built in Phase 2. It runs in two steps:

```
source/*.xlsx ──► import/transform.py ──► import/output/*.json, *.csv, import-report.md ──► import/Load-Data.ps1 ──► SharePoint lists
                      ▲                                                                          ▲
     import/config/role-assignments.csv                                             provision/list-schema.json
     import/config/project-milestones.csv
```

## 1. What gets loaded

| List | Items | Notes |
|---|---|---|
| Source Documents | 27 | 13 from _Doc Ref (4 marked *Provided = No*), plus 14 permits and plans the register cites (D-14). The permits carry their decision-notice and folder links. |
| Project Milestones | 6 | Names only, with **dates blank** until you supply them (D-10) |
| Role Assignments | 10 | Environmental and Sustainability Team → Daniel Blunt (+ Ethan Little copied) · Environmental Manager → Sam Moraes · Project Manager → Daniel Blunt · Environmental Advisor → Richard Kightley · the other 6 roles unassigned |
| Obligations | 225 | 208 obligations + 8 sub-conditions + 7 info-only + 2 needing review. No Name row 7 became Tasks (D-06), and CEMP row 11 was split into CEMP-006 (weekly) and CEMP-006B (monthly audit) (D-05). |
| Occurrences | 21 | 3 notification splits (commencement ✓ / completion pending): BC-DA-005, MP-TNAR-002, MP-NAC-003. Plus 5 Bridge Creek monitoring obligations × Year 1/2/5, undated until milestones are known. Recurring periods come from the Phase 5 occurrence generator. |
| Obligation Tasks | 4 | NN-DA-001 works (a)–(d). (a) is ticked. |
| Evidence | 71 | Every hyperlink in the DA sheets, with relative links resolved to full SharePoint URLs |
| Follow-ups | 5 | Import review items (D-12): RPP-001, LEG-016, ADR-001, CEMP-036, CEMP-023 |
| Status History | 246 | One *Import* entry per obligation and occurrence recording its starting status |
| Report Components | 26 | Reporting Matrix ✕ marks, each linked to the obligation that produces the report |

**Starting statuses:** 183 Not started, 31 Underway, 4 Complete (each has evidence), and 7 N/A – Info only.

Everything the transform can't decide on its own is listed in **`import/output/import-report.md`**:

- 25 obligations with no owner (5 unassigned roles)
- 35 items waiting on milestone dates
- 30 evidence-gap cells
- 30 Excel review comments to triage

## 2. Re-running safely

Both scripts are safe to re-run.

- **Transform:** it's deterministic, so the same workbook and config always give the same output. After editing either config CSV, run `python import/transform.py` again.
- **Loader:** it upserts on each list's key and never creates duplicates or deletes anything.

| List | Key | Fields the loader keeps in line with the register | Fields the team owns: set on create, filled if blank, **never overwritten** |
|---|---|---|---|
| Obligations | Obligation ID | Requirement, references, action, evidence required, timing detail, due rule, Extra Fields and so on | Status, Status Reason, Outcome, Compliance Notes, Owner, Responsible Role, Frequency, Stage, Risk, dates, Active, flags |
| Occurrences | Occurrence Key | Obligation, summary, period label | Status, dates, assignee, evidence URL, notes |
| Obligation Tasks | Obligation ID + Task Order | Task text, obligation | Done, assignee, dates, notes |
| Evidence | Obligation ID + Occurrence Key + URL | Title | Date, Added By, notes |
| Follow-ups | Obligation ID + Title | Reason | Owner, date, status, resolution |
| Source Documents | Document Key | Title, category, reference | Version, date, links, provided, comments |
| Role Assignments / Project Milestones | Title | – | People / dates / notes |
| Status History | Item + type + source | *(create only, never updated)* | |
| Report Components | Content item + report | All | |

`-OverwriteLiveFields` also resets the team-owned fields to the register values. **Use it only before go-live.**

So if you name a Construction Manager in `import/config/role-assignments.csv`, re-running the transform and loader fills the Owner on those 25 obligations. It leaves every other owner alone. The same applies to milestone dates in `import/config/project-milestones.csv` (format DD/MM/YYYY).

## 3. Steps

### 3.1 Transform (any machine with Python 3.10+)
```bash
pip install -r requirements.txt
python import/transform.py            # writes import/output/
python -m pytest import               # 24 tests, including "nothing lost" and Marine Plants
```
Open `import/output/import-report.md` and the CSVs in Excel to review the output before loading.

### 3.2 Load (PowerShell 7.2+ with PnP.PowerShell, the same app registration as Phase 2)
```powershell
# 1. Dry run: prints counts per list and writes import/output/load-plan.csv (every planned create/update)
./import/Load-Data.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -ClientId <app-id> -DryRun
# 2. Load
./import/Load-Data.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -ClientId <app-id>
# 3. Run it again: every list should report Created 0, Updated 0
./import/Load-Data.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/<site> -ClientId <app-id>
```
Each real run writes `import/output/load-log.csv`, and the script exits with an error if any item failed. `-Lists Obligations,Evidence` loads a subset; lookups to lists not being loaded still resolve.

**People:** the loader adds each person to the site the first time it sees them. If an email can't be resolved, it warns and leaves that column blank, and the rest of the load continues. Emails use `FirstnameLastname@dtinfrastructure.com.au`. Correct them in `import/config/role-assignments.csv` if any differ.

**Dates:** date-only values are sent at 12:00 so the day can't shift, and SharePoint stores the date in the site's Brisbane time zone. Run Phase 2 with `-SetRegionalSettings`, or set the regional settings manually, **before** loading.

## 4. What to check in the browser

1. **Counts:** each list's item count (*Site contents*) matches the table in section 1.
2. **Obligations › By Register view:** 11 groups (10 with items). Open **MP-TNAR-001** and check:
   - Status Underway, notes "Compare to GIS as evidence"
   - Source Document *SARA Decision Notice 2308-36075 SDA…*
   - Extra Fields show the TP002 original values and `Unmapped (J): Ongoing`
3. **Occurrences:** BC-DA-005|COMMENCE is Complete with an evidence link, and BC-DA-005|COMPLETE is Not started.
4. **Obligation Tasks:** NN-DA-001 has four tasks, with (a) ticked.
5. **Evidence:** open two or three links and confirm they reach the right SharePoint folder. They are TP002 sharing links, so they open for anyone who already had access.
6. **Dates:** NN-DA-010 shows Due Date **31/07/2026**.
7. **People:** Owner shows names, not emails, on CEMP-001 (Daniel Blunt) and PFAS-001 (Sam Moraes).
8. **Follow-ups › Open view:** 5 items, each with *Has Open Follow-up* = Yes on its obligation.
9. **Version history** on any obligation shows version 1.0 created by the account that ran the load.

## 5. Tests

| Test | What it proves |
|---|---|
| `python -m pytest import` | Counts per list and per register; the Marine Plants blocks, IDs, realignment of row 23, inheritance, info-only parent, notification splits, DAF quote and permit documents; No Name tasks; Bridge Creek citations and hidden columns; the CEMP split; roles and owners; follow-ups; milestone dates flowing through; **every non-empty source cell appears in the output**; deterministic output |
| `pwsh import/tests/Test-Load.ps1` | Against a mock SharePoint (provisioned by the Phase 2 script): the dry run writes nothing and produces the plan; a full load creates all 641 items with correct lookups, people, links and Brisbane dates; a re-run makes **zero writes**; team edits survive; edited source fields are repaired; blank owners are filled; `-OverwriteLiveFields` works; a single list loads alone; an unknown person is warned and left blank |

The mock SharePoint copies the behaviours the loader depends on (typed field values, date-only storage, unique columns), but it isn't SharePoint. The dry run and the re-run in step 3.2 are the real proof on your tenant.
