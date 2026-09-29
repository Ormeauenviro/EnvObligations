# Phase 2 – SharePoint list schema

**Source of truth:** [`provision/list-schema.json`](../provision/list-schema.json). The column, choice and view tables at the end of this page are generated from it by `python provision/build_schema_docs.py`. So are the click-through checklists in [`manual-provisioning.md`](manual-provisioning.md), and `Provision-Lists.ps1` reads the same file, so the three can't drift apart. To change the schema, edit the JSON, re-run the generator, and re-run the provisioning script.

This schema implements all Phase 1 decisions (D-01 to D-16, accepted 29/09/2026).

---

## 1. Lists at a glance

| List | URL | Holds | Expected size | Links to parent by |
|---|---|---|---|---|
| **Obligations** | `Lists/Obligations` | One item per obligation (the register rows) | ≈ 230 | – |
| **Occurrences** | `Lists/Occurrences` | One item per period of a recurring or milestone obligation | ≈ 1,050 a year, so **past 2,000 in year 2** | `Obligation` lookup + `ObligationID` text |
| **Obligation Tasks** | `Lists/ObligationTasks` | Tick-off checklist items | Hundreds | `Obligation` lookup + `ObligationID` text |
| **Evidence** | `Lists/Evidence` | Evidence links (attachments allowed for site photos) | ≈ 80 at import, then grows | `ObligationID` (+ `OccurrenceKey`) text |
| **Follow-ups** | `Lists/FollowUps` | Missing, overdue, evidence-gap and import-review follow-ups | Hundreds | `ObligationID` (+ `OccurrenceKey`) text |
| **Status History** | `Lists/StatusHistory` | Append-only log of every status change | Thousands, **past 2,000** | `ObligationID` text + `Title` = item key |
| **Source Documents** | `Lists/SourceDocuments` | _Doc Ref, plus permits and plans cited in the register (D-14) | ≈ 40 | Obligations → `SourceDocument` lookup |
| **Report Components** | `Lists/ReportComponents` | Reporting Matrix, one item per ✕ | 28 | `ProducingObligationID` text |
| **Role Assignments** | `Lists/RoleAssignments` | Responsible Role → people (D-11) | 10 | Obligations `ResponsibleRole` = `Title` |
| **Project Milestones** | `Lists/ProjectMilestones` | Key dates that anchor due-date rules (D-10) | 6 | Obligations `DueAnchor` = `Title` |

**Relationships:**

```
Source Documents ◄──lookup── Obligations ◄──lookup + ObligationID── Occurrences
                                 ▲    ▲  ◄──lookup + ObligationID── Obligation Tasks
          Role Assignments ──────┘    │  ◄──ObligationID / OccurrenceKey── Evidence, Follow-ups, Status History
 (ResponsibleRole = Title)            │  ◄──ProducingObligationID── Report Components
          Project Milestones ─────────┘
 (DueAnchor = Title)
```

Only the brief's parent links are lookup columns. Every list uses at most 2 lookup columns and 3 person columns, well under the 12-per-view lookup threshold.

## 2. Keys

| Key | Format | Examples | Used for |
|---|---|---|---|
| **Obligation ID** (Obligations `Title`, unique) | `<prefix>-<nnn>[suffix]`, numbered in sheet order (D-15) | `CEMP-001`, `BC-DA-012`, `MP-NAC-006B` | Upsert key for the loader, and the text link used by every child list |
| **Occurrence Key** (Occurrences `Title`, unique) | `<Obligation ID>\|<period code>` | `CEMP-009\|2026-09`, `CEMP-021\|2026-W40`, `BC-DA-041\|Y2`, `NN-DA-002\|COMPLETE` | Duplicate guard for the occurrence generator (a create fails on the unique index), and the child link for Evidence and Follow-ups |
| **Document Key** (Source Documents, unique) | Short code | `SDA-2402-39300`, `MD-13-320`, `CMP-OMF` | Loader upsert key |

**Period codes:**

| Frequency | Period code | Label example |
|---|---|---|
| Weekly | `YYYY-Www` (ISO week, Monday to Sunday) | Week ending 04/10/2026 |
| Monthly | `YYYY-MM` | September 2026 |
| Quarterly | `YYYY-Qn` (calendar quarter) | Jul–Sep 2026 |
| Six-monthly | `YYYY-Hn` | Jul–Dec 2026 |
| Annual | `YYYY` | 2026 |
| Milestone | `Y1`, `Y2`, `Y5` | Year 1 |
| Notification split (D-05) | `COMMENCE`, `COMPLETE` | Completion notice |

Daily obligations roll up into one weekly occurrence: `YYYY-Www`, labelled *Daily checks – week ending 04/10/2026*.

## 3. How the status rules map onto columns

| Rule | Implementation |
|---|---|
| Statuses | Obligations `Status` is one of Not started / Underway / Complete / Missing / N/A – Info only. Occurrences use the same values without *N/A – Info only*. A period that isn't needed (for example, no discharge that month) is set to *Complete* with a written justification. |
| **Missing** needs a reason and creates a Follow-up | The app requires `StatusComment` before saving Missing. Flow 1 creates a Follow-up (`ReasonType` = Missing, `Reason` = the comment, `Owner` = the obligation owner, `FollowUpDate` = today + 7), then sets `HasOpenFollowUp` = Yes. |
| Resolving a Follow-up needs a comment | The app requires `ResolutionComment` and sets `ResolvedBy` and `ResolvedDate`. Flow 4 clears `HasOpenFollowUp` when no Open follow-ups remain for that Obligation ID or Occurrence Key. |
| **Overdue** is derived | Nothing stores Overdue as a status. Flow 2 sets `IsOverdue` each morning on items where Due Date / Period Due < today, the item isn't Complete or N/A, and `Active` = Yes. The app also works it out live, so an item is overdue on screen straight away. |
| **Complete** needs evidence or a justification | The app blocks Complete unless an Evidence item exists for the Obligation ID or Occurrence Key (or `EvidenceURL` is set on the occurrence), or `StatusComment` holds a justification. |
| **Compliance Outcome** is separate | `ComplianceOutcome` (Not assessed / Compliant / Non-compliant) is independent of `Status`. |
| Recurring obligations roll up | When an occurrence changes, the app (and flows 2 and 3) copy the *current* occurrence's status and Period Due into the obligation's `Status` and `DueDate`, and its key into `CurrentOccurrenceKey`. The current occurrence is the earliest one that isn't Complete, or the latest one if all are Complete. |
| Every change is logged | Every status change (app, flow or import) writes one **Status History** item: item key, type, old and new status, who, when (`ChangedAt`, date and time) and the comment. |
| Once-off due dates from project milestones (D-10) | When `DueAnchor` is set, Due Date = the milestone's Actual Date (or Planned Date if there is no actual) + `DueOffsetDays`. The loader calculates it, and a Phase 5 flow recalculates when a milestone date changes. |

## 4. Delegation and index plan

Power Apps only *delegates* (sends to SharePoint) filters built from supported operators. Anything else is evaluated on the first 500–2,000 rows only and silently misses the rest. Every filter the app needs is designed as a delegable expression on an **indexed** column. The indexes also keep list views and flow queries working once a list passes SharePoint's 5,000-item view threshold.

| Screen need | Delegable expression | Column(s) indexed |
|---|---|---|
| Register – register, stage, status, frequency, role | `SourceRegister.Value = x` (and so on; choice `=` is delegable) | SourceRegister, ProjectStage, Status, Frequency, ResponsibleRole |
| Register – owner | `Owner.Email = x` | Owner |
| Register – risk, including "not rated" | `RiskSort = n` (A=1 … D=4, not rated=5). This avoids `IsBlank()` on a choice column, which isn't delegable. | RiskSort |
| Register – Overdue / Missing flags | `IsOverdue = true`; `Status.Value = "Missing" \|\| HasOpenFollowUp = true` | IsOverdue, Status, HasOpenFollowUp |
| Register – text search | `StartsWith(Title, txt) \|\| StartsWith(RequirementSummary, txt)` | Title, RequirementSummary |
| Register – sort | `SortByColumns` on DueDate, SourceOrder, RiskSort or Title. Sorting on a choice column isn't delegable, which is another reason `RiskSort` exists. | DueDate, RiskSort |
| Detail tabs (Occurrences, Tasks, Evidence, Follow-ups, History) | `ObligationID = varOb.Title` | ObligationID on every child list |
| My tasks | `Assignee.Email = me && PeriodDue <= Today()+7 && (Status.Value = "Not started" \|\| Status.Value = "Underway" \|\| Status.Value = "Missing")` | Assignee, PeriodDue, Status |
| Follow-up queue | `Status.Value = "Open"`, sorted by `RiskSort` then `Created` (age) | Status, RiskSort |
| Occurrence generator (flow) | `Title eq '<key>'` to check for an existing occurrence | Title (unique) |

**Delegation limits and how I handle them.** These trade-offs are for you to confirm in Phase 4.

1. **"Contains" keyword search isn't delegable** on SharePoint. `Search()`, `in` and substring matching all evaluate locally. The delegable search above matches the *start* of the Obligation ID or the requirement summary. Because the **Obligations** list will stay around 230–300 items (far below the 2,000 data-row limit), I'll also offer an optional *keyword* box on the Register. It shows a delegation warning but is complete in practice, and a guard warns if the list ever nears 2,000. This is never used on Occurrences or Status History.
2. **`CountRows` isn't delegable.** Dashboard counts come from **Obligations** (small), and occurrence counts use date-windowed queries (next 14 or 30 days, overdue). Each of those returns far fewer than 2,000 rows.
3. **Filtering on lookup or person-multi columns isn't used anywhere.** Child lists carry `ObligationID` as indexed text, and `AdditionalPeople` (multi-person) is for notifications only.
4. **`<>` on a choice column** is avoided in favour of OR-ed `=` tests, which delegate reliably.

## 5. Design choices you should know about

| Choice | Why | Trade-off |
|---|---|---|
| The built-in **Title** column holds each list's key (Obligation ID, Occurrence Key and so on) | Title is always present, indexed-capable and required. It's also the column Power Apps and flows handle most easily. | In Power Apps the column is still called `Title` in formulas. The display name shows as *Obligation ID*. |
| **Copies** of the requirement summary (`ObligationSummary`) and risk (`RiskSort`) on Occurrences, Tasks and Follow-ups | Galleries and the follow-up queue can show and sort by them without one lookup per row, which is slow and not delegable | If an obligation's summary or risk changes, the app updates the copies on its open children. |
| `IsOverdue` flag, maintained daily | Makes "Overdue" a simple delegable filter in views, flows and emails | Can be up to a day stale. The app works it out live as well. |
| Single-choice **Project Stage** and **Responsible Role**, plus a **Supporting Roles** text column (D-03, D-11) | Filtering multi-choice columns isn't delegable | Items with several roles are filtered by their primary role only. |
| **Project Milestones** list, with `DueAnchor` and `DueOffsetDays` on Obligations (extra to the brief) | "On completion" and Year 1/2/5 dates (D-10) can then be recalculated from one place when a milestone date changes | One extra small list and two columns |
| `Active` (Yes/No) on Obligations | Lets you close out an obligation (for example after construction) so the generator and overdue check skip it without deleting it | – |
| **Attachments on only for Evidence** | Site users can attach a photo from a phone | Evidence normally stays in SharePoint document libraries as links. Attachments are stored in the list. |
| The Tasks list is named **Obligation Tasks** (`Lists/ObligationTasks`) | Avoids confusion with Planner/To Do "Tasks" in Power Apps and Power Automate pickers | – |
| Lookups use **restrict delete** | An obligation with occurrences or tasks can't be deleted, so the audit trail is kept | To retire an obligation, set `Active` = No. |
| **500 major versions** per item on every list | Full audit trail of edits alongside Status History | Version storage counts against the site quota (it's small for list items). |
| **Status History** is a normal list | M365 standard licensing only | It isn't tamper-proof. Optionally restrict edit rights (see [`manual-provisioning.md`](manual-provisioning.md) step 5). Versioning records any edit. |

## 6. Localisation

- **Site regional settings:** English (Australia) (DD/MM/YYYY), time zone (UTC+10:00) Brisbane, which has no daylight saving, and Monday as the first day of the week. `Provision-Lists.ps1 -SetRegionalSettings` sets these; the manual steps are in [`manual-provisioning.md`](manual-provisioning.md) step 0.
- All due and period dates are **Date only** columns, so they don't shift when viewed from another time zone. Only `Status History › Changed At` stores a time.
- Flows convert `utcNow()` with `convertFromUtc(utcNow(), 'E. Australia Standard Time')` before any date comparison (Phase 5).
- Display names and choices use Australian spelling (for example "Follow-up" and "Six-monthly").

---

## 7. Lists, columns and views (generated)

"Index" = indexed column. "Unique" = unique values enforced, which also implies an index.

<!-- BEGIN GENERATED: lists -->

### Source Documents

`Lists/SourceDocuments` · Source document register (from TP002 _Doc Ref plus permits and plans cited in the register).

Versioning on · attachments off · 3 indexed columns (limit 20) · 0 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Document Title | `Title` | Single line of text | Yes |  |  | Built-in Title column, renamed |
| Document Key | `DocKey` | Single line of text | Yes | Unique |  | Stable key used for upserts, e.g. SDA-2402-39300, MD-13-320. |
| Category | `Category` | Choice (drop-down) |  | Yes |  | [DocCategory](#choice-doccategory) |
| # / Reference | `Reference` | Single line of text |  |  |  |  |
| Version | `Version` | Single line of text |  |  |  |  |
| Document Date | `DocDate` | Date (date only) |  |  |  |  |
| Document Link | `DocLink` | Hyperlink |  |  |  |  |
| Folder Link | `FolderLink` | Hyperlink |  |  |  |  |
| Provided | `Provided` | Yes/No |  | Yes | Yes | No = listed in red on _Doc Ref (not yet provided). |
| Comments | `Comments` | Multiple lines of text (plain) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Document Key, Document Title, Category, # / Reference, Version, Document Date, Provided, Document Link | Sort: **Category** ascending, then **Document Key** ascending |
| Not Provided |  | Document Key, Document Title, Category, # / Reference, Comments | Filter: Show items only when **Provided** is equal to `No` |

### Project Milestones

`Lists/ProjectMilestones` · Key project dates that anchor due-date rules (D-10).

Versioning on · attachments off · 1 indexed columns (limit 20) · 0 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Milestone | `Title` | Single line of text | Yes | Unique |  | Built-in Title column, renamed |
| Planned Date | `PlannedDate` | Date (date only) |  |  |  |  |
| Actual Date | `ActualDate` | Date (date only) |  |  |  |  |
| Notes | `Notes` | Multiple lines of text (plain) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Milestone, Planned Date, Actual Date, Notes | Sort: **Planned Date** ascending |

### Role Assignments

`Lists/RoleAssignments` · Maps each Responsible Role to real people. Obligation Owner is resolved from Primary Person.

Versioning on · attachments off · 2 indexed columns (limit 20) · 2 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Responsible Role | `Title` | Single line of text | Yes | Unique |  | Built-in Title column, renamed |
| Primary Person | `PrimaryPerson` | Person (single) |  | Yes |  | Becomes the Owner of obligations with this role. |
| Additional People | `AdditionalPeople` | Person (multiple) |  |  |  | Copied on notifications. Not used for filtering. |
| Notes | `Notes` | Multiple lines of text (plain) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Responsible Role, Primary Person, Additional People, Notes | Sort: **Responsible Role** ascending |

### Obligations

`Lists/Obligations` · Environmental compliance obligations (replaces the TP002.1 register sheets).

Versioning on · attachments off · 16 indexed columns (limit 20) · 2 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Obligation ID | `Title` | Single line of text | Yes | Unique |  | Built-in Title column, renamed |
| Source Register | `SourceRegister` | Choice (drop-down) | Yes | Yes |  | [SourceRegister](#choice-sourceregister) |
| Source Order | `SourceOrder` | Number (0 decimals) |  |  |  | Row order within the source register; used for sorting and the TP002 export. |
| Source Cell | `SourceCell` | Single line of text |  |  |  | Original workbook location, e.g. 'Bridge Creek DA Conditions'!A14:N14. |
| Section/Group | `SectionGroup` | Single line of text |  |  |  |  |
| Condition # | `ConditionNo` | Single line of text |  |  |  |  |
| Parent Obligation ID | `ParentObligationID` | Single line of text |  |  |  | Set on sub-conditions, e.g. MP-NAC-006B has parent MP-NAC-006. |
| Project Location/Activity | `LocationActivity` | Single line of text |  |  |  |  |
| Project Stage | `ProjectStage` | Choice (drop-down) |  | Yes |  | [ProjectStage](#choice-projectstage) |
| Requirement Summary | `RequirementSummary` | Single line of text | Yes | Yes |  | First 250 characters of the requirement; used for delegable StartsWith search. |
| Requirement | `Requirement` | Multiple lines of text (plain) | Yes |  |  |  |
| Source Reference | `SourceReference` | Multiple lines of text (plain) |  |  |  |  |
| Source Document | `SourceDocument` | Lookup |  | Yes |  | → Source Documents (Title); restrict delete |
| Compliance Action/Controls | `ComplianceAction` | Multiple lines of text (plain) |  |  |  |  |
| Practical Implementation | `PracticalImplementation` | Multiple lines of text (plain) |  |  |  |  |
| Evidence Required | `EvidenceRequired` | Multiple lines of text (plain) |  |  |  | What evidence demonstrates compliance (D-02). |
| Stakeholder Review/Approval/Submission | `Stakeholder` | Multiple lines of text (plain) |  |  |  |  |
| Risk Rating | `RiskRating` | Choice (drop-down) |  |  |  | [RiskRating](#choice-riskrating). Per DTI-RM-ST001. Blank = not rated. |
| Risk Sort | `RiskSort` | Number (0 decimals) |  | Yes |  | A=1, B=2, C=3, D=4, not rated=5. Maintained by the app and loader for delegable sorting. |
| Frequency | `Frequency` | Choice (drop-down) | Yes | Yes |  | [Frequency](#choice-frequency) |
| Timing Detail | `TimingDetail` | Multiple lines of text (plain) |  |  |  |  |
| Recurrence Start | `RecurrenceStart` | Date (date only) |  |  |  | First period the occurrence generator creates (recurring obligations). |
| Recurrence End | `RecurrenceEnd` | Date (date only) |  |  |  | Last period to generate; blank = until Active is set to No. |
| Due Anchor | `DueAnchor` | Choice (drop-down) |  |  |  | [DueAnchor](#choice-dueanchor). Project milestone a once-off due date is calculated from. |
| Due Offset (days) | `DueOffsetDays` | Number (0 decimals) |  |  |  | Once/Milestone: days after the Due Anchor. Recurring: days after period end that each occurrence is due. |
| Due Date | `DueDate` | Date (date only) |  | Yes |  | Real due date. For recurring obligations, the current occurrence's Period Due. |
| Due Rule | `DueRule` | Single line of text |  |  |  | Original due-date text from TP002. |
| Owner | `Owner` | Person (single) |  | Yes |  |  |
| Responsible Role | `ResponsibleRole` | Choice (drop-down) |  | Yes |  | [ResponsibleRole](#choice-responsiblerole) |
| Supporting Roles | `SupportingRoles` | Multiple lines of text (plain) |  |  |  |  |
| Status | `Status` | Choice (drop-down) | Yes | Yes | Not started | [Status](#choice-status) |
| Status Reason/Justification | `StatusComment` | Multiple lines of text (plain) |  |  |  | Reason for Missing, or justification for Complete without an evidence link. |
| Compliance Outcome | `ComplianceOutcome` | Choice (drop-down) | Yes | Yes | Not assessed | [ComplianceOutcome](#choice-complianceoutcome) |
| Compliance Notes | `ComplianceNotes` | Multiple lines of text (plain) |  |  |  |  |
| Last Reviewed | `LastReviewed` | Date (date only) |  |  |  |  |
| Info Only | `InfoOnly` | Yes/No |  | Yes | No |  |
| Active | `Active` | Yes/No |  | Yes | Yes | No = closed out; the occurrence generator and overdue check skip it. |
| Has Open Follow-up | `HasOpenFollowUp` | Yes/No |  | Yes | No | Maintained by flows. |
| Is Overdue | `IsOverdue` | Yes/No |  | Yes | No | Maintained by the daily overdue flow. Overdue is derived, never a status. |
| Current Occurrence Key | `CurrentOccurrenceKey` | Single line of text |  |  |  | Occurrence the status rolls up from (recurring obligations). |
| Extra Fields | `ExtraFields` | Multiple lines of text (plain) |  |  |  | Labelled text for workbook content with no dedicated column. |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Obligation ID, Source Register, Condition #, Requirement Summary, Project Stage, Frequency, Due Date, Owner, Responsible Role, Status, Compliance Outcome, Risk Rating, Has Open Follow-up, Is Overdue | Sort: **Source Register** ascending, then **Source Order** ascending |
| Open |  | Obligation ID, Source Register, Requirement Summary, Frequency, Due Date, Owner, Status, Risk Rating, Is Overdue | Filter: Show items only when **Status** is equal to `Not started` OR **Status** is equal to `Underway` OR **Status** is equal to `Missing`<br>Sort: **Due Date** ascending |
| Overdue |  | Obligation ID, Source Register, Requirement Summary, Due Date, Owner, Status, Risk Rating | Filter: Show items only when **Is Overdue** is equal to `Yes`<br>Sort: **Due Date** ascending |
| Missing or Follow-up |  | Obligation ID, Source Register, Requirement Summary, Owner, Status, Status Reason/Justification, Has Open Follow-up | Filter: Show items only when **Has Open Follow-up** is equal to `Yes` OR **Status** is equal to `Missing`<br>Sort: **Risk Sort** ascending |
| By Register |  | Obligation ID, Section/Group, Condition #, Requirement Summary, Frequency, Owner, Status, Compliance Outcome | Sort: **Source Order** ascending<br>Group By: **Source Register** (collapsed) |
| My Obligations |  | Obligation ID, Source Register, Requirement Summary, Frequency, Due Date, Status, Is Overdue | Filter: Show items only when **Owner** is equal to `[Me]`<br>Sort: **Due Date** ascending |
| Info Only |  | Obligation ID, Source Register, Requirement Summary, Source Reference | Filter: Show items only when **Info Only** is equal to `Yes`<br>Sort: **Source Order** ascending |

### Occurrences

`Lists/Occurrences` · One item per period for recurring and milestone obligations. Expected to exceed 2,000 items.

Versioning on · attachments off · 9 indexed columns (limit 20) · 2 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Occurrence Key | `Title` | Single line of text | Yes | Unique |  | Built-in Title column, renamed |
| Obligation | `Obligation` | Lookup | Yes | Yes |  | → Obligations (Title); restrict delete |
| Obligation ID | `ObligationID` | Single line of text | Yes | Yes |  | Copy of the obligation's ID for delegable filtering. |
| Obligation Summary | `ObligationSummary` | Single line of text |  |  |  | Copy of Requirement Summary so galleries avoid per-row lookups. |
| Period Label | `PeriodLabel` | Single line of text | Yes |  |  |  |
| Period Start | `PeriodStart` | Date (date only) | Yes | Yes |  |  |
| Period End | `PeriodEnd` | Date (date only) |  |  |  |  |
| Period Due | `PeriodDue` | Date (date only) | Yes | Yes |  |  |
| Assignee | `Assignee` | Person (single) |  | Yes |  |  |
| Status | `Status` | Choice (drop-down) | Yes | Yes | Not started | [OccurrenceStatus](#choice-occurrencestatus) |
| Status Reason/Justification | `StatusComment` | Multiple lines of text (plain) |  |  |  |  |
| Compliance Outcome | `ComplianceOutcome` | Choice (drop-down) | Yes |  | Not assessed | [ComplianceOutcome](#choice-complianceoutcome) |
| Completed Date | `CompletedDate` | Date (date only) |  |  |  |  |
| Evidence URL | `EvidenceURL` | Hyperlink |  |  |  |  |
| Notes | `Notes` | Multiple lines of text (plain) |  |  |  |  |
| Risk Sort | `RiskSort` | Number (0 decimals) |  | Yes |  |  |
| Is Overdue | `IsOverdue` | Yes/No |  | Yes | No |  |
| Has Open Follow-up | `HasOpenFollowUp` | Yes/No |  |  | No |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Occurrence Key, Obligation ID, Period Label, Period Due, Assignee, Status, Completed Date, Evidence URL, Is Overdue | Sort: **Period Due** descending |
| Open |  | Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Assignee, Status, Is Overdue | Filter: Show items only when **Status** is equal to `Not started` OR **Status** is equal to `Underway` OR **Status** is equal to `Missing`<br>Sort: **Period Due** ascending |
| Due Next 30 Days |  | Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Assignee, Status | Filter: Show items only when **Period Due** is greater than or equal to `[Today]` AND **Period Due** is less than or equal to `[Today]+30`<br>Sort: **Period Due** ascending |
| Overdue |  | Occurrence Key, Obligation ID, Obligation Summary, Period Due, Assignee, Status | Filter: Show items only when **Is Overdue** is equal to `Yes`<br>Sort: **Period Due** ascending |
| My Occurrences |  | Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Status | Filter: Show items only when **Assignee** is equal to `[Me]`<br>Sort: **Period Due** ascending |

### Obligation Tasks

`Lists/ObligationTasks` · Checklist items that can be ticked off against an obligation.

Versioning on · attachments off · 5 indexed columns (limit 20) · 2 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Task | `Title` | Single line of text | Yes |  |  | Built-in Title column, renamed |
| Obligation | `Obligation` | Lookup | Yes | Yes |  | → Obligations (Title); restrict delete |
| Obligation ID | `ObligationID` | Single line of text | Yes | Yes |  |  |
| Obligation Summary | `ObligationSummary` | Single line of text |  |  |  |  |
| Task Order | `TaskOrder` | Number (0 decimals) |  |  |  |  |
| Assignee | `Assignee` | Person (single) |  | Yes |  |  |
| Due Date | `DueDate` | Date (date only) |  | Yes |  |  |
| Done | `Done` | Yes/No |  | Yes | No |  |
| Done Date | `DoneDate` | Date (date only) |  |  |  |  |
| Notes | `Notes` | Multiple lines of text (plain) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Obligation ID, Task Order, Task, Assignee, Due Date, Done, Done Date | Sort: **Obligation ID** ascending, then **Task Order** ascending |
| Open Tasks |  | Obligation ID, Task, Assignee, Due Date | Filter: Show items only when **Done** is equal to `No`<br>Sort: **Due Date** ascending |
| My Tasks |  | Obligation ID, Task, Due Date, Done | Filter: Show items only when **Assignee** is equal to `[Me]` AND **Done** is equal to `No`<br>Sort: **Due Date** ascending |

### Evidence

`Lists/Evidence` · Links to evidence held in SharePoint, against an obligation and optionally an occurrence.

Versioning on · attachments on · 4 indexed columns (limit 20) · 1 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Title | `Title` | Single line of text | Yes |  |  | Built-in Title column, renamed |
| Obligation ID | `ObligationID` | Single line of text | Yes | Yes |  |  |
| Occurrence Key | `OccurrenceKey` | Single line of text |  | Yes |  |  |
| SharePoint URL | `EvidenceURL` | Hyperlink | Yes |  |  |  |
| Date | `EvidenceDate` | Date (date only) |  | Yes |  |  |
| Added By | `AddedBy` | Person (single) |  | Yes |  |  |
| Notes | `Notes` | Multiple lines of text (plain) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Obligation ID, Occurrence Key, Title, SharePoint URL, Date, Added By | Sort: **ID** descending |

### Follow-ups

`Lists/FollowUps` · Items needing follow-up: Missing obligations, overdue items, evidence gaps and import review.

Versioning on · attachments off · 7 indexed columns (limit 20) · 2 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Follow-up | `Title` | Single line of text | Yes |  |  | Built-in Title column, renamed |
| Obligation ID | `ObligationID` | Single line of text | Yes | Yes |  |  |
| Occurrence Key | `OccurrenceKey` | Single line of text |  | Yes |  |  |
| Obligation Summary | `ObligationSummary` | Single line of text |  |  |  |  |
| Reason Type | `ReasonType` | Choice (drop-down) | Yes | Yes | Missing | [FollowUpReasonType](#choice-followupreasontype) |
| Reason | `Reason` | Multiple lines of text (plain) | Yes |  |  |  |
| Owner | `Owner` | Person (single) |  | Yes |  |  |
| Follow-up Date | `FollowUpDate` | Date (date only) |  | Yes |  |  |
| Status | `Status` | Choice (drop-down) | Yes | Yes | Open | [FollowUpStatus](#choice-followupstatus) |
| Resolution Comment | `ResolutionComment` | Multiple lines of text (plain) |  |  |  |  |
| Resolved By | `ResolvedBy` | Person (single) |  |  |  |  |
| Resolved Date | `ResolvedDate` | Date (date only) |  |  |  |  |
| Risk Sort | `RiskSort` | Number (0 decimals) |  | Yes |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Follow-up, Obligation ID, Reason Type, Owner, Follow-up Date, Status, Created | Sort: **Created** descending |
| Open |  | Follow-up, Obligation ID, Obligation Summary, Reason Type, Reason, Owner, Follow-up Date, Created | Filter: Show items only when **Status** is equal to `Open`<br>Sort: **Risk Sort** ascending, then **Created** ascending |
| My Open Follow-ups |  | Follow-up, Obligation ID, Reason Type, Reason, Follow-up Date | Filter: Show items only when **Owner** is equal to `[Me]` AND **Status** is equal to `Open`<br>Sort: **Follow-up Date** ascending |

### Status History

`Lists/StatusHistory` · Append-only log of every status change (app, flows and import).

Versioning on · attachments off · 4 indexed columns (limit 20) · 1 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Item ID | `Title` | Single line of text | Yes | Yes |  | Built-in Title column, renamed |
| Item Type | `ItemType` | Choice (drop-down) | Yes | Yes |  | [ItemType](#choice-itemtype) |
| Obligation ID | `ObligationID` | Single line of text | Yes | Yes |  |  |
| Old Status | `OldStatus` | Single line of text |  |  |  |  |
| New Status | `NewStatus` | Single line of text | Yes |  |  |  |
| Changed By | `ChangedBy` | Person (single) |  |  |  |  |
| Changed At | `ChangedAt` | Date and time | Yes | Yes |  |  |
| Comment | `Comment` | Multiple lines of text (plain) |  |  |  |  |
| Change Source | `ChangeSource` | Choice (drop-down) |  |  | App | [ChangeSource](#choice-changesource) |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Changed At, Item ID, Item Type, Obligation ID, Old Status, New Status, Changed By, Change Source, Comment | Sort: **Changed At** descending |
| Last 30 Days |  | Changed At, Item ID, Item Type, Old Status, New Status, Changed By, Comment | Filter: Show items only when **Changed At** is greater than or equal to `[Today]-30`<br>Sort: **Changed At** descending |

### Report Components

`Lists/ReportComponents` · Reporting Matrix: which content items go into which report.

Versioning on · attachments off · 2 indexed columns (limit 20) · 0 lookup/person columns (+ Created By/Modified By)

| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |
|---|---|---|---|---|---|---|
| Content Item | `Title` | Single line of text | Yes |  |  | Built-in Title column, renamed |
| Report Name | `ReportName` | Choice (drop-down) | Yes | Yes |  | [ReportName](#choice-reportname) |
| Report Frequency | `ReportFrequency` | Choice (drop-down) |  |  |  | [ReportFrequency](#choice-reportfrequency) |
| Report Source Reference | `ReportSourceRef` | Single line of text |  |  |  |  |
| Producing Obligation ID | `ProducingObligationID` | Single line of text |  | Yes |  |  |
| Sort Order | `SortOrder` | Number (0 decimals) |  |  |  |  |

| View | Default | Columns | Filter / sort / group |
|---|---|---|---|
| All Items | ✓ | Content Item, Report Frequency, Producing Obligation ID, Report Source Reference | Sort: **Sort Order** ascending<br>Group By: **Report Name** (expanded) |

### Choice sets

<a id="choice-sourceregister"></a>**SourceRegister**: CEMP and General Requirements · PFAS CMP Requirements · ADR WWBW Conditions · Bridge Creek DA Conditions · No Name DA Conditions · Marine Plants DA Conditions · No Name Operational Works · External CHMP Conditions · RPP Exemption · Specification Requirements · Legislation

<a id="choice-projectstage"></a>**ProjectStage**: All stages · Design · Pre-construction · Pre-construction and construction · Construction · Construction and operation · Operation · Completion · Post-construction · Rehabilitation and monitoring

<a id="choice-frequency"></a>**Frequency**: Once · Daily · Weekly · Monthly · Quarterly · Six-monthly · Annual · Milestone · Trigger-based · Ongoing

<a id="choice-status"></a>**Status**: Not started · Underway · Complete · Missing · N/A – Info only

<a id="choice-occurrencestatus"></a>**OccurrenceStatus**: Not started · Underway · Complete · Missing

<a id="choice-complianceoutcome"></a>**ComplianceOutcome**: Not assessed · Compliant · Non-compliant

<a id="choice-riskrating"></a>**RiskRating**: A · B · C · D

<a id="choice-responsiblerole"></a>**ResponsibleRole**: Environmental and Sustainability Team · Environmental Manager · Environmental Advisor · Project Manager · Construction Manager · Site Manager · Site Supervisor · Engineering and Design Team · Contractor · Asset Owner / Operator

<a id="choice-dueanchor"></a>**DueAnchor**: Construction start · Practical completion · Rehabilitation completion · Revegetation commencement · Temporary barrier removal · Restoration works completion

<a id="choice-followupstatus"></a>**FollowUpStatus**: Open · Resolved

<a id="choice-followupreasontype"></a>**FollowUpReasonType**: Missing · Overdue · Evidence gap · Import review · Other

<a id="choice-itemtype"></a>**ItemType**: Obligation · Occurrence · Task · Follow-up

<a id="choice-changesource"></a>**ChangeSource**: App · Flow · Import

<a id="choice-doccategory"></a>**DocCategory**: Specification · Heritage · WWBW · RPP Exemption · CMP · Development approval · Management plan · Contract · Standard · Legislation

<a id="choice-reportname"></a>**ReportName**: Monitoring data made available to TMR · PFAS Monthly Compliance Report · CMP Monthly Audit Report · Compliance Register · CMP Close Out Report

<a id="choice-reportfrequency"></a>**ReportFrequency**: Monthly · Post construction

<!-- END GENERATED: lists -->
