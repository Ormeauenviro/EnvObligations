# Phase 1 – Data audit and mapping

**Source:** `source/DTI-HSEQ-TP002.1_Environmental_Compliance_Obligations_Register_Rev2_DRAFT_DH.xlsx` (TP002.1 Rev 2 DRAFT, QTMP Ormeau Maintenance Facility)
**Status:** Approved 29/09/2026. **All decisions D-01 to D-16 were accepted as recommended.** Phase 2 implements them in [`list-schema.md`](list-schema.md).
**Evidence:** Every count and value below comes from `import/audit_workbook.py`. Its full output (the row-by-row classification, column fill counts, every distinct value, every hyperlink and every cell comment) is in [`data-audit-inventory.md`](data-audit-inventory.md). To regenerate it:

```bash
pip install -r requirements.txt
python import/audit_workbook.py
```

Decisions I need from you are marked **⚠ D-nn** and collected in [section 7](#7-decisions-needed-before-phase-2).

---

## 1. Headline findings

1. **The register holds 225 importable rows in 10 obligation sheets:** 208 obligations, 8 sub-conditions, 7 info-only ("Noted") rows and 2 rows I can't parse (section 6). The hidden legacy and stub rows on Marine Plants (4) and the blank rows are extra to that total.
2. **Only the three DA sheets have any progress data.** CEMP, PFAS, ADR, CHMP, RPP, Specification and Legislation have never had anything entered in Compliance Status, Compliance Notes, Due Date, Evidence or Last Reviewed. Once migrated, these 151 rows will start as *Not started / Not assessed* unless you tell me otherwise (**⚠ D-01**).
3. **Legislation isn't empty.** It has 16 rows, but only Requirement (C) and Reference/Act (D) are filled: 15 usable rows plus one Act with no requirement text. **No Name Operational Works is empty** (only A1 is formatted).
4. **The existing status drop-down isn't used on any sheet.** `Summary!A20:A25` (Compliant / Not Compliant / Completed / In Progress / Overdue / Not Started) is the validation source, but the DA sheets use free text such as *Not completed – ongoing*, *Ongoing* and *Not completed – rehabilitation*. The Summary COUNTIFs only count the Specification sheet, which a threaded comment on `Summary!A18` confirms. Summary figures therefore can't be trusted, and the Dashboard will replace them.
5. **Marine Plants DA Conditions has three blocks.** A hidden legacy block (rows 1–19 are hidden) is followed by one block per decision notice, each with its own group row and repeated header: 2308‑36075 SDA (TNAR) at rows 21–30 and 2309‑36966 SDA (NAC) at rows 32–45. The **legacy block uses a different 10-column layout**, and one current row (row 23) is shifted one column to the right (section 1.3).
6. **Evidence in the DA sheets is hyperlinks, not text.** There are 77 hyperlinks: 34 on Bridge Creek, 12 on No Name and 31 on Marine Plants (25 in data rows, 4 on the decision-notice group rows and 2 on the hidden legacy rows). Most are *relative* links (`../../../../:f:/s/DTI-20151005_QTMP/…`), and I resolve them to `https://dtinfrastructurecomau.sharepoint.com/:f:/s/…`. Many point to the same few folders (for example, *Weekly Environmental Inspection Checklist_Full Report* is linked 38 times). They are folder links, not item-level evidence.
7. **The "Evidence" column is mostly a description of the evidence expected** (for example, *Approved plans, site inspections* or *Survey records, as-constructed drawings*), not evidence itself. PFAS's *Compliance Check* column is the same kind of thing. I propose a separate **Evidence Required** column (**⚠ D-02**).
8. **Bridge Creek has 31 cells with pasted AI-citation fragments**, such as `[2402-39300...A Decision | PDF]`, and **55 cells that contain only non-breaking spaces**. Columns D (Source/Reference) and G (Responsibility) are **hidden** but hold data.
9. **Risk Rating is sparse.** It exists only on CHMP, RPP and Specification (44 rows: 4 × C, 40 × D). A and B are never used, and 181 of the 225 importable rows have no rating.
10. **Only 3 real dates exist in any Due Date column** (27/10/2023 and 28/10/2023 in the hidden legacy Marine Plants rows, and 31/07/2026 on NN condition 10). Every other value is a rule such as *Ongoing*, *Trigger based*, *Completion* or *Scheduled*. To turn those into dates I need project milestone dates (**⚠ D-10**).
11. **There are 36 threaded review comments** (plus 5 legacy notes on the Risk Rating headers) (reviewers: A. Brunott, R. Kightley). They hold open questions such as "EP Reg apply to this?" and "Is there a project audit schedule…?". I propose carrying them into Extra Fields, and optionally raising them as Follow-ups (**⚠ D-12**).

---

## 2. Sheet-by-sheet

Every obligation sheet puts group headers on row 4 and column headers on row 5, with data from row 6. The exceptions are Marine Plants' second and third blocks.

| Sheet | Purpose | Header row | Columns (row 5) | Rows with content | Obligations / sub / info / can't parse |
|---|---|---|---|---|---|
| Instructions | How to use the template. C30:C33 = Risk A–D (drop-down source) | – | – | – | reference |
| _Template | Blank standard 12-column layout | 5 | A–L (standard) | 0 | reference |
| Summary | Project details and COUNTIF totals. A20:A25 = status drop-down | – | – | – | reference |
| **CEMP and General Requirements** | EMP/CEMP annex, CMP and ISO 14001 monitoring, inspection and reporting | 5 (Excel table `Table13`, A5:L65) | 12 | 60 | 60 / 0 / 0 / 0 |
| **PFAS CMP Requirements** | PFAS CMP Rev C, Compliance Matters 1–3 | 5 | 14 | 19 | 19 / 0 / 0 / 0 |
| **ADR WWBW Conditions** | Accepted development requirements for waterway barrier works | 5 | 11 | 7 | 6 / 0 / 1 / 0 |
| **Bridge Creek DA Conditions** | SARA 2402‑39300 SDA conditions 1–14, plus Fish Salvage, Rehab & Remediation, Fish Passage Monitoring and Aquatic Plant Monitoring plan commitments | 5 | 13 (+ unheaded N) | 42 | 42 / 0 / 0 / 0 |
| **No Name DA Conditions** | SARA 2506‑46598 SDA conditions 1–11 | 5 | 14 | 12 | 11 / 1 / 0 / 0 |
| **Marine Plants DA Conditions** | SARA 2308‑36075 SDA (TNAR) and 2309‑36966 SDA (NAC) | 5 (legacy, hidden), 22, 33 | 10 (legacy) / 14 | 26 | 12 / 7 / 1 / 0 + 4 legacy |
| No Name Operational Works | Placeholder | – | – | 0 | empty |
| **External CHMP Conditions** | Jabree CHMP and Danggan Balun Partial Clearance Report | 5 | 12 (standard) | 14 | 11 / 0 / 3 / 0 |
| **RPP Exemption** | Riverine protection permit exemption (WSS/2013/726) | 5 | 12 (standard) | 3 | 2 / 0 / 0 / 1 |
| **Specification Requirements** | Subcontract, Sch A5, Sch C1, MD‑13‑320, MD‑15‑3 | 5 | 12 (standard) | 32 | 30 / 0 / 2 / 0 ¹ |
| Reporting Matrix | 20 report content items × 5 reports (MONTHLY ×4, POST CONSTRUCTION ×1) | 4 | A + B–F | 20 | reference → Report Components |
| **Legislation** | Acts and regulations (only C and D filled) | 5 | 12 (standard) | 16 | 15 / 0 / 0 / 1 |
| _Doc Ref | Source document register. Red text = document not yet provided | 4 | 6 | 13 | reference → Source Documents |

¹ The 2 info-only rows are rows 11 and 15. Rows 35 and 37 read "Noted – relevant to rehabilitation / landscaping plan" and have no Responsibility. The script counts them among the 30 obligations, but they may be info-only (**⚠ D-07**).

### 2.1 Header variations (same meaning, different wording)

| Canonical field | Header variants found |
|---|---|
| Requirement | Compliance Requirement · Requirement · ADR Compliance Requirement (Bridge Creek, copied from ADR) · ADR Compliance Requirement Summary (No Name, Marine Plants) |
| Source Reference | Reference · Source/ Reference |
| Compliance Action/Controls | Compliance Controls · Compliance Action · "Compliance Action " (trailing space) |
| Timing | Timing / Frequency · Timing/ Trigger · *(absent on PFAS, CHMP, RPP, Spec, Legislation)* |
| Section/Group | Management Plan/ Subplan (CEMP col A) · Compliance Matter (PFAS col C) |
| Evidence | Evidence · "Evidence " · Supporting Documents |
| Links | Links · Link · *(unheaded Bridge Creek col N)* |
| PFAS only | Relevant Legislation / Standard or Code · Practical Implementation · Compliance Check |
| CEMP only | Compliance last reviewed |
| Standard layout only | Risk Rating · Stakeholder Review/ Approval/ Submission *(never filled on any sheet)* |

### 2.2 Group and heading rows

The brief expected section/group heading rows between data rows. **None exist on any sheet apart from Marine Plants.** Grouping is carried in a column instead:

- **CEMP col A** (*Management Plan/Subplan*) has 13 groups, such as *1. Compliance Management Plan*, *Acid sulfate soil* and *Air quality and dust*. Two cells hold two lines (*General ⏎ Legal Compliance*); I keep the first line as the group and the rest in Extra Fields. Col A → **Section/Group**.
- **PFAS col C** (*Compliance Matter 1/2/3: …*) → **Section/Group**. PFAS col A holds one sentence repeated on all 19 rows, which goes to **Project Location/Activity**.
- **DA sheets**: the decision notice/permit is the grouping, stored as **Source Document** (lookup) plus the ID prefix.

### 2.3 Marine Plants DA Conditions – parsed structure

| Rows | What it is | Proposed handling |
|---|---|---|
| 1–19 | **Hidden.** Title and the old row-5 header, which uses the **10-column layout** (A Location, B Stage, C Requirement, D Reference, E Action, F Responsibility, G Due Date, H Status, I Notes, J Evidence) | Header only |
| 6 | Legacy summary: TNAR, *SARA Approval to Disturb Marine Plants*, 2308‑36075 SDA, due 27/10/2023, *Completed*, J = hyperlink "Approval" | **Don't import as an obligation.** Becomes the **Source Document** row for 2308‑36075 SDA (with the approval link). The notification action is covered by condition 2 below. |
| 7 | Legacy summary: TNAR, 2309‑**36996** SDA, due 28/10/2023, *Compliant*, J = hyperlink "Approval" | Same as row 6, but the permit number **conflicts** with row 32 (2309‑**36966**) and A says TNAR although that permit is NAC (**⚠ D-08**) |
| 8, 9 | Stubs: "NAC" in col A only | Drop |
| 10–20 | Blank (hidden) | Skip |
| 21 | **Group row, TNAR:** B = hyperlink to *2308‑36075 SDA – Decision notice.pdf*, C = hyperlink "Decision documents" | Context for rows 23–30. Both links go on the Source Document item. |
| 22 | **Repeated header** in the 14-column DA layout (same as No Name) | Header for block 2 |
| 23–30 | TNAR conditions 1, 2, 3, 4, 5a, 5b, 6a, 6b | 6 obligations + 2 sub-conditions (5b, 6b) |
| 31 | Blank | Skip |
| 32 | **Group row, NAC:** B = hyperlink to *2309‑36966 SDA – Decision – approval with conditions.pdf*, C = "Decision documents" | Context for rows 34–45 |
| 33 | **Repeated header** | Header for block 3 |
| 34–45 | NAC conditions 1, 2 (parent, status *noted*), 2a, 2b, 2c, 3, 4, 5, 6a, 6b, 7a, 7b | 6 obligations + 5 sub-conditions (2a, 2b, 2c, 6b, 7b) + 1 info-only (2) |

**Column misalignment, row 23 (TNAR condition 1):** from column J onward the values are shifted one column right compared with every other row and with the identical No Name condition 1.

| Col | Header | Row 23 value | What it actually is |
|---|---|---|---|
| J | Compliance Status | Ongoing | stray value (duplicate of the due-date idea) |
| K | Compliance Notes | Not completed – ongoing | **status** |
| L | Evidence | Compare to GIS as evidence | **notes** |
| M | Link | 2. Marine Plant DAs (hyperlink) | link |
| N | Link | Weekly Environmental Inspection Checklist (hyperlink) | link |

Proposed fix: status = K, notes = L, Evidence Required = blank, and the stray J value goes to Extra Fields as `Unmapped (J): Ongoing` (**⚠ D-08**).

**Inherited values on sub-condition rows:** rows 28, 30, 36–38, 43 and 45 leave Location and Stage blank. They inherit both from the parent condition row. The literal text "Same project" in col A is replaced with the block's location (row 23 or row 34).

**Links in column N on rows 36–38** swap M and N compared with other rows. That doesn't matter, because every hyperlinked cell becomes an Evidence item whichever column it's in.

### 2.4 Other sheet-specific issues

| Sheet | Row/Cell | Issue | Proposed handling |
|---|---|---|---|
| Bridge Creek | C, D, F (31 cells) | Pasted citation fragments `[2402-39300...A Decision \| PDF]` and `[…itions (1) \| PDF]` | Strip from the text and keep the originals in Extra Fields → `Citations:` (**⚠ D-09**) |
| Bridge Creek | I and J, rows 22–49 (55 cells) | Cells that contain only a non-breaking space | Treat as blank |
| Bridge Creek | Cols D and G | Hidden columns (Source/Reference, Responsibility) | Import normally |
| Bridge Creek | N12 | Value in an unheaded column: DAF letter quote *"The Department believes that appropriate steps will be taken…"* | Append to Compliance Notes as `DAF correspondence: …` |
| Bridge Creek | Rows 20–21 | Blank gap between condition 14 and the Fish Salvage Plan rows | Skip. Confirm SARA 2402‑39300 has only 14 conditions (**⚠ D-08**) |
| Bridge Creek | D7–D19 | Source Reference is just "Condition 2" … "Condition 14" | Condition # is parsed out of it, and Source Document = 2402‑39300 SDA |
| No Name | Row 7 (C = `1:a-d`, D7:I7 merged) | Sub-row that lists the four works (a–d) under condition 1. Status "(a) – completed" | Four **Tasks** on NN condition 1, with (a) marked Done (**⚠ D-06**) |
| No Name, Marine Plants | NN L10, MP L26, MP L41 | The Evidence column holds the DAF letter quote instead of an evidence description | Move the quote to Compliance Notes. Leave Evidence Required blank. |
| ADR | A6 | Location is an editorial note: *"To change all of these to Creek crossings that required ADR WWBW"* | Location blank. Note goes into Extra Fields and a Follow-up (**⚠ D-12**) |
| ADR | Col F (Timing) | Blank on every row | Frequency comes from the requirement text (section 4.3) |
| CEMP | Row 11 | Two deliverables in one row (Weekly inspection checklist + CMP Monthly Audit; timing "Weekly ⏎ Monthly") | Split into CEMP‑xxx and CEMP‑xxx‑B (**⚠ D-05**) |
| CEMP | Row 28 | Stage *Pre construction* but timing *Post construction* | Flag. Keep Stage and put Timing in Timing Detail. |
| CEMP | Row 41 | Timing *"Quarterly / Six monthly?"* | Needs your answer (**⚠ D-04**) |
| CEMP | Rows 26, 49–52 | No Compliance Action | Import blank |
| CEMP | Row 60 | No Responsibility | Default role = Environmental and Sustainability Team (**⚠ D-11**) |
| PFAS | All rows | No Timing, Due Date, Status, Notes or Evidence. Col E (Legislation) and col K (Compliance Check) are fully populated. | Legislation → Extra Fields. Compliance Check → Evidence Required (**⚠ D-02**) |
| RPP | C6 | Requirement is just `?)` (truncated) | Can't parse (section 6) |
| Legislation | C6–C18 | Five cells start with a tab | Trimmed |
| Legislation | Row 21 | *EPBC Act – significant impact identified during works* in D, no requirement | Can't parse (section 6) |
| Specification | Rows 35, 37 | "Noted – relevant to rehabilitation / landscaping plan", no Responsibility | **⚠ D-07** |
| All | – | 209 multi-line cells, 201 cells with leading or trailing whitespace | Keep line breaks (SharePoint multi-line plain text). Trim each line. |
| External CHMP | Used range to row 48 | Formatting only below row 19 | Ignore |
| Workbook | External link | Link to Downer *DG-DM-TP008 Change Register.xlsm* (not used by any formula) | Ignore |

### 2.5 Non-obligation sheets

- **Instructions** gives the Risk list A, B, C, D and cites **DTI-RM-ST001**. The Risk Rating header comment on each sheet cites **DA-ZH-PR028 / Annex B** instead, and the RPP comment cites DG-ZH-FM028.1. The choice values are A–D either way. Only the help text differs; I'll use DTI-RM-ST001 as you specified.
- **_Template** defines the standard 12-column layout. It is used for the export layout (Phase 5).
- **Summary** holds project details (Project QTMP – Ormeau Maintenance Facility, Location Ormeau, Client Downer (RTS), PM Sean Nicolls, Updated 27/09/2023 "REV 2 REVIEW") and six COUNTIFs on `'Specification Requirements'!J:J` only. The project details become app settings/header text, and the counts are replaced by the Dashboard.
- **Reporting Matrix** has 20 content items (A5:A24) × 5 reports (B4:F4, each "MONTHLY ⏎ <name>" or "POST CONSTRUCTION ⏎ <name>") with "X" marks: 26 X marks in total (corrected from 28 in Phase 3). Threaded comments on B4–F4 give each report's source clause. Proposed **Report Components** load: one row per (item, report) X, with Report Name, Report Frequency, Report Source Ref and a link to the obligation that produces the report:

  | Col | Report | Frequency | Source (comment) | Producing obligation (proposed) |
  |---|---|---|---|---|
  | B | Monitoring data made available to TMR on a monthly basis | Monthly | CMP (B) s13.3 | CEMP row 10 |
  | C | PFAS Monthly Compliance Report | Monthly | CMP (B) s13.3 | CEMP row 9 |
  | D | CMP Monthly Audit Report | Monthly | CMP (B) s13.1 | CEMP row 11 (the monthly half after the D-05 split) |
  | E | Compliance Register | Monthly | EMP (D) s6.4 & 8.1; CEMP (B) s8.3 | CEMP row 44 (**⚠ D-13**: 44 or 45?) |
  | F | CMP Close Out Report | Post construction | CMP (B) s15 | CEMP row 8 |

- **_Doc Ref** has 13 documents. Rows 7, 9, 10 and 11 (MD‑12‑164, MD‑15‑315/316/317) are in **red** (not yet provided). Versions are mixed types (4, 1.1, "C", 2.03), so I'll store them as text. The register cites many **documents that aren't in _Doc Ref**, and I propose adding them to Source Documents during import:
  - SARA 2402‑39300 SDA, 2506‑46598 SDA, 2308‑36075 SDA and 2309‑36966 SDA (with their decision notice links)
  - EMP (D), CEMP (B) and its annexes, and CMP (B), which the CEMP sheet cites. _Doc Ref lists CMP Rev **C**, but the CEMP sheet cites CMP (**B**).
  - Fish Salvage Plan, Rehabilitation & Remediation Plan, Fish Passage Monitoring Plan, Aquatic Plant Monitoring Plan and Marine Plant Restoration Plan (WGJV, Sept …)
  - Subcontract for Ormeau Facility Part A, Schedule A5, ISO 14001:2015
  - Each Act on the Legislation sheet (optional, **⚠ D-14**)

---

## 3. Column mapping to the Obligations list

The target columns are the brief's Obligations schema plus two proposed additions, marked ★ (**⚠ D-02**, **⚠ D-03**). In the table, "–" means the sheet has no such column, "*derived*" means the value is computed, and "Std" means the standard 12-column layout used by CHMP, RPP, Specification and Legislation.

| Obligations column | CEMP | PFAS | ADR | Bridge Creek | No Name | Marine Plants (blocks 2–3) | Std |
|---|---|---|---|---|---|---|---|
| Obligation ID | *derived* | *derived* | *derived* | *derived* | *derived* | *derived* | *derived* |
| Source Register | *sheet name* | ← | ← | ← | ← | ← | ← |
| Section/Group | A | C | – | – | – | *permit short name (TNAR/NAC)* | – |
| Condition # | – | – | – | *parsed from D* ("Condition 12" → 12) | C | C | – |
| Project Location/Activity | – | A | A | A | A | A (with "Same project" replaced; blank inherits) | A |
| Project Stage | B → normalised | B | B | B | B | B (blank inherits) | B |
| Requirement | C | D | C | C (citations stripped) | D | D | C |
| Source Reference | D | G | D | D (hidden col) | E | E (leading newline trimmed) | D |
| Compliance Action/Controls | E | H | E | E | F | F | E |
| Practical Implementation | – | F | – | – | – | – | – |
| ★ Evidence Required | L (empty) | K *Compliance Check* | K | K | L | L | L *Supporting Documents* (empty) |
| Stakeholder Review/Approval/Submission | – | – | – | – | – | – | H (empty everywhere) |
| Risk Rating | – | – | – | – | – | – | G |
| Frequency | F → normalised | *default* | *from requirement* | F → normalised | G → normalised | G → normalised | *default* |
| Timing Detail | F (raw) | – | F (empty) | F (raw) | G (raw) | G (raw) | – |
| Due Date | H (empty) | J (empty) | H (empty) | H if a real date | I if a real date | I if a real date | I (empty) |
| Due Rule | – | – | – | H (raw text) | I (raw text) | I (raw text) | – |
| Owner (Person) | *resolved from Responsible Role via Role Assignments* | ← | ← | ← | ← | ← | ← |
| Responsible Role | G → primary role | I | G | G (hidden col) | H | H | F |
| ★ Supporting Roles | G → other roles | I | G | G | H | H | F |
| Status | I → normalised | L | I | I | J | J (K on row 23) | J |
| Compliance Outcome | *derived* (section 4.1) | ← | ← | ← | ← | ← | ← |
| Compliance Notes | J | M | J | J (+ N12 quote) | K | K | K |
| Last Reviewed | K (empty) | – | – | – | – | – | – |
| Info Only | *derived* | ← | ← | ← | ← | ← | ← |
| Has Open Follow-up | *No at import* | ← | ← | ← | ← | ← | ← |
| Source Document (lookup) | *parsed from D* | *from G* (CMP Rev C) | *from D* (ADR 2018) | 2402‑39300 SDA, or plan name from D | 2506‑46598 SDA | 2308‑36075 / 2309‑36966 SDA | *parsed from D* |
| Extra Fields | *see below* | E *Legislation*, comments | comments | citations, comments | comments | stray J (row 23), comments | comments |

**Evidence list (child items):** every hyperlinked cell in the Links/Evidence columns (Bridge Creek L–M, No Name M–N, Marine Plants M–N, and the hidden legacy J6/J7) becomes one **Evidence** item: Title = display text, URL = resolved absolute URL, Added By = the import account, Date = blank. Link cells with **no** hyperlink hold placeholders (`?`, `..`, `N/A – until completion`, `N/A – no current records`, `drawings?`, `monthly reports?`). These are *evidence gaps*, not evidence. They go into Extra Fields as `Links (not hyperlinked): …`, and the transform lists them in an import report (**⚠ D-12**).

**Extra Fields format:** multi-line plain text with one `Label: value` block per item, for example:

```
Source cell range: 'Bridge Creek DA Conditions'!A14:N14
Relevant Legislation / Standard or Code: Environmental Protection Act 1994 - S319 and S440ZG …
Citations: [2402-39300...A Decision | PDF]
Links (not hyperlinked): ?
Excel comment (E13, Allison Brunott): EP Reg apply to this?
Unmapped (J): Ongoing
```

Nothing from any cell is dropped. Every non-empty cell either maps to a column, becomes an Evidence item, or appears in Extra Fields.

### Derived fields

| Field | Rule |
|---|---|
| **Obligation ID** | `<prefix>-<nnn>`, numbered in sheet order within each register (**⚠ D-15** gives the alternatives). Prefixes: CEMP, PFAS, ADR, BC-DA, NN-DA, MP-TNAR, MP-NAC, CHMP, RPP, SPEC, LEG. Sub-conditions keep their parent's number plus a suffix (for example `MP-NAC-006B`). The ID is assigned once at import and never reused. It is the upsert key. |
| **Source Register** | Choice = the sheet name (11 values including the empty *No Name Operational Works*, so it can be used later) |
| **Info Only** | Yes when the action, responsibility or status starts with Note / Noted / Noted – see below. Status becomes *N/A – Info only*. |
| **Condition #** | Text, not a number, because values include `5a`, `6b` and `2c` |

---

## 4. Distinct values and proposed normalisation

The complete lists with counts per sheet are in [inventory section 5](data-audit-inventory.md#5-distinct-values-obligation-info-only-sub-condition-flagged-and-legacy-rows).

### 4.1 Compliance Status → Status + Compliance Outcome

The target Status values are *Not started*, *Underway*, *Complete*, *Missing* and *N/A – Info only*. The target Outcome values are *Compliant*, *Non-compliant* and *Not assessed*. Overdue is **derived**: it is never stored as a status.

| Found value | Where (count) | → Status | → Outcome | Notes |
|---|---|---|---|---|
| *(blank)* | All non-DA sheets (151), BC rows 22–49 (27, including NBSP-only cells) | **Not started** | Not assessed | **⚠ D-01.** Many are plainly under way (weekly inspections in 2026), but I won't guess. |
| Not completed | BC (6) | **⚠ D-01b.** Proposed: *Underway* where Timing = At all times / Duration of works (BC C1, C2, C6); *Not started* for completion-triggered items (BC C3, C8); split for BC C5 (below) | Not assessed | *Not completed* is ambiguous between not started and in progress |
| Not completed – ongoing | NN (4), MP (9) | Underway | Not assessed | Notes such as "DAF inspection completed (2025)" suggest *Compliant*, but I won't infer an outcome |
| Not completed – rehabilitation | NN (1), MP (7) | Not started | Not assessed | Every one is noted as "In Landscaping Scope". **⚠ D-01c:** Not started, or Underway? |
| Ongoing | BC (6), NN (5), MP (3) | Underway | Not assessed | |
| Completed | BC (3), NN (1), MP legacy (1) | Complete | Not assessed | All 4 current rows have an evidence hyperlink, so they meet the "Complete needs evidence" rule |
| Compliant | MP legacy row 7 (1) | *(row not imported, D-08)* | Compliant | |
| (a) – completed | NN row 7 | → Task (a) Done = Yes | – | D-06 |
| noted | MP row 35 | N/A – Info only | Not assessed | |
| Drop-down values never used: *Compliant, Not Compliant, Completed, In Progress, Overdue, Not Started* | Summary A20:A25 | Not Compliant → Outcome *Non-compliant*; In Progress → *Underway*; Completed → *Complete*; Not Started → *Not started*; Compliant → Outcome *Compliant*; Overdue → *derived* | | Kept for the export mapping back to TP002 |

**BC condition 5 / MP TNAR 2 / MP NAC 3 (commencement and completion notifications):** each row combines two notices, one before works and one after completion. The commencement notice is evidenced (hyperlinks to *DAF Notifications* or *Pre-works notification*), but the completion notice is still outstanding ("Required at end of conditions"). I propose splitting each into two **Occurrences** (Commencement notice → Complete, Completion notice → Not started), with the obligation rolling up to Underway (**⚠ D-05**). *Phase 3 note: No Name condition 2 was listed here originally, but it covers the completion notice only, so it isn't split.*

### 4.2 Project Stage

A multi-select Stage would suit the data, but filtering a multi-value choice column **isn't delegable** in SharePoint. I propose a **single-choice** column with combined values instead:

| Found value (count) | → Proposed Stage |
|---|---|
| All (20) | All stages |
| Pre construction (12), Pre-construction (1) | Pre-construction |
| Pre and during construction (23) | Pre-construction and construction |
| During construction (77), Construction (21) | Construction |
| Construction / Operation (8), Design / Construction / Operation (2) | Construction and operation |
| Construction / Rehabilitation (1) | Construction and operation ⚠ – or Rehabilitation and monitoring (BC C3 covers maintenance, monitoring and rehab) |
| Operation / Maintenance (1) | Operation |
| Completion (2), Decommissioning (1) | Completion |
| Pre-construction / Completion (1) | Completion ⚠ (BC C5; see the notification split above) |
| Post construction (3) | Post-construction |
| Rehabilitation (6), Rehabilitation Monitoring (5), Monitoring (4) | Rehabilitation and monitoring |
| Operational Works (13), Temporary Access (2) – Marine Plants | Construction ⚠ – these are the *approval type*, not a stage. Original kept in Extra Fields. |
| *(blank)* – Legislation (16) | All stages |

That gives a choice list of: All stages · Design · Pre-construction · Pre-construction and construction · Construction · Construction and operation · Operation · Completion · Post-construction · Rehabilitation and monitoring. **⚠ D-03** asks you to approve this list; *Design* is unused but kept for future sheets.

### 4.3 Timing/Frequency → Frequency (choice) + Timing Detail (raw)

Frequency choices are: Once / Daily / Weekly / Monthly / Quarterly / Six-monthly / Annual / Trigger-based / Ongoing, plus a proposed **Milestone** (**⚠ D-04**). The raw text is always kept in Timing Detail.

| Found value (sheet, count) | → Frequency | Flag |
|---|---|---|
| Daily (CEMP 2) | Daily | Occurrence volume, section 5 |
| Weekly (CEMP 10) | Weekly | |
| Monthly / "Monthly " (CEMP 8) | Monthly | |
| Monthly – due by the 5th day of each month (CEMP) | Monthly | Due Rule: "5th of following month" |
| Quarterly (CEMP 2) | Quarterly | |
| Quarterly after revegetation commences (CEMP) | Quarterly | Start date = revegetation start (unknown) |
| Quarterly / Six monthly? (CEMP row 41) | **⚠ D-04** | Reviewer's own question |
| Annual (CEMP) | Annual | |
| Weekly ⏎ Monthly (CEMP row 11) | Weekly + Monthly | Split (D-05) |
| Daily ⏎ Weekly (CEMP row 33, ESC inspections) | Weekly ⚠ | Daily detail kept in text (section 5) |
| Weekly ⏎ Ad hoc – following rainfall… (CEMP 2) | Weekly | Rainfall trigger kept in Timing Detail |
| Ad hoc ⏎ Monthly as part of reporting to TMR (CEMP 2) | Monthly | Ad hoc part kept in Timing Detail |
| Monthly during standard storage / operations ⏎ Ad hoc… (CEMP) | Monthly | |
| Ad hoc; Ad hoc – during … / following rainfall … / prior to release / when spoil removed … / one week before disposal … (CEMP 14) | Trigger-based | |
| Weekly check sampling during discharge period ⏎ Ad hoc… (CEMP) | Trigger-based | Weekly only while discharging |
| Daily during the discharge event (CEMP) | Trigger-based | |
| Prior to disposal or release (CEMP) | Trigger-based | |
| Prior to construction works commencing (CEMP 3) | Once | Construction has started, so each is either Complete or **Missing** (D-01) |
| Prior to construction works commencing / At the commencement of vibration-generating activities (CEMP) | Trigger-based | |
| Prior to clearing works commencing – no more than two weeks prior … (CEMP) | Trigger-based | One per clearing event |
| Post construction (CEMP 3) | Once | Due Rule "Post-construction". CEMP row 28 conflicts with its stage. |
| Post construction – within four weeks of construction completion (CEMP) | Once | Due = completion + 4 weeks (D-10) |
| At all times (BC 8, NN 7, MP 6); Duration of works (BC 3); For duration of works (MP 4); Construction; Construction phase (3); Construction and operation; Rehabilitation phase; Ongoing; As specified in plans (BC) | Ongoing | |
| As required (BC 5); When fish stranding/entrapment occurs (NN) | Trigger-based | |
| Daily during dewatering; Daily monitoring (BC 3) | Trigger-based ⚠ | Daily only while dewatering; fish salvage is already Complete. Could be Daily if dewatering is continuing (**⚠ D-04**). |
| During dewatering (2), During pumping activities, During salvage, During planting (BC) | Trigger-based | |
| Prior to dewatering; Following salvage works; Within 2 weeks of rehabilitation completion; Within 10 business days of completion and before post-works notification (BC) | Once | |
| Years 1, 2 and 5 (BC 2); Years 1, 2 and 5 post-construction; Monitoring milestones; Reporting milestones (BC) | **Milestone** ⚠ | 3 Occurrences each (Y1, Y2, Y5), dated from completion (D-10) |
| 5-20 business days before start; within 15 business days of completion (BC) | Once ×2 | Notification split (D-05) |
| By no later than 31 Jul 2026 (NN) | Once | Due Date 31/07/2026 |
| Within 15 business days of completion; Within 10 business days of removal … (NN) | Once | |
| (a) Start… (b) Completion… / (a) At least 5 business days … (b) Within 15 business days … (MP 2) | Once ×2 | Notification split |
| (a)/(b) As soon as reasonably practicable … (MP 4) | Once | |
| (a) As stated in the Marine Plant Restoration Plan (MP 2) | Ongoing ⚠ | Needs the plan's schedule |
| (b)(i) As stated. (b)(ii) Within 10 business days of completion of restoration works and within 5 years … (MP 2) | Milestone ⚠ | |
| *(no timing column)* PFAS, ADR, CHMP, RPP, Specification, Legislation | **Ongoing** by default | Exceptions: PFAS row 11 → Six-monthly (per Compliance Check), PFAS row 12 → Daily; *Pre construction* stage rows → Once; ADR row 6 → Once (pre-works notification). **⚠ D-04** |

### 4.4 Due Date → Due Date + Due Rule

| Found value (count) | → Due Date | → Due Rule (raw text kept) |
|---|---|---|
| 31/07/2026 (NN C10) | 31/07/2026 | "By no later than 31 Jul 2026" |
| 27/10/2023, 28/10/2023 (MP legacy rows 6–7) | not imported (D-08) | – |
| Ongoing (24), Ongoing during works, During works (7), During construction (2) | blank | Ongoing |
| Trigger based (7), Triggered by milestones, As required (2) | blank | Trigger-based |
| Scheduled (5), As per monitoring schedule, As per plan (2), As per approved plan (4), As per restoration schedule (2), Rehabilitation period (2) | blank until D-10 | "Per plan schedule" ⚠ – dates are in the monitoring/restoration plans, which aren't in the workbook |
| Completion (2), Upon Completion, Completion of works, Completion of salvage, Post works (4) | blank until D-10 | "On completion" ⚠ – needs the practical completion date |
| 2 weeks post rehabilitation | blank until D-10 | "Rehabilitation completion + 2 weeks" |
| Prior to works | blank | "Before works start" (already passed) |
| Daily | blank | Daily |
| *(blank)* – every non-DA sheet | blank | blank |

### 4.5 Responsibility → Responsible Role (single choice, filterable) + Supporting Roles (text)

**⚠ D-11.** One cell often names several roles (up to four, separated by line breaks or " / "). Filtering a multi-choice column isn't delegable, so I propose:

- **Responsible Role**: single choice. It is the *first internal DTI role* in the cell, and it drives the Owner through Role Assignments.
- **Supporting Roles**: plain text, holding the rest of the cell. It isn't filtered.

Distinct role tokens after splitting on line breaks and " / ":

| Token (occurrences) | → Role choice | Internal? | Question |
|---|---|---|---|
| Environmental and Sustainability Team (104) | Environmental and Sustainability Team | ✓ | |
| Environmental Manager (45) | Environmental Manager | ✓ | Is this the same person/team as above? The PFAS CMP means the Builder's *Project Environmental Manager*. |
| Environmental Advisor (7) | Environmental Advisor | ✓ | Your role? |
| Project Manager (34) | Project Manager | ✓ | |
| Construction Manager (19) | Construction Manager | ✓ | |
| Site Supervisor (4) | Site Supervisor | ✓ | |
| Site Manager (3) | Site Manager | ✓ | Same as Site Supervisor? |
| Engineering and Design Team (12) | Engineering and Design Team | ✓ | |
| DT Infrastructure (11) | → Project Manager ⚠ | ✓ | It's the company, not a role. OK to map to Project Manager? |
| Contractor (13) | Contractor ⚠ | ✗ | In the DA sheets DTI *is* the contractor. Map to Construction Manager? |
| Asset Owner (1), Maintenance Manager (1) | Asset Owner / Operator ⚠ | ✗ | Operation-phase obligations (BC C11, C12). Who at TMR/QR? |
| All DTI Workers (4) | Supporting only | – | |
| Ecologist, Fauna Spotter-Catcher, Air quality consultant, Environmental consultant, Suitably qualified person(s) (12) | Supporting only (external specialist) | ✗ | The primary role becomes the internal role in the same cell. CEMP row 35 (spotter-catcher only) → Environmental and Sustainability Team. |
| Note / Noted (6) | – (Info Only) | | |
| *(blank)*: CEMP row 60, Spec rows 35 and 37, Legislation (16) | Environmental and Sustainability Team ⚠ | | Default owner for unassigned rows? |

Proposed Responsible Role choices: Environmental and Sustainability Team · Environmental Manager · Environmental Advisor · Project Manager · Construction Manager · Site Manager · Site Supervisor · Engineering and Design Team · Contractor · Asset Owner / Operator.

### 4.6 Risk Rating

| Found | Count | → |
|---|---|---|
| D | 40 (CHMP 8, RPP 3, Spec 29) | D |
| C | 4 (CHMP 1, Spec 3) | C |
| *(blank)* | 181 (every sheet except CHMP, RPP and Spec, plus 5 CHMP/Spec rows) | **blank = "Not rated"** ⚠ D-16. The Follow-up queue sorts by risk, so unrated items need a position (I propose sorting them after D). |

---

## 5. Design implications for Phase 2 (for your awareness)

- **Occurrence volume vs the 2,000 threshold.** With the proposed frequencies there are about 3 Daily, 14 Weekly, 13 Monthly, 3 Quarterly, 1 Six-monthly and 1 Annual recurring obligations. That is roughly 1,100 daily + 730 weekly + 160 monthly, or about 2,000 Occurrences a year. I'll propose rolling **Daily** obligations up into one *weekly* Occurrence ("daily checks completed for week ending …"), which brings the total to about 1,050 a year. Either way the Occurrences list will pass 2,000 items within the project, so every filter on it will use indexed Obligation ID, Period Due, Status and Assignee columns.
- **Lookup count.** Obligations has one lookup (Source Document), and each child list has one (Obligation). All are well under the 12-per-view limit.
- **"My tasks" depends on Owner being a Person column.** Role-only rows resolve through Role Assignments at import and again when assignments change (a flow).

---

## 6. Rows I can't confidently parse

| # | Sheet!Row | Content | Why | Proposed |
|---|---|---|---|---|
| 1 | RPP Exemption!6 | C = `?)`, D = RPP exemption s…, E = Work Instruction…, Risk D | Requirement text is missing or truncated | Import as an obligation with Requirement "⚠ Requirement text missing in TP002 – see source" and raise a Follow-up |
| 2 | Legislation!21 | D = *EPBC Act – significant impact identified during works* only | No requirement | Import with Requirement = D text and raise a Follow-up to confirm scope |
| 3 | Marine Plants!6–7 (hidden) | Legacy approval summaries; row 7 permit number conflicts (36996 vs 36966) and site (TNAR vs NAC) | Superseded layout with conflicting data | Source Document rows only (D-08) |
| 4 | Marine Plants!8–9 (hidden) | "NAC" only | Stub | Drop |
| 5 | Marine Plants!23 | Shifted one column from J onward | Misaligned | Remap as in section 2.3 (D-08) |
| 6 | No Name!7 | `1:a-d`, list of 4 works, "(a) – completed" | Sub-list inside a condition | 4 Tasks (D-06) |
| 7 | ADR!6 col A | Editorial note in Location | Not location data | Extra Fields and a Follow-up |
| 8 | CEMP!11 | Two deliverables in one row | Mixed frequencies | Split (D-05) |
| 9 | CEMP!41 | "Quarterly / Six monthly?" | Unresolved | D-04 |
| 10 | CEMP!28 | Stage Pre construction vs timing Post construction | Contradictory | Keep both. Flag in Follow-up. |
| 11 | Spec!35, 37 | "Noted – relevant to rehabilitation / landscaping plan" | Info-only or obligation? | D-07 |
| 12 | No Name!10, Marine!26, 41 (col L) | DAF letter quote in the Evidence column | Content in the wrong column | Notes |

---

## 7. Decisions (all accepted as recommended, 29/09/2026)

| ID | Decision | My recommendation |
|---|---|---|
| **D-01** | Starting status for the 151 rows with no status (and BC rows 22–49) | *Not started / Not assessed*, Last Reviewed blank. The Dashboard shows "never reviewed" so the team can work through them. Alternatively, give me a bulk rule (for example "During construction + Weekly/Monthly = Underway"). |
| D-01b | BC "Not completed" rows | Underway for C1, C2 and C6; Not started for C3 and C8; C5 split (D-05) |
| D-01c | "Not completed – rehabilitation" | Not started |
| **D-02** | Add an **Evidence Required** column (multi-line) for the "Evidence" description and PFAS *Compliance Check* | Yes |
| **D-03** | Project Stage as a single choice with the 10 values in 4.2, plus **Supporting Roles** (text) | Yes |
| **D-04** | Add a **Milestone** frequency. Answer CEMP row 41 (Quarterly or Six-monthly?), the BC dewatering daily monitoring (still active?), and the default *Ongoing* for sheets without a timing column | Add Milestone. Default Ongoing. |
| **D-05** | Split combined rows: CEMP row 11 into two obligations, and the four commencement/completion notification conditions into two Occurrences each | Yes |
| **D-06** | No Name condition 1 (a)–(d) becomes 4 Tasks, and other sub-conditions (5b, 6b, 2a–c, 7b) become their own obligations with parent-derived IDs | Yes |
| **D-07** | Specification rows 35 and 37: info-only? | Info-only |
| **D-08** | Marine Plants legacy rows 6–9 become Source Documents only. Confirm which permit number is right (2309‑36966 or 36996). Confirm the row 23 remap. Confirm Bridge Creek has only 14 conditions. | As proposed. 36966 matches the decision notice filename. |
| **D-09** | Strip Bridge Creek AI-citation fragments, keeping them in Extra Fields | Yes |
| **D-10** | Key project dates for due-date rules: construction start, practical completion (construction completion), rehabilitation completion, and any dewatering periods | Needed to date the Milestone and "On completion" items. Until then they stay undated with a Due Rule. |
| **D-11** | Role mapping in 4.5: DT Infrastructure → Project Manager? Contractor → Construction Manager? Env Manager vs Env & Sustainability Team? Site Manager vs Site Supervisor? Who is Asset Owner / Operator? Default owner for blank rows? Then give me **names** per role for Role Assignments. | As listed |
| **D-12** | At import, auto-create Follow-ups for: the 2 unparseable rows, the ADR editorial note, evidence-gap placeholders (30 cells: 9 are "?", "..", "drawings?" or "monthly reports?", and 21 are "N/A – until completion / no current records"), and the 36 Excel review comments | Follow-ups for the unparseable rows and the ADR note only. Comments and placeholders go to Extra Fields and an import report you can triage. |
| **D-13** | Reporting Matrix column E (Compliance Register): produced by CEMP row 44 or row 45? | Row 44 |
| **D-14** | Add every cited plan, permit and Act to Source Documents, not just the 13 in _Doc Ref | Yes: plans and permits. Acts optional. |
| **D-15** | Obligation ID scheme: (a) sequential per register (`BC-DA-012` = 12th Bridge Creek row), or (b) condition-based where a condition exists (`BC-DA-C12`), sequential otherwise | (a). It's stable, uniform and short, and the condition number has its own column. |
| **D-16** | Unrated risk: leave blank ("Not rated") or default to D | Blank. Sort after D. |

---

## 8. What to check before approving

1. Skim [`data-audit-inventory.md`](data-audit-inventory.md) sections 2–3 and confirm the row classifications, especially Marine Plants and No Name.
2. Answer or accept each decision in section 7. A reply such as "accept all except D-11: …" is enough.
3. Send me the role → person names (D-11) and the project dates (D-10) if you have them. They aren't needed for Phase 2 (the lists), only for Phase 3 (the data load).
4. The source workbook has been committed to `source/` so the transform is reproducible. It holds internal SharePoint links, so keep the repo private. If you'd rather it stayed out of git, tell me and I'll move it to `.gitignore` and read it from a path you supply.

There are no browser steps in Phase 1.
