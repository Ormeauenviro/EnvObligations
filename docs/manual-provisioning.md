# Manual provisioning (fallback)

Use this guide if the Entra app registration for PnP.PowerShell isn't approved. It builds exactly what `provision/Provision-Lists.ps1` builds, through the SharePoint browser UI. Allow about 3–4 hours for all 10 lists. Obligations (40 columns) is the longest.

**You need:** Site Owner permission on the target site.

**Tip:** do the lists in the order given. Source Documents and Obligations must exist before the lists that look them up.

---

## Step 0 – Site regional settings (once)

1. Open the site, then choose ⚙ **Settings › Site information › View all site settings**.
2. Under *Site Administration*, open **Regional settings**:
   - **Time zone:** (UTC+10:00) Brisbane
   - **Locale:** English (Australia), which gives DD/MM/YYYY dates
   - **First day of week:** Monday
   - **Time format:** 12 Hour
3. Click **OK**.

## Step 1 – Understand the "internal name" rule (important)

SharePoint fixes a column's **internal name** from the name you type when you *first* create it. It can never be changed afterwards. The app and flows refer to the internal names in this guide, so you must:

1. Create each column with its **internal name** exactly as shown (for example `SourceRegister`, with no spaces).
2. Save it, then click it again and change *Column name* to the **display name** (for example *Source Register*).

If you type the display name first, the internal name becomes `Source_x0020_Register` and the app formulas won't match. If that happens, delete the column and recreate it.

## Step 2 – Where the settings are

For each list, open it and choose ⚙ **Settings › List settings** (the classic settings page). Every setting in this guide is on that page:

| Task | Where |
|---|---|
| Rename the list, set its description | *List name, description and navigation* |
| Versioning | *Versioning settings* |
| Attachments | *Advanced settings* |
| Add a column | *Columns › Create column* |
| Rename, require or make unique an existing column | *Columns › (column name)* |
| Indexes | *Columns › Indexed columns › Create a new index* |
| Views | *Views › (view name)*, or *Create view › Standard view* |

## Step 3 – Column settings by type

When you create a column, pick the type and settings shown in the checklist. For every column:

- Set *Require that this column contains information* to **Yes** only when the checklist says Required = Yes.
- Set *Enforce unique values* to **Yes** only when Unique = Yes. SharePoint will offer to index the column; accept.
- For choice columns, type the choices **one per line, exactly as shown**, including capitals and the en dash in *N/A – Info only*. Copy and paste from this page.
- For Default: on choice columns, set *Default value* to the value shown. On Yes/No columns, set *Default value* to Yes or No as shown.
- Clear *Add to all content types* and *Add to default view*. The views are set up in step 4.

## Step 4 – Views

For each view in the checklist:

1. Open the view, or create it with *Create view › Standard view*, give it the view name, and choose *Public view*.
2. **Columns:** tick only the columns listed, and set *Position from Left* in the listed order.
3. **Sort, Filter and Group By:** as listed. *[Today]* and *[Me]* are typed literally into the filter value box. For *[Today]+30*, type `[Today]+30`.
4. **Item Limit:** 100, with *Display items in batches of the specified size*.
5. Tick *Make this the default view* where the checklist says so.

## Step 5 – Optional: protect Status History

To make Status History effectively append-only for normal users, open **Status History › List settings › Advanced settings** and set *Create and Edit access* to **Create items and edit items that were created by the user**. Versioning still records any edit.

## Changing a column type

SharePoint can't change a column's type in place, and `Provision-Lists.ps1` warns rather than forcing it. If a column was created with the wrong type:

- **If the list has no data yet:** delete the column and recreate it with the internal name.
- **If it has data:** create a temporary column, copy the values (Power Automate or grid view), delete the old column, recreate it with the correct type *and the same internal name*, then copy the values back.

## Verification (either method)

- [ ] 10 lists on the site's navigation, with the URLs in the checklist
- [ ] Each list: *Versioning settings* shows 500 versions. Attachments are on for **Evidence** only.
- [ ] Obligations *Indexed columns* page shows 16 indexes. Occurrences shows 9.
- [ ] New Obligations item form: Status defaults to *Not started*, Compliance Outcome to *Not assessed*, Active to *Yes*
- [ ] Creating two Obligations with the same Obligation ID fails with "value must be unique"
- [ ] A date column shows DD/MM/YYYY

---

## Checklists (generated from `provision/list-schema.json`)

<!-- BEGIN GENERATED: checklists -->

### 1. Source Documents

- [ ] **Create the list:** *New › List › Blank list*. Name it `SourceDocuments` (no spaces, so the URL is `Lists/SourceDocuments`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Source Documents** and paste the description: *Source document register (from TP002 _Doc Ref plus permits and plans cited in the register).*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Document Title**, and leave *Require* on **Yes**.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `DocKey` | Document Key | Single line of text; Maximum number of characters: 255 | Yes | Yes | – |
  | 2 | `Category` | Category | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Specification`<br>`Heritage`<br>`WWBW`<br>`RPP Exemption`<br>`CMP`<br>`Development approval`<br>`Management plan`<br>`Contract`<br>`Standard`<br>`Legislation` | No | No | – |
  | 3 | `Reference` | # / Reference | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `Version` | Version | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 5 | `DocDate` | Document Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 6 | `DocLink` | Document Link | Hyperlink or Picture; Format URL as: Hyperlink | No | No | – |
  | 7 | `FolderLink` | Folder Link | Hyperlink or Picture; Format URL as: Hyperlink | No | No | – |
  | 8 | `Provided` | Provided | Yes/No (check box) | No | No | Yes |
  | 9 | `Comments` | Comments | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Document Key (`DocKey`), Category (`Category`), Provided (`Provided`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Document Key, Document Title, Category, # / Reference, Version, Document Date, Provided, Document Link.
    - Sort: **Category** ascending, then **Document Key** ascending
  - **Not Provided**. Columns: Document Key, Document Title, Category, # / Reference, Comments.
    - Filter: Show items only when **Provided** is equal to `No`

### 2. Project Milestones

- [ ] **Create the list:** *New › List › Blank list*. Name it `ProjectMilestones` (no spaces, so the URL is `Lists/ProjectMilestones`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Project Milestones** and paste the description: *Key project dates that anchor due-date rules (D-10).*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Milestone**, and leave *Require* on **Yes**. Set *Enforce unique values* to **Yes** and accept the prompt to index it.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `PlannedDate` | Planned Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 2 | `ActualDate` | Actual Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 3 | `Notes` | Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Title (Milestone).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Milestone, Planned Date, Actual Date, Notes.
    - Sort: **Planned Date** ascending

### 3. Role Assignments

- [ ] **Create the list:** *New › List › Blank list*. Name it `RoleAssignments` (no spaces, so the URL is `Lists/RoleAssignments`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Role Assignments** and paste the description: *Maps each Responsible Role to real people. Obligation Owner is resolved from Primary Person.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Responsible Role**, and leave *Require* on **Yes**. Set *Enforce unique values* to **Yes** and accept the prompt to index it.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `PrimaryPerson` | Primary Person | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 2 | `AdditionalPeople` | Additional People | Person or Group; Allow multiple selections: Yes; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 3 | `Notes` | Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Title (Responsible Role), Primary Person (`PrimaryPerson`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Responsible Role, Primary Person, Additional People, Notes.
    - Sort: **Responsible Role** ascending

### 4. Obligations

- [ ] **Create the list:** *New › List › Blank list*. Name it `Obligations` (no spaces, so the URL is `Lists/Obligations`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Obligations** and paste the description: *Environmental compliance obligations (replaces the TP002.1 register sheets).*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Obligation ID**, and leave *Require* on **Yes**. Set *Enforce unique values* to **Yes** and accept the prompt to index it.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `SourceRegister` | Source Register | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`CEMP and General Requirements`<br>`PFAS CMP Requirements`<br>`ADR WWBW Conditions`<br>`Bridge Creek DA Conditions`<br>`No Name DA Conditions`<br>`Marine Plants DA Conditions`<br>`No Name Operational Works`<br>`External CHMP Conditions`<br>`RPP Exemption`<br>`Specification Requirements`<br>`Legislation` | Yes | No | – |
  | 2 | `SourceOrder` | Source Order | Number; Number of decimal places: 0 | No | No | – |
  | 3 | `SourceCell` | Source Cell | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `SectionGroup` | Section/Group | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 5 | `ConditionNo` | Condition # | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 6 | `ParentObligationID` | Parent Obligation ID | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 7 | `LocationActivity` | Project Location/Activity | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 8 | `ProjectStage` | Project Stage | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`All stages`<br>`Design`<br>`Pre-construction`<br>`Pre-construction and construction`<br>`Construction`<br>`Construction and operation`<br>`Operation`<br>`Completion`<br>`Post-construction`<br>`Rehabilitation and monitoring` | No | No | – |
  | 9 | `RequirementSummary` | Requirement Summary | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 10 | `Requirement` | Requirement | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | Yes | No | – |
  | 11 | `SourceReference` | Source Reference | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 12 | `SourceDocument` | Source Document | Lookup (information already on this site); Get information from: **Source Documents**; In this column: **Title**; Enforce relationship behaviour: **Yes › Restrict delete** (accept the prompt to index) | No | No | – |
  | 13 | `ComplianceAction` | Compliance Action/Controls | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 14 | `PracticalImplementation` | Practical Implementation | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 15 | `EvidenceRequired` | Evidence Required | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 16 | `Stakeholder` | Stakeholder Review/Approval/Submission | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 17 | `RiskRating` | Risk Rating | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`A`<br>`B`<br>`C`<br>`D` | No | No | – |
  | 18 | `RiskSort` | Risk Sort | Number; Number of decimal places: 0 | No | No | – |
  | 19 | `Frequency` | Frequency | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Once`<br>`Daily`<br>`Weekly`<br>`Monthly`<br>`Quarterly`<br>`Six-monthly`<br>`Annual`<br>`Milestone`<br>`Trigger-based`<br>`Ongoing` | Yes | No | – |
  | 20 | `TimingDetail` | Timing Detail | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 21 | `RecurrenceStart` | Recurrence Start | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 22 | `RecurrenceEnd` | Recurrence End | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 23 | `DueAnchor` | Due Anchor | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Construction start`<br>`Practical completion`<br>`Rehabilitation completion`<br>`Revegetation commencement`<br>`Temporary barrier removal`<br>`Restoration works completion` | No | No | – |
  | 24 | `DueOffsetDays` | Due Offset (days) | Number; Number of decimal places: 0 | No | No | – |
  | 25 | `DueDate` | Due Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 26 | `DueRule` | Due Rule | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 27 | `Owner` | Owner | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 28 | `ResponsibleRole` | Responsible Role | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Environmental and Sustainability Team`<br>`Environmental Manager`<br>`Environmental Advisor`<br>`Project Manager`<br>`Construction Manager`<br>`Site Manager`<br>`Site Supervisor`<br>`Engineering and Design Team`<br>`Contractor`<br>`Asset Owner / Operator` | No | No | – |
  | 29 | `SupportingRoles` | Supporting Roles | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 30 | `Status` | Status | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Not started`<br>`Underway`<br>`Complete`<br>`Missing`<br>`N/A – Info only` | Yes | No | Not started |
  | 31 | `StatusComment` | Status Reason/Justification | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 32 | `ComplianceOutcome` | Compliance Outcome | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Not assessed`<br>`Compliant`<br>`Non-compliant` | Yes | No | Not assessed |
  | 33 | `ComplianceNotes` | Compliance Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 34 | `LastReviewed` | Last Reviewed | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 35 | `InfoOnly` | Info Only | Yes/No (check box) | No | No | No |
  | 36 | `Active` | Active | Yes/No (check box) | No | No | Yes |
  | 37 | `HasOpenFollowUp` | Has Open Follow-up | Yes/No (check box) | No | No | No |
  | 38 | `IsOverdue` | Is Overdue | Yes/No (check box) | No | No | No |
  | 39 | `CurrentOccurrenceKey` | Current Occurrence Key | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 40 | `ExtraFields` | Extra Fields | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Title (Obligation ID), Source Register (`SourceRegister`), Project Stage (`ProjectStage`), Requirement Summary (`RequirementSummary`), Source Document (`SourceDocument`), Risk Sort (`RiskSort`), Frequency (`Frequency`), Due Date (`DueDate`), Owner (`Owner`), Responsible Role (`ResponsibleRole`), Status (`Status`), Compliance Outcome (`ComplianceOutcome`), Info Only (`InfoOnly`), Active (`Active`), Has Open Follow-up (`HasOpenFollowUp`), Is Overdue (`IsOverdue`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Obligation ID, Source Register, Condition #, Requirement Summary, Project Stage, Frequency, Due Date, Owner, Responsible Role, Status, Compliance Outcome, Risk Rating, Has Open Follow-up, Is Overdue.
    - Sort: **Source Register** ascending, then **Source Order** ascending
  - **Open**. Columns: Obligation ID, Source Register, Requirement Summary, Frequency, Due Date, Owner, Status, Risk Rating, Is Overdue.
    - Filter: Show items only when **Status** is equal to `Not started` OR **Status** is equal to `Underway` OR **Status** is equal to `Missing`
    - Sort: **Due Date** ascending
  - **Overdue**. Columns: Obligation ID, Source Register, Requirement Summary, Due Date, Owner, Status, Risk Rating.
    - Filter: Show items only when **Is Overdue** is equal to `Yes`
    - Sort: **Due Date** ascending
  - **Missing or Follow-up**. Columns: Obligation ID, Source Register, Requirement Summary, Owner, Status, Status Reason/Justification, Has Open Follow-up.
    - Filter: Show items only when **Has Open Follow-up** is equal to `Yes` OR **Status** is equal to `Missing`
    - Sort: **Risk Sort** ascending
  - **By Register**. Columns: Obligation ID, Section/Group, Condition #, Requirement Summary, Frequency, Owner, Status, Compliance Outcome.
    - Sort: **Source Order** ascending
    - Group By: **Source Register** (collapsed)
  - **My Obligations**. Columns: Obligation ID, Source Register, Requirement Summary, Frequency, Due Date, Status, Is Overdue.
    - Filter: Show items only when **Owner** is equal to `[Me]`
    - Sort: **Due Date** ascending
  - **Info Only**. Columns: Obligation ID, Source Register, Requirement Summary, Source Reference.
    - Filter: Show items only when **Info Only** is equal to `Yes`
    - Sort: **Source Order** ascending

### 5. Occurrences

- [ ] **Create the list:** *New › List › Blank list*. Name it `Occurrences` (no spaces, so the URL is `Lists/Occurrences`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Occurrences** and paste the description: *One item per period for recurring and milestone obligations. Expected to exceed 2,000 items.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Occurrence Key**, and leave *Require* on **Yes**. Set *Enforce unique values* to **Yes** and accept the prompt to index it.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `Obligation` | Obligation | Lookup (information already on this site); Get information from: **Obligations**; In this column: **Title**; Enforce relationship behaviour: **Yes › Restrict delete** (accept the prompt to index) | Yes | No | – |
  | 2 | `ObligationID` | Obligation ID | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 3 | `ObligationSummary` | Obligation Summary | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `PeriodLabel` | Period Label | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 5 | `PeriodStart` | Period Start | Date and Time; Date and Time Format: Date Only; Display Format: Standard | Yes | No | – |
  | 6 | `PeriodEnd` | Period End | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 7 | `PeriodDue` | Period Due | Date and Time; Date and Time Format: Date Only; Display Format: Standard | Yes | No | – |
  | 8 | `Assignee` | Assignee | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 9 | `Status` | Status | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Not started`<br>`Underway`<br>`Complete`<br>`Missing` | Yes | No | Not started |
  | 10 | `StatusComment` | Status Reason/Justification | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 11 | `ComplianceOutcome` | Compliance Outcome | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Not assessed`<br>`Compliant`<br>`Non-compliant` | Yes | No | Not assessed |
  | 12 | `CompletedDate` | Completed Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 13 | `EvidenceURL` | Evidence URL | Hyperlink or Picture; Format URL as: Hyperlink | No | No | – |
  | 14 | `Notes` | Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 15 | `RiskSort` | Risk Sort | Number; Number of decimal places: 0 | No | No | – |
  | 16 | `IsOverdue` | Is Overdue | Yes/No (check box) | No | No | No |
  | 17 | `HasOpenFollowUp` | Has Open Follow-up | Yes/No (check box) | No | No | No |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Title (Occurrence Key), Obligation (`Obligation`), Obligation ID (`ObligationID`), Period Start (`PeriodStart`), Period Due (`PeriodDue`), Assignee (`Assignee`), Status (`Status`), Risk Sort (`RiskSort`), Is Overdue (`IsOverdue`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Occurrence Key, Obligation ID, Period Label, Period Due, Assignee, Status, Completed Date, Evidence URL, Is Overdue.
    - Sort: **Period Due** descending
  - **Open**. Columns: Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Assignee, Status, Is Overdue.
    - Filter: Show items only when **Status** is equal to `Not started` OR **Status** is equal to `Underway` OR **Status** is equal to `Missing`
    - Sort: **Period Due** ascending
  - **Due Next 30 Days**. Columns: Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Assignee, Status.
    - Filter: Show items only when **Period Due** is greater than or equal to `[Today]` AND **Period Due** is less than or equal to `[Today]+30`
    - Sort: **Period Due** ascending
  - **Overdue**. Columns: Occurrence Key, Obligation ID, Obligation Summary, Period Due, Assignee, Status.
    - Filter: Show items only when **Is Overdue** is equal to `Yes`
    - Sort: **Period Due** ascending
  - **My Occurrences**. Columns: Occurrence Key, Obligation ID, Obligation Summary, Period Label, Period Due, Status.
    - Filter: Show items only when **Assignee** is equal to `[Me]`
    - Sort: **Period Due** ascending

### 6. Obligation Tasks

- [ ] **Create the list:** *New › List › Blank list*. Name it `ObligationTasks` (no spaces, so the URL is `Lists/ObligationTasks`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Obligation Tasks** and paste the description: *Checklist items that can be ticked off against an obligation.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Task**, and leave *Require* on **Yes**.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `Obligation` | Obligation | Lookup (information already on this site); Get information from: **Obligations**; In this column: **Title**; Enforce relationship behaviour: **Yes › Restrict delete** (accept the prompt to index) | Yes | No | – |
  | 2 | `ObligationID` | Obligation ID | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 3 | `ObligationSummary` | Obligation Summary | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `TaskOrder` | Task Order | Number; Number of decimal places: 0 | No | No | – |
  | 5 | `Assignee` | Assignee | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 6 | `DueDate` | Due Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 7 | `Done` | Done | Yes/No (check box) | No | No | No |
  | 8 | `DoneDate` | Done Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 9 | `Notes` | Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Obligation (`Obligation`), Obligation ID (`ObligationID`), Assignee (`Assignee`), Due Date (`DueDate`), Done (`Done`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Obligation ID, Task Order, Task, Assignee, Due Date, Done, Done Date.
    - Sort: **Obligation ID** ascending, then **Task Order** ascending
  - **Open Tasks**. Columns: Obligation ID, Task, Assignee, Due Date.
    - Filter: Show items only when **Done** is equal to `No`
    - Sort: **Due Date** ascending
  - **My Tasks**. Columns: Obligation ID, Task, Due Date, Done.
    - Filter: Show items only when **Assignee** is equal to `[Me]` AND **Done** is equal to `No`
    - Sort: **Due Date** ascending

### 7. Evidence

- [ ] **Create the list:** *New › List › Blank list*. Name it `Evidence` (no spaces, so the URL is `Lists/Evidence`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Evidence** and paste the description: *Links to evidence held in SharePoint, against an obligation and optionally an occurrence.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Enabled**.
- [ ] **Title column:** click *Title*, rename it to **Title**, and leave *Require* on **Yes**.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `ObligationID` | Obligation ID | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 2 | `OccurrenceKey` | Occurrence Key | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 3 | `EvidenceURL` | SharePoint URL | Hyperlink or Picture; Format URL as: Hyperlink | Yes | No | – |
  | 4 | `EvidenceDate` | Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 5 | `AddedBy` | Added By | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 6 | `Notes` | Notes | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Obligation ID (`ObligationID`), Occurrence Key (`OccurrenceKey`), Date (`EvidenceDate`), Added By (`AddedBy`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Obligation ID, Occurrence Key, Title, SharePoint URL, Date, Added By.
    - Sort: **ID** descending

### 8. Follow-ups

- [ ] **Create the list:** *New › List › Blank list*. Name it `FollowUps` (no spaces, so the URL is `Lists/FollowUps`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Follow-ups** and paste the description: *Items needing follow-up: Missing obligations, overdue items, evidence gaps and import review.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Follow-up**, and leave *Require* on **Yes**.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `ObligationID` | Obligation ID | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 2 | `OccurrenceKey` | Occurrence Key | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 3 | `ObligationSummary` | Obligation Summary | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `ReasonType` | Reason Type | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Missing`<br>`Overdue`<br>`Evidence gap`<br>`Import review`<br>`Other` | Yes | No | Missing |
  | 5 | `Reason` | Reason | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | Yes | No | – |
  | 6 | `Owner` | Owner | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 7 | `FollowUpDate` | Follow-up Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 8 | `Status` | Status | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Open`<br>`Resolved` | Yes | No | Open |
  | 9 | `ResolutionComment` | Resolution Comment | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 10 | `ResolvedBy` | Resolved By | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 11 | `ResolvedDate` | Resolved Date | Date and Time; Date and Time Format: Date Only; Display Format: Standard | No | No | – |
  | 12 | `RiskSort` | Risk Sort | Number; Number of decimal places: 0 | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Obligation ID (`ObligationID`), Occurrence Key (`OccurrenceKey`), Reason Type (`ReasonType`), Owner (`Owner`), Follow-up Date (`FollowUpDate`), Status (`Status`), Risk Sort (`RiskSort`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Follow-up, Obligation ID, Reason Type, Owner, Follow-up Date, Status, Created.
    - Sort: **Created** descending
  - **Open**. Columns: Follow-up, Obligation ID, Obligation Summary, Reason Type, Reason, Owner, Follow-up Date, Created.
    - Filter: Show items only when **Status** is equal to `Open`
    - Sort: **Risk Sort** ascending, then **Created** ascending
  - **My Open Follow-ups**. Columns: Follow-up, Obligation ID, Reason Type, Reason, Follow-up Date.
    - Filter: Show items only when **Owner** is equal to `[Me]` AND **Status** is equal to `Open`
    - Sort: **Follow-up Date** ascending

### 9. Status History

- [ ] **Create the list:** *New › List › Blank list*. Name it `StatusHistory` (no spaces, so the URL is `Lists/StatusHistory`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Status History** and paste the description: *Append-only log of every status change (app, flows and import).*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Item ID**, and leave *Require* on **Yes**. Index it (see the indexes step).
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `ItemType` | Item Type | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Obligation`<br>`Occurrence`<br>`Task`<br>`Follow-up` | Yes | No | – |
  | 2 | `ObligationID` | Obligation ID | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 3 | `OldStatus` | Old Status | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `NewStatus` | New Status | Single line of text; Maximum number of characters: 255 | Yes | No | – |
  | 5 | `ChangedBy` | Changed By | Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users | No | No | – |
  | 6 | `ChangedAt` | Changed At | Date and Time; Date and Time Format: Date & Time; Display Format: Standard | Yes | No | – |
  | 7 | `Comment` | Comment | Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No | No | No | – |
  | 8 | `ChangeSource` | Change Source | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`App`<br>`Flow`<br>`Import` | No | No | App |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Title (Item ID), Item Type (`ItemType`), Obligation ID (`ObligationID`), Changed At (`ChangedAt`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Changed At, Item ID, Item Type, Obligation ID, Old Status, New Status, Changed By, Change Source, Comment.
    - Sort: **Changed At** descending
  - **Last 30 Days**. Columns: Changed At, Item ID, Item Type, Old Status, New Status, Changed By, Comment.
    - Filter: Show items only when **Changed At** is greater than or equal to `[Today]-30`
    - Sort: **Changed At** descending

### 10. Report Components

- [ ] **Create the list:** *New › List › Blank list*. Name it `ReportComponents` (no spaces, so the URL is `Lists/ReportComponents`) and click Create. Then, in *List settings › List name, description and navigation*, rename it to **Report Components** and paste the description: *Reporting Matrix: which content items go into which report.*
- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** and *Keep the following number of versions* to **500**.
- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to **Disabled**.
- [ ] **Title column:** click *Title*, rename it to **Content Item**, and leave *Require* on **Yes**.
- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column and change *Column name* to the **display name** (details in step 3):

  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |
  |---|---|---|---|---|---|---|
  | 1 | `ReportName` | Report Name | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Monitoring data made available to TMR`<br>`PFAS Monthly Compliance Report`<br>`CMP Monthly Audit Report`<br>`Compliance Register`<br>`CMP Close Out Report` | Yes | No | – |
  | 2 | `ReportFrequency` | Report Frequency | Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No; Choices (one per line):<br>`Monthly`<br>`Post construction` | No | No | – |
  | 3 | `ReportSourceRef` | Report Source Reference | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 4 | `ProducingObligationID` | Producing Obligation ID | Single line of text; Maximum number of characters: 255 | No | No | – |
  | 5 | `SortOrder` | Sort Order | Number; Number of decimal places: 0 | No | No | – |

- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column (unique and lookup columns may already be indexed): Report Name (`ReportName`), Producing Obligation ID (`ProducingObligationID`).
- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, *Display items in batches*.

  - **All Items** · **Make this the default view**. Columns: Content Item, Report Frequency, Producing Obligation ID, Report Source Reference.
    - Sort: **Sort Order** ascending
    - Group By: **Report Name** (expanded)

<!-- END GENERATED: checklists -->
