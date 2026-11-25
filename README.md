# TransAxle Learning & Development Dashboard

A single-file web dashboard (`index_v.1.0.html`) that reads an uploaded team learning tracker Excel workbook and shows learning hours, target achievement, readiness programmes, person-to-person comparisons and a yearly training timeline. It is designed to be hosted as a static page (for example **GitHub Pages**) with **no server**.

- **Single file:** all HTML, CSS and JavaScript are inside `index_v.1.0.html`. No libraries, no server, no database.
- **Upload, don't link:** the user uploads the workbook; no file name or path is hard-coded.
- **Kept for the tab only:** the uploaded workbook survives page reloads in the same tab and is **deleted automatically when the tab is closed** (`sessionStorage`).
- **Private:** the workbook is read inside the browser and never sent anywhere.
- **No hard-coded data:** every name, year, method, training and number shown comes from the workbook.
- **Read-only:** the dashboard never modifies the workbook.

---

## 1. Purpose

| Goal | How the dashboard supports it |
|---|---|
| Track learning hours against the yearly target (default **40.00 h**) | KPI cards, target-achievement donut, per-employee bar chart with a dashed target line |
| Monitor readiness programmes (BTV Batch 1/2, PLATON Basic/Advanced) | Readiness KPI cards and chart with participation % and hours |
| Compare up to 4 people | Compare section with cascading filters, charts and person cards |
| See which trainings were conducted each year | Timeline grouped by year with Technical / Behavioral filters and search |
| Verify numbers against Excel | Participants panel lists names, hours, the employee's Excel column and the workbook row numbers |
| Find data-entry problems | Data quality panel lists missing names, invalid hours, blank years, etc. |

---

## 2. Files

| File | Purpose |
|---|---|
| `index_v.1.0.html` | The dashboard. For GitHub Pages, publish it as `index.html` (or link to it). |
| Any `.xlsx` learning tracker in the required format | Uploaded by the user. **File name and location do not matter.** Not stored in the repository. |
| `DT-LD-Dashbaord_README.md` | This document |

---

## 3. Quick start

### Hosting on GitHub Pages

1. Commit `index_v.1.0.html` to the repository (rename to `index.html` if it should be the site's home page).
2. Enable **Settings → Pages** for the branch.
3. Open the site URL. **Do not commit the workbook** — each user uploads it in the browser.

### Using the dashboard

1. Open the page. A welcome screen appears.
2. Click **Upload workbook…** (or drag and drop the `.xlsx` anywhere on the page).
3. The dashboard loads. Reloading the page in the **same tab** keeps it; **closing the tab deletes it** automatically.
4. To update: edit the workbook in Excel, save, then **Upload workbook…** again (or drag it in). In Edge/Chrome, **Refresh data** can re-read the same file directly after asking for permission.

### Where the data comes from

```mermaid
flowchart TD
    A([Page opens / Refresh data]) --> B{File uploaded or dropped<br/>in this tab?}
    B -- Yes --> R{Can the browser<br/>re-read it?}
    R -- Yes --> L[Read the file]
    R -- No --> C
    B -- No --> C{Copy kept<br/>for this tab?}
    C -- Yes --> S[Use the tab copy]
    C -- No --> W[Welcome screen:<br/>Upload workbook / drag & drop]
    L --> V{Required sheets and<br/>headers present?}
    S --> V
    V -- Yes --> D[Dashboard]
    L -- valid --> K[(Keep copy for this tab<br/>sessionStorage)]
    K -. tab closed .-> X[Deleted by the browser]
    V -- No --> E[Error naming what is missing] --> W
```

---

## 4. Configuration

All settings are in the `CONFIG` block at the top of the `<script>` in `index_v.1.0.html`. There is **no file path or file-name setting**; the workbook is always uploaded.

### 4.1 Tab storage

| Item | Detail |
|---|---|
| Where | The browser's `sessionStorage` for this site, in **this tab only** |
| What | The last **valid** uploaded workbook (name, modified time and file contents) |
| Lifetime | Survives reloads and **Refresh data** in the same tab; **removed automatically when the tab or browser is closed** |
| Sharing | Not shared between tabs, users, browsers or devices; a new tab starts at the welcome screen |
| Size | Fits workbooks up to roughly 3 MB; larger files still work but must be uploaded again after a reload (a notice says so) |
| Older data | Any workbook saved by an earlier version of the page (IndexedDB) is deleted on load |

### 4.2 `CONFIG` (layout and rules)

| Setting | Default | Meaning |
|---|---|---|
| `mainSheet` | `Efficiency measures _Summary` | Learning-data worksheet (exact name) |
| `trainingListSheet` | `Sheet1` | Maintained training list |
| `employeeSheet` | `Employees` | Team roster |
| `headerRow` | `3` | Header row in the main sheet |
| `firstDataRow` | `5` | First learning record row |
| `firstEmployeeColumn` | `I` | First employee-name column in the header row |
| `trainingListColumn` / `trainingListFirstRow` | `B` / `2` | Where the training list starts |
| `employeeRow` / `employeeFirstColumn` | `2` / `B` | Where the roster starts |
| `requiredHeaders` | Method, Training/Learning, Sub Topic, Year | Missing any of these stops loading |
| `targetHours` | `40` | Learning target in hours |
| `maxPeople` | `4` | Maximum people in Compare |
| `colors` | Blue, Purple, Green, Amber | Colours for P1–P4 |
| `belowTargetColor` | Red `#DC2626` | Totals > 0 and below target |
| `zeroHoursColor` | Grey `#94A3B8` | Totals of exactly 0 hours |
| `onTargetColor` | Teal `#0D9488` | Target achieved (summary charts) |
| `kpiCategories` | BTV Batch 1/2, PLATON Basic/Advanced | Readiness programme matching rules |
| `trainingTypes` | Behavioral / Technical | Method → type mapping for the timeline |
| `timelineCollapsedCount` | `8` | Cards shown per year before "Show more" |
| `chartTopCategories` | `7` | Categories shown before grouping into "Other" |

---

## 5. Workbook format (fixed)

The workbook must be an **`.xlsx`** file with these three worksheets. Worksheet names must match **exactly**, including spaces and capitals.

### 5.1 Main sheet — `Efficiency measures _Summary`

Matrix layout: one **row per learning record**, one **column per employee**.

```
        A        B        C                   D           E                 F      G         H                      I            J            ...
Row 3   Sl_No    Method   Training/Learning   Sub Topic   Trainer name...   Year   Remarks   TOTAL LEARNING...      Employee 1   Employee 2   ...
Row 4   (formula totals per employee — ignored by the dashboard)
Row 5   1        TecCom   E-drive Module 2                                  2026             (formula)              32                        ...
Row 6   ...
```

| Area | Rule |
|---|---|
| **Header row (row 3), columns A–H** | Headers are found by name (case and extra spaces ignored), not by fixed position. **Required:** `Method`, `Training/Learning`, `Sub Topic`, `Year`. **Optional:** `Sl_No`, `Trainer name (if internal)`, `Remarks`, `TOTAL LEARNING (In Hours)` |
| **Employee names (row 3, from column I)** | One name per column, read left to right **until the first blank header cell**. Anything after a gap (for example helper lists far to the right) is ignored |
| **Row 4** | Ignored. Excel formula totals are not used |
| **Column H (TOTAL LEARNING)** | Ignored for calculations, so hours are never double-counted |
| **Records (row 5 onward)** | Each row is one training record. A row is ignored as empty when Method, Training/Learning and Sub Topic are blank **and** no employee has hours in it |
| **Hour cells** | See 5.4 |
| **Year** | Number (`2026`) or text (`"2026"`). Blank → shown as **Year not specified** |

### 5.2 Roster sheet — `Employees`

| Area | Rule |
|---|---|
| Names | **Row 2, from column B**, one per column, left to right |
| Blanks | Skipped |
| Duplicates | Counted once; reported in Data quality |
| Role | Defines the **team** used for KPIs and summary charts |

### 5.3 Training list — `Sheet1`

| Area | Rule |
|---|---|
| Trainings | **Column B, from row 2**, one per row |
| Role | Used only to warn when a record's Training/Learning value is not in this list. If the sheet is missing, the check is skipped with a warning |

### 5.4 Learning-hour cell rules

```mermaid
flowchart TD
    A[Employee hour cell] --> B{Empty?}
    B -- Yes --> Z0[Counted as 0]
    B -- No --> C{Formula?}
    C -- Yes, no saved value --> W1[0 + warning<br/>'Formula without cached value']
    C -- Yes, saved value --> D
    C -- No --> D{Value type}
    D -- Number --> E{Negative?}
    D -- Text that is a number e.g. '2' --> E
    D -- Other text e.g. 'two', 'NA' --> W2[0 + warning<br/>'Invalid learning-hour value']
    D -- "Excel error e.g. #35;N/A" --> W2
    E -- Yes --> W3[Excluded + warning<br/>'Negative hours']
    E -- No --> F{Greater than 0?}
    F -- Yes --> OK[Counted as learning hours<br/>person = participant]
    F -- No --> Z0
```

### 5.5 Name matching

- Leading/trailing spaces are removed and repeated spaces collapsed.
- Matching between sheets is **not case-sensitive** (`AJITH KUMAR` = `Ajith Kumar`).
- The name shown is taken from the `Employees` sheet when available.

---

## 6. Data restrictions and assumptions

| # | Restriction / assumption | Effect if not followed |
|---|---|---|
| 1 | File must be **`.xlsx`** (Office Open XML). `.xls`, `.csv`, `.xlsb` are not supported; `.xlsm` is untested | Error: "not a valid .xlsx workbook" |
| 2 | Workbook must **not be password-protected / encrypted** | Error: workbook cannot be read |
| 3 | The three worksheet names must match exactly | Error panel lists the missing sheet and available sheets |
| 4 | Required headers must be in **row 3, before column I** | Error panel lists the missing header(s) |
| 5 | Employee header columns must be **contiguous** from column I | Names after the first blank header cell are ignored |
| 6 | Hours must be **numbers** (decimals allowed) | Text → counted as 0 and reported |
| 7 | Hours must be **≥ 0** | Negative values excluded and reported |
| 8 | The workbook must have been **saved by Excel** (so formula results are cached) | Formula cells without saved results count as 0 and are reported |
| 9 | Year should be a plain **number or text**, not an Excel date | Dates appear as serial numbers |
| 10 | Merged cells: only the top-left cell holds the value | Other merged cells are treated as blank |
| 11 | Hidden rows/columns are **still read** | Hide ≠ exclude; delete or clear data instead |
| 12 | KPI and summary totals include **only people in the `Employees` sheet** | Others are shown in Compare ("Main sheet only"), timeline and Data quality |
| 13 | Totals are calculated from individual cells | May differ from Excel's cached `TOTAL` cells if those are hard-coded or stale |
| 14 | Saving the workbook while the dashboard reads it | Read may fail; click **Refresh data** again |
| 15 | File name and location are **not** restricted | Any `.xlsx` loads if the format matches |
| 16 | Only the **last valid** workbook is kept, per tab | Uploading a new valid workbook replaces it; invalid uploads are not kept |
| 17 | The workbook is kept only while the tab is open | Closing the tab deletes it; upload again in a new tab |

---

## 7. Data flow

```mermaid
flowchart LR
    subgraph Source
        X[(Workbook .xlsx)]
    end
    subgraph Load["Load (every page load / Refresh)"]
        L1{File uploaded / dropped<br/>this session?}
        L2[Read uploaded file]
        L3[(Copy kept for this tab<br/>sessionStorage)]
        L4[Welcome screen]
    end
    subgraph Parse["Parse (in the browser, read-only)"]
        P1[Unzip .xlsx]
        P2[Read sheets:<br/>Main · Employees · Sheet1]
        P3[Validate sheet names<br/>and required headers]
        P4[Normalise names, years,<br/>hours → records]
        P5[Collect data-quality warnings]
    end
    subgraph Show["Dashboard"]
        S1[KPI cards & summary charts]
        S2[Compare up to 4 people]
        S3[Timeline by year]
        S4[Participants panel]
        S5[Data quality panel]
    end
    X --> L1
    L1 -- Yes --> L2 --> P1
    L1 -- No --> L3
    L3 -- tab copy --> P1
    L3 -- nothing saved --> L4 --> L2
    P5 -. valid upload .-> L3
    P1 --> P2 --> P3
    P3 -- Missing sheet/header --> E[Error panel<br/>+ welcome screen]
    P3 -- OK --> P4 --> P5
    P4 --> S1 & S2 & S3
    S1 & S3 -- click a bar/card --> S4
    S1 -- click a name --> S2
    P5 --> S5
```

## 8. User flow

```mermaid
flowchart TD
    A([Open dashboard]) --> B{Workbook loaded?}
    B -- No --> C[Welcome screen] --> D[Upload workbook /<br/>drag & drop] --> B
    B -- Yes --> E[Choose Year<br/>or All Years]
    E --> F[Review Learning summary<br/>KPIs + charts]
    F --> G{Need detail?}
    G -- Who attended? --> H[Click method / readiness bar<br/>or timeline card] --> I[Participants panel<br/>names · hours · Excel column · rows]
    G -- Compare people --> J[Click names in<br/>'Learning hours by employee'<br/>or use Compare rows]
    J --> K[Narrow with Method → Training →<br/>Sub Topic → Learning Hours]
    K --> L[Charts + person cards update]
    L -- click selected name again --> M[Person removed]
    G -- Data issues? --> N[Open Data quality panel]
    I & L & N --> O[Edit workbook in Excel → Save]
    O --> P[Click Refresh data] --> F
```

---

## 9. Dashboard sections

### 9.1 Header
Workbook name, last modified time, last refresh time, selected year, learning target, **Year** filter, **Refresh data**, **Upload workbook…**

Before a workbook is loaded, a **welcome screen** shows the drop area, the **Upload workbook…** button, a privacy note and the required-format checklist (generated from `CONFIG`). An amber notice appears only when something needs attention (for example the original file could not be re-read).

### 9.2 Learning summary (team = `Employees` sheet)

| KPI | Calculation |
|---|---|
| Total Employees | Unique names in the `Employees` sheet |
| Total Learning Hours | Sum of all team members' hour cells for the selected year |
| Avg Hours / Employee | Total Learning Hours ÷ Total Employees (0.00 if no employees) |
| At or above 40 h | Team members with total ≥ 40.00 |
| Below 40 h | Team members with total < 40.00 (includes 0 h) |
| BTV Training – Batch 1 / 2 | Method or Training contains "BTV" **and** Training/Learning = "Batch 1" / "Batch 2" |
| PLATON Readiness – Basic | Method or Training contains "PLATON" **and** Sub Topic = "Basic" or "Basics" |
| PLATON Readiness – Advanced | Method or Training contains "PLATON" **and** Sub Topic = "Advanced" |

Readiness cards show participants (hours > 0), hours, and % of the team.

**Charts:** Target achievement donut (achieved / below / no hours), Learning hours by method, Readiness programmes, Learning hours by year (click to filter), Learning hours by employee (click to add/remove from Compare).

### 9.3 Compare team learning
- Starts with one row; **Name is required**. `+` adds rows up to 4; `−` appears when there are 2+ rows.
- Filters cascade: **Method → Training/Learning → Sub Topic → Learning Hours** and only list values the selected person actually has.
- A person cannot be selected twice.
- Charts: Total learning hours (with target line), Breakdown by Method, Breakdown by Training/Learning, and person cards (filtered total, overall total, items, % of target, unique methods, training categories).

### 9.4 Training conducted by year
Records grouped by **Year → Method + Training/Learning + Sub Topic**, with participants, hours and average. Filter by **All / Technical / Behavioral** and search by training or sub topic. Click a card to open the participants panel.

Type mapping (configurable in `CONFIG.trainingTypes`):

| Type | Rule |
|---|---|
| Behavioral | Method = `HR`, or Training/Learning contains "Behavioral"/"Behavioural" |
| Technical | Method is one of TecCom, Product Hands On learning, Focus competency group, PLATON, BTV training, External Training, Internal training, Upskilling |
| Other/Unclassified | Anything else (e.g. Daimler Mandatory Training, Others) |

### 9.5 Participants panel (verification)
Opens on the right side. Shows the title, participants, total and average hours, **workbook row numbers**, and each participant's **hours and Excel column** with a Compare/Remove button. **Copy list** copies names and hours for pasting into Excel.

### 9.6 Data quality
Lists, without stopping the dashboard:
- Roster names missing from the main sheet (shown with 0.00 h)
- Main-sheet names missing from the roster
- Duplicate names
- Invalid, negative, or uncached formula hour cells (with cell reference)
- Records missing Year / Method / Training/Learning (with row number)
- Training values not in `Sheet1`
- Headers that could not be matched
- Count of empty rows ignored

---

## 10. Colour and status rules

| Total hours | Colour | Status text |
|---|---|---|
| 0.00 | Grey | No learning hours (below 40-hour target) |
| > 0 and < 40.00 | Red | Below 40-hour target |
| ≥ 40.00 | Person colour (Compare) / Teal (summary) | Target achieved |

- Status is evaluated on the **unrounded** value (39.999 shows as 40.00 but is still below target).
- All hours are displayed with **two decimals**.
- Status is always shown as text as well as colour.

---

## 11. Browser support

| Browser | Status |
|---|---|
| Microsoft Edge / Google Chrome (current) | Fully supported, including re-reading a selected file on Refresh |
| Firefox (113+) | Supported; re-select the file after each save when opened from disk |
| Safari (16.4+) | Supported |
| Internet Explorer | Not supported |

Animations are switched off automatically when the operating system's "reduce motion" setting is on.

---

## 12. Troubleshooting

| Message / symptom | Cause | Fix |
|---|---|---|
| Welcome screen instead of the dashboard | Nothing uploaded in this tab yet (new tab, tab was closed, other device) | **Upload workbook…** or drag and drop |
| "… was renamed, moved or deleted" / "did not allow re-reading" | The original file can't be re-read | The copy uploaded earlier in this tab is shown; upload the workbook again to update |
| "This browser blocked direct access to …" | Browser limits direct file access (e.g. VS Code's built-in browser) | Click **Upload workbook…** again — the standard file chooser is used |
| "… is too large to keep for this tab" | Workbook exceeds the tab storage limit | The dashboard still works; upload again after reloading |
| "… is not an .xlsx workbook" | A `.csv`, `.xls` or other file was chosen | Save from Excel as **Excel Workbook (*.xlsx)** |
| "Worksheet … was not found" | Sheet renamed | Restore the exact sheet name or update `CONFIG` |
| "Missing expected header(s)…" | Header text changed or moved | Restore header text in row 3 (columns A–H) |
| "not a valid .xlsx workbook" / "corrupt" | Still saving, or damaged | Wait for Excel to finish saving, then **Refresh data** |
| A person is missing from KPIs | Not in the `Employees` sheet | Add the name to row 2 of `Employees` |
| Total differs from Excel's TOTAL column | Excel total cell is hard-coded/stale | Dashboard uses individual cells; fix the Excel formula |
| Numbers didn't change after editing | The tab still holds the old copy | Save in Excel, then **Upload workbook…** again (or **Refresh data** in Edge/Chrome) |

---

## 13. Verification checklist

- [ ] Dashboard loads and shows the workbook name and modified time.
- [ ] Total Employees equals the number of names in `Employees` row 2.
- [ ] A participants panel's names and hours match the Excel row/columns shown.
- [ ] Year filter changes KPIs, charts, Compare and timeline.
- [ ] Clicking a name in "Learning hours by employee" adds it to Compare; clicking again removes it.
- [ ] Totals of 0 h are grey, < 40 h red, ≥ 40 h coloured.
- [ ] Edit an hour in Excel → Save → **Upload workbook…** → value updates.
- [ ] Upload a workbook with any file name from any folder → dashboard loads with the same figures.
- [ ] Reload the page in the same tab → the workbook loads automatically.
- [ ] Close the tab and open the site in a new tab → welcome screen (data cleared).
- [ ] The workbook's modified time is unchanged after using the dashboard.
