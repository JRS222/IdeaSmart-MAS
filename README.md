# Comprehensive Maintenance Software Guide

## Overview

This maintenance software is a PowerShell-based GUI application (built on Windows Forms) that manages:

- **Parts Rooms** — local inventory plus same-day parts rooms at other sites
- **Parts Books** — searchable cross-reference across MS handbook volumes
- **Work Tracking** — reactive calls, work orders, PM checklists, weekly eDAC worksheets, and an append-only event log ("Historian")
- **Search** — NSN, OEM/Part Number, and description search across all sources
- **Reporting** — on-demand work-tracking reports exportable as text

The application is driven by a single script, `UI-Script.ps1`, plus a small on-disk data tree (config + CSVs + JSON). An installation wizard is built into the UI and can be re-run from the **Settings** tab at any time.

---

## Initial Setup

### Step 1: First Launch

Right-click `UI-Script.ps1` and choose **Run with PowerShell**.

If no `Config.json` exists yet, the **Installation Setup** wizard opens automatically. It will:

1. Ask for a **Data root** folder (e.g. `C:\Maintenance` or a network share).
2. Optionally ask for a **CSV source folder** containing pre-built `Sites.csv`, `Parsed-Parts-Volumes.csv`, and `Machines.csv`. If left blank, header-only blanks are created and you can fill them in later.
3. Create the following directory tree under the chosen root:

```
<Data Root>/
├── Config.json
├── Dropdown CSVs/
│   ├── Sites.csv
│   ├── Parsed-Parts-Volumes.csv
│   └── Machines.csv
├── Parts Room/
│   ├── <LocalSite>.csv
│   ├── <LocalSite>.xlsx
│   ├── Same Day Parts Room/
│   │   ├── <Site A>.csv
│   │   └── <Site A>.html
│   └── error_log.txt
├── Parts Books/
│   └── <Book Name>/
│       ├── Volumes-to-URL.csv
│       ├── SectionNames.txt
│       ├── HTML and CSV Files/
│       │   ├── Figure <sec>-<fig>.html
│       │   └── Figure <sec>-<fig>.csv
│       ├── CombinedSections/
│       │   └── Section <n>.csv
│       └── <Book Name>.xlsx
└── Work Tracking/
    ├── Jobs.json
    ├── Breaks.json
    ├── Historian.jsonl
    ├── Weekly Worksheets/
    │   └── YYYY-MM-DD.json
    ├── PM Checklists/
    │   └── <ACRONYM>_<MNO>_<CHECKNO>_<YYYY-MM-DD>.{html,json}
    └── Reports/
```

Existing files are never overwritten by the wizard.

### Step 2: Required CSV Files (in `Dropdown CSVs/`)

| File | Header | Required? | Purpose |
|---|---|---|---|
| `Sites.csv` | `Site ID,Full Name` | Yes, for parts-room downloads and same-day sites | Site lookup. |
| `Parsed-Parts-Volumes.csv` | `Full Name,MS Book No,Volume` | Yes, for Parts Book Creator | Catalog of MS parts handbook volumes. |
| `Machines.csv` | `Machine Acronym,Machine Number` | Seeded automatically; new machines are appended when jobs are saved | Machine registry. |

Example `Sites.csv`:

```csv
Site ID,Full Name
ABC,Alpha Bravo Center
XYZ,X-Ray Yankee Zone
```

You can regenerate `Sites.csv` and `Parsed-Parts-Volumes.csv` at any time from the **Settings** tab by pasting the `<select>` HTML from the relevant MTSC / eMARS page.

> **Note:** Earlier versions of this software used `Actions.csv`, `Causes.csv`, and `Nouns.csv` for a Call Log workflow. Those files are **no longer used** and are not created by the installer. If they exist on disk they are harmless — the script never reads them.

---

## The Interface at a Glance

The main window has five top-level tabs:

1. **Parts Books** — open the local parts room, run the Parts Book Creator, launch any installed book.
2. **Work Tracking** — Job Log, PM Checklist, Weekly Worksheet, and Historian sub-tabs.
3. **Actions** — maintenance and housekeeping commands.
4. **Search** — NSN / OEM / Description search across all parts sources.
5. **Settings** — identity, work-tracking config, work budget, and installation utilities.

---

## Parts Books Tab

### Buttons
- **Open Parts Room** — opens the local parts-room `.xlsx` workbook in Excel.
- **Create Parts Book** — launches the Parts Book Creator (see below).
- **<Book Name>** — one button per installed book; opens its `.xlsx`.

### Creating a Parts Book
1. Click **Create Parts Book**.
2. If `Parsed-Parts-Volumes.csv` is missing or empty, the Parts Volumes wizard will run first — it opens MTSC, asks you to paste the parts-book dropdown HTML, and writes the CSV.
3. In the book picker, check the volumes you want and click **Process Selected**.
4. For each book, the UI:
   - Opens the MTSC tree page in your browser.
   - Asks you to paste the page source (Ctrl+U, Ctrl+A, Ctrl+C).
   - Parses the tree into `Volumes-to-URL.csv` + `SectionNames.txt`.
   - Downloads every figure HTML, converts it to CSV.
   - Merges section CSVs against your local parts room.
   - Builds a workbook with one worksheet per section, hyperlinked to the original figure HTML.
   - Renames worksheets to full section names.

Already-installed books are greyed out in the picker.

---

## Work Tracking Tab

### Job Log

Records reactive calls and work orders against a machine.

- **+ New Reactive Call** and **+ New Work Order** open the Job editor.
- **Refresh** re-reads `Jobs.json`, `Breaks.json`, and the budget panel.
- **Send all ready** pushes every unsent, valid job into the Historian.

Double-click any row to edit or delete it.

**Job editor fields:**
- **Machine ID** — any of `DBCS 49`, `dbcs-49`, `DBCS_0049` normalize to `DBCS 49`. Bare numbers (e.g. `49`) become `MNO 49`.
- **Start** / **End** — 24-hour HH:mm, with **Now** buttons.
- **Hours** — computed live from Start/End.
- **Work Order #** — required if duration > 30 min or any parts are attached.
- **Description** — required under the same rule.
- **Parts Used** — list with **+ Add Part** / **Remove**. Part search supports NSN and Description; quantities are capped at the on-hand Qty.

New machine IDs are automatically registered in `Machines.csv` when a job is saved.

**Bottom panel:**
- **Work Budget (Today)** — shows scheduled non-work time (lunch, wash-up, breaks, startup, paperwork), your work target, hours logged today (from jobs + selected PM tasks), break time, and either the remaining hours or an **OVER BY** warning. If work exceeds the target, a grievance reminder is shown.
- **Breaks & Non-Work Time (Today)** — log paid breaks, lunch, wash-up, etc. Each break has a type, start time, duration, and notes.

### PM Checklist

Manages the PM checklists you pull from eDAC.

- **+ Import PM Checklist** — paste the raw HTML of a checklist popup; it's parsed, archived as `<ACRONYM>_<MNO>_<CHECKNO>_<YYYY-MM-DD>.{html,json}`, and displayed.
- **Scan Archive** — reparses any non-canonical HTML in the archive folder and collapses duplicates.
- **Reload** — re-reads all JSONs from disk.
- **Send Checklist** — pushes the selected date's completed tasks to the Historian as one event per machine.
- **Date picker** / **Today** — filters the machine tabs.
- **Open Archive Folder** — opens the archive directory.

Each machine is a tab. The task grid supports:
- Checkbox to mark a task complete.
- Double-click to open the task editor (mark complete, override estimated time, read the full instruction).
- Row coloring:
  - 🟧 **Orange** — overdue + Power Off
  - 🟩 **Green** — overdue + Power On
  - 🟨 **Yellow** — future due date

Selections persist in a per-checklist `.session.json` sidecar.

### Weekly Worksheet

Works with the eDAC **Weekly Assignment Worksheet**.

- **Fetch from eDAC** — downloads the configured `WorkTracking.eDacUrl` and imports every date section it contains.
- **Paste HTML** — imports from clipboard-style pasted HTML.
- **Sync Actual Times** — applies the D1 rule:
  - Checklist No = `000` and PM data exists → actual hours = sum of selected PM task times.
  - No PM data → actual hours = estimated hours.
  - Otherwise → leave manual entry alone.
- **Send Day** — pushes the current worksheet to the Historian as a single event (rows flagged **AUDIT** are noted but not counted in the summary).
- **Reload** / **Today** — reloads the JSON for the selected date.
- **Date picker** — switches to a different day's worksheet.

Double-click a row to override its actual hours or clear an auto value. Flagged rows are highlighted yellow. Rows auto-filled by sync show a ⚡ indicator.

### Historian

Read-only viewer for the append-only `Historian.jsonl` event log.

- **Filters** — Kind (`reactive`, `workorder`, `pm.checklist`, `worksheet`), Machine, and free-text search across summary and detail.
- **Refresh** — re-reads and re-filters.
- **Open Historian File** — opens the JSONL in your default editor.
- **Debug Dump** — writes a diagnostic snapshot to the console for troubleshooting.
- **Generate Report** — see below.
- **Clear Filters** — resets to All / (any) / empty.

Double-click a row for a full detail view including the raw event payload.

**Generate Report** produces a monospaced text report covering any selected subset of:
- **[1] Jobs** — reactive calls and work orders with totals.
- **[2] PM Checklists** — per-machine checklist, task count, hours.
- **[3] Weekly Worksheets** — row counts, estimated vs. actual with diffs.
- **[4] Work Budget Roll-up** — per-day job + PM hours against your target, plus break time logged.

Reports can be copied to the clipboard or saved under `Work Tracking/Reports/`.

---

## Actions Tab

### Maintenance
| Action | Behavior |
|---|---|
| **Update Files** | Master refresh — same-day sites, local parts room, then parts books. |
| **Update Parts Books** | Refreshes Qty and Location in every section CSV and worksheet. |
| **Update Parts Room** | Downloads the local site's stockroom HTML and merges with the existing CSV (preserves book-reference columns, marks missing parts as "Not in current inventory"). |

### Parts
| Action | Behavior |
|---|---|
| **Take a Part Out** | *(not yet implemented)* |
| **Search for a Part** | Switches to the Search tab. |
| **Request a Part to be Ordered** | *(not yet implemented)* |

### Work Management
| Action | Behavior |
|---|---|
| **Request a Work Order** | *(not yet implemented)* |
| **Make an MTSC Ticket** | Opens the MTSC ticket login page. |
| **Search Knowledge Base** | Opens a small dialog; submits the query to ServiceNow KB search. |

### Parts Book / Parts Room Registry
| Action | Behavior |
|---|---|
| **Add Parts Book** | Adds a volume from `Parsed-Parts-Volumes.csv` to the config and creates its folder. |
| **Remove Parts Book** | Removes a book from the config (files on disk are preserved). |
| **Add Same Day Parts Room** | Pick sites from `Sites.csv`; the app downloads their stockroom HTML and creates CSVs. |
| **Remove Same Day Parts Room** | Removes sites from the config (cached CSVs are preserved). |
| **Add 1-Day / 2-Day Parts Room** | *(future)* |

> The **Actions** tab is a UI feature. It is unrelated to the legacy `Actions.csv` file, which is no longer used.

---

## Search Tab

Searches the local parts room, all same-day parts rooms, and every installed book's combined sections.

### Inputs
- **NSN** — digits plus `*` wildcard. Non-digits are stripped, so `1234-56-78` and `12345678` behave the same. `1234*` finds all NSNs starting with `1234`.
- **OEM** — alphanumerics plus `*`. Punctuation is stripped and case is ignored.
- **Description** — space-separated tokens; every token must be present.

### Result Panels
1. **Availability** — Local parts room.
2. **Same Day Parts Availability** — One row per matching part per site, with `Site Name`.
3. **Cross Reference** — Combined section rows from every book.

### Bottom Buttons
- **Open Figures** — opens the HTML figure for each checked cross-reference row.
- **Take Part(s) Out** — enabled when any Availability / Same Day row is checked.

---

## Settings Tab

### Identity
- **Technician Name** — stamped on every Historian event as `capturedBy`.
- **Supervisor Email** — informational.

### Work Tracking
- **eDAC Worksheet URL** — used by **Fetch from eDAC**.
- **Reactive → WO threshold (min)** — currently informational; default 15.

### Work Budget
- Day Length (hours)
- Lunch (min, 0 to skip)
- Wash-up at Lunch (min)
- Startup (min)
- Paid Break 1 / 2 (min)
- Wash-up End-of-Day (min)
- Paperwork (min)
- Work Target (hours)

Zero-valued items are skipped in the budget roll-up.

### Installation
- **Run Installation Wizard** — reopens the setup dialog. Safe to run again; existing data is preserved.
- **Open Root Folder** / **Open Scripts Folder** — shortcuts.
- **Regenerate Parts Volumes List** — re-runs the parsed-parts-volumes paste wizard.
- **Convert Figure HTML → CSV** — batch-converts a folder of `Figure <sec>-<fig>.html` files into CSVs.
- **Regenerate Sites List** — re-runs the sites paste wizard.

Buttons at the bottom of the tab:
- **Save Settings** — writes `Config.json`.
- **Reload from File** — discards unsaved edits and re-reads `Config.json`.

---

## Machine ID Convention

Machine IDs are canonicalized everywhere:

| Input | Stored as |
|---|---|
| `DBCS 49` | `DBCS 49` |
| `dbcs-49` | `DBCS 49` |
| `DBCS_0049` | `DBCS 49` |
| `49` | `MNO 49` |

The same number with different acronyms is allowed (`ASD 47` and `DBCS 47` can coexist). Saving a new `ACRONYM N` combination adds it to `Machines.csv`.

---

## Historian Format

`Historian.jsonl` is newline-delimited JSON. Every event includes:

- `id` — GUID
- `source` — `work_tracking`
- `kind` — `reactive`, `workorder`, `pm.checklist`, or `worksheet`
- `machineId`
- `eventTime` — ISO 8601
- `summary` — one-line description
- `detail` — multi-line detail block
- `payload` — structured data (the job, the checklist, or the worksheet)
- `capturedAt`, `capturedBy`

The file is **append-only**. Once an event is written it cannot be removed from the UI.

---

## Known Behaviors & Limitations

### Data Refresh
Several lists hold data in memory and are only re-read when you click **Refresh** or change the date filter. If you edit `Jobs.json`, `Breaks.json`, or add new PM checklists on disk outside the app, use **Refresh** (or **Reload** on the PM Checklist sub-tab).

### Excel via COM
Updating parts-room workbooks and creating parts books uses Excel COM automation. Excel must be installed on the machine running the UI. Very large parts rooms can take a while — the progress bar updates as cells are written.

### Historian is Append-Only
There is no built-in way to delete or edit an event once it has been pushed. If you send a job and then realize it was wrong, edit the job and re-send; both events will remain in the log.

### Parts on Jobs
Parts attached to a job are stored inside `Jobs.json` with the job itself. They are **not** removed from the parts-room CSV when you click **Take Part(s) Out** — that action is currently a stub.

### Search Cross-Contamination (historical)
Earlier versions of this software had a shared search state between the old Labor Log and Search tabs. That design is gone: the Search tab now runs its queries independently, and the Historian has its own text filter. If you see stale results, click **Refresh** or clear the search box.

### Notification Indicators
Work orders that still need a work-order number are validated at save time (the Save button refuses the job), so there is no longer a passive "notification dot." The work-order number field is simply required when the rule applies.

### Actions Not Yet Implemented
The following buttons display an informational dialog and do nothing else:

- **Take a Part Out**
- **Request a Part to be Ordered**
- **Request a Work Order**
- **Add 1-Day Parts Room**
- **Add 2-Day Parts Room**

---

## Best Practices

1. **Back up `Historian.jsonl` and `Jobs.json` regularly** — they hold the only durable record of work tracking.
2. **Run Update Files at least once a day** to keep availability fresh. It updates same-day rooms, local parts room, and parts books in one pass.
3. **Assign work order numbers promptly.** The job editor enforces this on save (>30 min or any parts attached), but it cannot retroactively flag jobs created outside the UI.
4. **Use the Historian report** rather than screen-scraping the list when you need something to hand to a supervisor. Reports are self-contained text.
5. **Keep `Machines.csv` curated.** New entries are appended automatically; if you see a bad acronym, edit the file directly (with the app closed).
6. **Don't edit `Config.json` by hand** unless you know the schema. Use **Settings** → **Save Settings** where possible.
7. **Verify search results** by spot-checking against `Parts Room/<LocalSite>.csv` when a query returns unexpectedly few or many rows.

---

## Troubleshooting

- **App won't start / no tabs appear** — confirm `Config.json` exists under `Scripts/PartsMgmt/`. If not, the installation wizard should have launched. If it didn't, run `UI-Script.ps1` from an elevated PowerShell prompt and watch the console.
- **"No CSV files found in …"** — the local parts room CSV is missing. Run **Actions → Update Parts Room** and pick a site, or drop a CSV into `Parts Room/`.
- **"Multiple CSV files found in …"** — the Search tab expects exactly one `.csv` (plus the `.xlsx`) at the root of `Parts Room/`. Move extras into `Same Day Parts Room/`.
- **Excel operations hang** — close any open workbooks with the same name, and end leftover `EXCEL.EXE` processes from Task Manager. The app releases COM objects after each operation.
- **Historian shows fewer events than expected** — filters persist until cleared. Click **Clear Filters** or run **Debug Dump** to see the raw event count.
- **A PM checklist is not appearing** — click **Scan Archive**. Canonical files (`ACRONYM_MNO_CHECKNO_YYYY-MM-DD.json`) require a matching `.json`; raw HTML files are parsed on scan.
- **PowerShell refuses to run the script** — set the execution policy for the current user:

  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```

- **Check the log** — `UI.log` in the script directory records every action, warning, and error. Use it first when something looks wrong.
