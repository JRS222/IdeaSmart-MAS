# First-Time Use Guide

A step-by-step walkthrough from double-clicking the script to a fully working setup. Each step takes a couple of minutes. Details and reference material live in [README.md](README.md).

---

## 1. What you need

- A Windows PC on the **USPS intranet** (the download features talk to MTSC / eMARS / eDAC).
- **Microsoft Excel** installed (parts-room workbooks and parts books are built through Excel).
- PowerShell — already on every Windows machine.
- The `UI-Script.ps1` file and this guide.

---

## 2. Launch the app

Right-click `UI-Script.ps1` → **Run with PowerShell**.

If Windows blocks the script with an execution-policy error, open PowerShell once and run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Then launch the script again.

---

## 3. Installation wizard

On first launch, the **Installation Setup** wizard opens automatically.

1. **Data root** — pick where your data lives, e.g. `C:\Maintenance` (or a network share if you move between PCs).
2. **CSV source folder** (optional) — only needed if you already have `Sites.csv`, `Parsed-Parts-Volumes.csv`, and `Machines.csv` from another install. Leave blank for a fresh start.
3. The wizard creates the full folder tree and seed files. **It never overwrites existing files**, so re-running it later is safe.

When the main window opens, you're on the **Parts Books** tab. Do the next two steps before anything else — the app needs the site and volume lists for most downloads.

---

## 4. Fill in the Sites list

Go to **Settings** → **Installation** → **Regenerate Sites List**.

- On success: a dialog says "Fetched and wrote N entries" (~290 sites).
- Verify: `Dropdown CSVs\Sites.csv` exists with header `Site ID,Full Name`.

> If the automatic fetch fails, a paste dialog appears instead — the app opens the eMARS page in your browser; paste the site `<select>` HTML as instructed. That's the fallback and works just as well.

---

## 5. Fill in the Parts Volumes list

Same tab: **Regenerate Parts Volumes List**.

- On success: a dialog reports ~96 handbook volumes.
- Verify: `Dropdown CSVs\Parsed-Parts-Volumes.csv` has header `MS Book No,Volume,Full Name` and no "All MS Handbooks" placeholder row.

---

## 6. Download your local parts room

Go to **Actions** → **Maintenance** → **Update Parts Room**.

Pick your site from the list. The app downloads the site's stockroom data and merges it into `Parts Room\<YourSite>.csv` (and builds the `.xlsx`). This gives Search something to find and the Parts Book Creator something to cross-reference against.

> The other download buttons live here too: **Add Same Day Parts Room** (other sites' stock), and **Update Files** (one-button daily refresh of everything).

---

## 7. Create your first parts book

1. Go to **Parts Books** → **Create Parts Book**.
2. Check one volume you want (e.g. a handbook for a machine you work on) and click **Process Selected**.
3. The app fetches the handbook tree, downloads every figure, converts them to CSVs, merges against your parts room, and builds an Excel workbook — one worksheet per section, hyperlinked to the figures.
4. A new button for the book appears on the Parts Books tab. Click it to open the workbook.

> If the automatic fetch fails it opens the tree page in your browser and asks you to paste the page source — follow the on-screen instructions; the result is identical.

---

## 8. Configure identity and your work budget

Go to **Settings**:

- **Identity → Technician Name** — this is stamped on every Historian event. Set it now.
- **Work Budget** — set Day Length, Lunch, breaks, wash-up, paperwork, and your Work Target. Zero values are skipped.
- Click **Save Settings**.

---

## 9. Log your first work

Go to **Work Tracking** → **Job Log**:

1. **+ New Reactive Call** or **+ New Work Order**.
2. Machine ID: `DBCS 49`, `dbcs-49`, or just `49` — anything reasonable is normalized.
3. Set Start/End (the **Now** buttons help), parts if any, and a description. The editor prompts for a Work Order # when required.
4. **Save**. The bottom panel shows your work budget for today updating live.

**PM Checklists:** use **+ Import PM Checklist** and paste the checklist popup's HTML from eDAC. Mark tasks complete as you go; **Send Checklist** pushes them to the Historian.

**Weekly Worksheet:** set the eDAC worksheet URL in Settings, then **Fetch from eDAC**. After logging work, click **Sync Actual Times** — your jobs' hours land on the worksheet automatically, and **Send Day** pushes the day to the Historian.

---

## 10. Your daily rhythm

| When | What | Where |
|---|---|---|
| Start of day | **Fetch from eDAC** (or Paste HTML) | Work Tracking → Weekly Worksheet |
| After work calls | Log jobs, mark PM tasks | Job Log / PM Checklist |
| Before leaving | **Sync Actual Times**, then **Send Day** | Weekly Worksheet |
| As needed | **Update Files** (fresh stock data), **Send all ready** (unsent jobs → Historian) | Actions / Job Log |
| End of week | **Generate Report** | Historian |

---

## 11. Where everything lives

| Thing | Location |
|---|---|
| Configuration | `<Data Root>\Config.json` |
| Sites / volumes / machines | `<Data Root>\Dropdown CSVs\` |
| Local parts room | `<Data Root>\Parts Room\<YourSite>.csv/.xlsx` |
| Same-day rooms | `<Data Root>\Parts Room\Same Day Parts Room\` |
| Installed books | `<Data Root>\Parts Books\<Book Name>\` |
| Jobs / breaks / events | `<Data Root>\Work Tracking\Jobs.json, Breaks.json, Historian.jsonl` |
| Worksheets / PM checklists / reports | `<Data Root>\Work Tracking\Weekly Worksheets\, PM Checklists\, Reports\` |
| App log | `UI.log` (next to `UI-Script.ps1`) |
| Failed-fetch snapshots | `<Data Root>\Debug Dumps\` |

---

## 12. If something goes wrong

1. **Read `UI.log` first** — every download, parse, and error is logged there.
2. A fetch that fails **falls back to a paste dialog** — finish the task manually; you're never blocked.
3. Failed automatic fetches save the page to **`<Data Root>\Debug Dumps\`** for diagnosis.
4. For anything else, the [README.md](README.md) troubleshooting section has the common fixes (Excel hangs, execution policy, missing CSVs, stale list views).

---

**Welcome aboard — and when in doubt: Update Files, Sync Actual Times, Send Day.**
