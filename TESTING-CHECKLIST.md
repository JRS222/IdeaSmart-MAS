# Testing Checklist - Auto-Fetch & Work-Order Duplication Fixes

Date: today's build. Test in this order; each section is independent enough to skip on failure and come back.

## 0. Before you start

- [ ] Close any running instance of the app and relaunch `UI-Script.ps1` (new code only loads at launch).
- [ ] Confirm you're on the network where the MTSC / eMARS pages open in your browser (USPS intranet).
- [ ] **Back up** the whole `Work Tracking` folder (especially `Jobs.json`, `Breaks.json`, `Historian.jsonl`, `Weekly Worksheets\*.json`) - the duplication fix is safe, but you're about to stress it.
- [ ] Keep `UI.log` open (tail it in Notepad++ / `Get-Content -Wait`) - every step writes to it.

---

## 1. Sites list - auto-fetch

**Steps:** Settings tab -> **Regenerate Sites List**

- [ ] No "Automatic fetch failed" dialog appears; instead a success box: "Fetched and wrote N entries".
- [ ] `Dropdown CSVs\Sites.csv` has ~290 rows; spot-check the first row is ABILENE, and spot-check a middle row.
- [ ] `UI.log` shows the chain: `Invoke-PmHttpGet: GET http://emarssu3...` -> `200` -> `New-SitesCsv: matched N <option> element(s)` -> `wrote N row(s)`.

**If it fails:** the paste fallback dialog appears - finish it manually so you're not blocked, then send me:
`Debug Dumps\sites-select-*.html` (if present) + a **Ctrl+U raw source** save of the criteria page.

---

## 2. Parts volumes list - auto-fetch

**Steps:** Settings tab -> **Regenerate Parts Volumes List**

- [ ] Success box reports ~96 entries (not 97).
- [ ] `Dropdown CSVs\Parsed-Parts-Volumes.csv` header is `MS Book No,Volume,Full Name`; **no** `all,all,"All MS Handbooks"` row.
- [ ] `UI.log` shows: GET of `index.php` -> then GET of `contents/partsSearch_information.php` -> `yielded a <select>` or `raw <option>s` -> `wrote 96 row(s)`.

**If it fails:** complete the paste fallback, then send me a **Ctrl+U raw source** save of
`https://www1.mtsc.usps.gov/apps/mtsc/contents/partsSearch_information.php`
(and `msbook_information.php` if convenient).

---

## 3. Parts Book Creator - tree auto-fetch (the big one)

**Setup:** pick a volume that is **not already installed** (installed books are greyed out).
If everything's installed, remove one book from the config via Actions -> Remove Parts Book
(files on disk are preserved) and re-create it as the test.

**Steps:** Parts Books tab -> **Create Parts Book** -> check the volume -> **Process Selected**

- [ ] **No browser window opens and no paste dialog appears.** You go straight to the download progress.
- [ ] `UI.log` shows: `Parts Book Creator: auto-fetched tree for <book>` plus
  `Get-PartsBookTreeFromWeb: found N section(s)` and `fetched M figure(s)`.
- [ ] Book folder (`Parts Books\<Book Name>\`) contains:
  - [ ] `Volumes-to-URL.csv` - one row per figure, URLs pointing at `content/printfigandtable.php?...&viewerflag=d&layout=L11`
  - [ ] `SectionNames.txt` - one `Section N <name>` line per section, numeric order
  - [ ] `HTML and CSV Files\Figure <sec>-<fig>.html` / `.csv` for every figure
  - [ ] `CombinedSections\Section N.csv` merged against your parts room
  - [ ] `<Book Name>.xlsx` with one worksheet per section, renamed to full section names
- [ ] Spot-check one figure CSV against the paper handbook (or the browser page) - NSN/part numbers sane.

**If it fails:** it should fall back to the browser + paste flow (your old path still works - complete it). Then capture for me:
- [ ] F12 -> **Network** -> click a section in the tree -> find the **`section_figures.php`** POST -> **Response** tab -> copy all -> save
- [ ] F12 -> **Network** -> reload the book page -> find the **`build_Tree_sections.php`** POST -> **Response** -> copy -> save
- [ ] `Debug Dumps\phbk-tree-*.html` if present

---

## 4. Work order duplication - regression tests

**Setup:** an existing work-order job for today, plus today's worksheet loaded (Sync once first so you start from a known state).

- [ ] **Edit WO number:** double-click a work order -> change its WO # -> Save.
  - [ ] Work Order tab shows exactly one row for it (no immediate duplicate).
- [ ] Go to Weekly Worksheet -> click **Sync**.
  - [ ] Worksheet: the old WO # row is gone; the new WO # appears exactly once.
  - [ ] Work Order tab: **no phantom job with the old WO #** appears.
- [ ] Click **Sync again** - nothing changes (idempotent; job count stable).
- [ ] **Delete a work order** (double-click -> Delete) -> Sync.
  - [ ] The job does **not** reappear in the Work Order tab.
  - [ ] Its worksheet row disappears.
- [ ] **Manual override respected:** double-click a worksheet row -> set Actual time -> Sync -> the value stays (not overwritten).
- [ ] **Fetch from eDAC** (if the page is available) -> re-check: no duplicated rows in either tab.
- [ ] `UI.log`: watch for `pruning stale auto row WO=...` and `Import-WorksheetJobsToJobLog: added N job(s)` - N should only count genuinely new eDAC work orders.

**If any duplicate appears:** grab `UI.log`, the day's `Weekly Worksheets\YYYY-MM-DD.json`, and `Jobs.json`, and send all three.

---

## 5. "The goods" - expected artifacts summary

| Feature | File | Expected |
|---|---|---|
| Sites | `Dropdown CSVs\Sites.csv` | ~290 rows, `Site ID,Full Name` |
| Volumes | `Dropdown CSVs\Parsed-Parts-Volumes.csv` | ~96 rows, no `all` placeholder |
| Book tree | `Parts Books\<Book>\Volumes-to-URL.csv` | 1 row/figure, `printfigandtable.php` URLs |
| Book tree | `Parts Books\<Book>\SectionNames.txt` | `Section N <full name>` lines |
| Book build | `Parts Books\<Book>\<Book>.xlsx` | one sheet per section, renamed |
| Jobs | `Work Tracking\Jobs.json` | no duplicate id / WO#+date pairs |
| Worksheet | `Work Tracking\Weekly Worksheets\YYYY-MM-DD.json` | no duplicate WO rows |

## 6. If anything misbehaves - send me this bundle

- [ ] `UI.log` (whole file)
- [ ] `Debug Dumps\` folder contents (especially `phbk-tree-*`, `sites-select-*`)
- [ ] Raw Ctrl+U saves: parts-search info page, eMARS criteria page
- [ ] F12 Network responses: `section_figures.php`, `build_Tree_sections.php`
- [ ] The affected `Jobs.json` and the day's worksheet JSON
