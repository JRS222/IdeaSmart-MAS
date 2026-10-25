Fixing the timer bug + adding book/section scope

Part 1 — The timer bug

Two separate bugs, both real.

Bug 1a: "Start Timer" fails validation when end time is blank

Test-JobFields requires both startTime and endTime. The Start Timer handler in Show-JobEditor builds a candidate from the current pickers, validates it (passes because the end-time picker defaults to "now"), then the running-state block nulls endTime after validation — so the initial save is fine. But the failure mode is the editor's Start Timer path itself: if you clear the end picker (or the picker ever lands on an invalid value), you get "End time is required" and the timer never starts. And once a running timer's job is later re-validated anywhere (Send-AllReadyJobs, Save), it fails for the same reason.

Fix — make Test-JobFields timer-aware:

```powershell
function Test-JobFields {
    param($Job)

    $errs = @()
    if ($null -eq $Job) { return @('No job to validate.') }

    $m = Format-MachineId "$($Job.machineId)"
    if ($null -eq $m) { $errs += 'Machine ID must look like "DBCS 49".' }

    # A running or paused timer legitimately has no end time yet.
    $timerActive = (Get-JobTimerState -Job $Job) -in @('running','paused')

    if ([string]::IsNullOrWhiteSpace($Job.startTime)) { $errs += 'Start time is required.' }
    if (-not $timerActive -and [string]::IsNullOrWhiteSpace($Job.endTime)) {
        $errs += 'End time is required.'
    }

    $h = Get-JobHours $Job
    if ($null -ne $h -and $h -le 0) { $errs += 'End time must be after start time.' }

    $needsDetail = ($null -ne $h -and $h -gt 0.5) -or (@($Job.parts).Count -gt 0)
    if ($needsDetail) {
        if ([string]::IsNullOrWhiteSpace($Job.description)) { $errs += 'Description required (>30 min or parts used).' }
        if ([string]::IsNullOrWhiteSpace($Job.workOrderNo)) { $errs += 'Work order # required (>30 min or parts used).' }
    }

    return ,$errs
}
```

And in the Start Timer handler, move the running-state setup before the validation call so Test-JobFields sees the running state:

```powershell
$startTimerBtn.Add_Click({
    # ...stop branch unchanged, see fix 1b...

    $candidate = [PSCustomObject]@{ ... }   # unchanged

    # Enter running state FIRST so Test-JobFields skips the end-time requirement
    if ((Get-JobTimerState -Job $Job) -notin @('running','paused','finished')) {
        $now = Get-Date
        $candidate.startTime               = $now.ToString('HH:mm')
        $candidate.endTime                 = $null
        $candidate.timerState              = 'running'
        $candidate.timerStartedAt          = $now.ToString("o")
        $candidate.timerPausedAt           = $null
        $candidate.timerPausedTotalSeconds = 0
        $candidate.timerEndedAt            = $null
        $candidate.timerComputedHours      = $null
    }

    $errs = Test-JobFields -Job $candidate
    if ($errs.Count -gt 0) { $lblVal.Text = "! " + $errs[0]; return }
    $lblVal.Text = ''

    $m = Format-MachineId $candidate.machineId
    if ($null -eq $m) {
        $lblVal.Text = "! Machine ID must look like 'DBCS 49'."
        return
    }
    $candidate.machineId = $m.Canonical
    if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

    $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate; StartTimer = $true }
    $form.Close()
})
```

The silent if ($null -eq $m) { return } was also a bug on its own — the user sees nothing when they fat-finger the machine ID.

Bug 1b: "Stop Timer" wipes start/end and discards unsaved editor edits

The Stop Timer branch in Show-JobEditor does two destructive things at once:

1. Calls Reset-JobTimer, which sets $j.startTime = ''; $j.endTime = '' on the persisted job
2. Returns Action = 'stop'

The Action = 'stop' branch is never handled by the caller — the double-click handler in Setup-JobLogSubTab only knows 'save' and 'delete'. So:

· Reset-JobTimer already committed the start/end wipe to Jobs.json
· The editor's description / WO# / parts edits are thrown away because they never make it back into the caller

Result: the user loses start time, end time, and any field edits they made in the editor. That's the "deletes all the stat" bug.

Fix — build a candidate that preserves the editor fields but resets only the timer, and return Action = 'save':

```powershell
$startTimerBtn.Add_Click({
    if ((Get-JobTimerState -Job $Job) -in @('running','paused')) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Stop this timer?`r`n`r`nThe elapsed time will NOT be recorded.`r`nYour other edits (description, WO#, parts) will be saved.",
            "Stop Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # Reset ONLY the timer fields; preserve everything the user has
        # typed into the editor so Save does not discard it.
        $reset = [PSCustomObject]@{
            id              = $Job.id
            kind            = $Job.kind
            machineId       = $txtMachine.Text.Trim()
            startTime       = ''
            endTime         = ''
            hoursOverride   = $Job.hoursOverride
            description     = $txtDesc.Text.Trim()
            workOrderNo     = $txtWo.Text.Trim()
            parts           = $script:__je_parts
            createdAt       = $Job.createdAt
            sentToHistorian = $Job.sentToHistorian
            sentAt          = $Job.sentAt
            timerState              = 'idle'
            timerStartedAt          = $null
            timerPausedAt           = $null
            timerPausedTotalSeconds = 0
            timerEndedAt            = $null
            timerComputedHours      = $null
        }

        $m = Format-MachineId $reset.machineId
        if ($m) {
            $reset.machineId = $m.Canonical
            if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }
        }

        $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $reset }
        $form.Close()
        return
    }

    # ...start branch from fix 1a...
})
```

Note we no longer call Reset-JobTimer here at all. Reset-JobTimer should stay in place for the timer-form's Cancel/Reset buttons, where a fresh slate really is the intent.

Bug 1c: opening the timer form silently starts a timer

Both Show-JobTimerForm and the list's Timer-column MouseClick handler call Start-JobTimer when the job is idle:

```powershell
if ((Get-JobTimerState -Job $job) -eq 'idle') {
    Start-JobTimer -JobId $jobId
    ...
}
```

That's why you can't "start the timer without changing end time" in the first place — you didn't choose to start it, opening the form did, and it unconditionally overwrites startTime with "now" and nulls endTime, destroying any manually-entered times.

Fix — stop auto-starting; add an explicit Start button to the timer form.

1. Remove the Start-JobTimer block from Show-JobTimerForm (top of the function).
2. Remove the if ((Get-JobTimerState -Job $job) -eq 'idle') { Start-JobTimer ... } block from Setup-JobLogSubTab's MouseClick handler.
3. Add a $startBtn to Show-JobTimerForm:

```powershell
$startBtn = New-Object System.Windows.Forms.Button
$startBtn.Text = "Start Timer"
$startBtn.Size = New-Object System.Drawing.Size(140, 44)
$startBtn.Location = New-Object System.Drawing.Point(180, 246)
$startBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$startBtn.FlatAppearance.BorderSize = 0
$startBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
$startBtn.ForeColor = [System.Drawing.Color]::White
$startBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$startBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$startBtn.Visible = $false
$form.Controls.Add($startBtn)

$startBtn.Add_Click({
    try {
        Start-JobTimer -JobId $JobId
        Refresh-JobLogList
        Refresh-WorkBudget
        & $refreshDisplay
    } catch { Write-Log "Start Timer error: $($_.Exception.Message)" }
}.GetNewClosure())
```

4. Show it only in the idle branch of $refreshDisplay:

```powershell
default {   # idle
    $lblState.Text = "Idle — not started"
    $lblElapsed.Text = "0:00"
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
    $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $false
    $startBtn.Visible = $true
}
```

…and $startBtn.Visible = $false in the running / paused / finished branches.

Now the flow is deliberate: opening the form shows current state; you explicitly click Start when you want to start, and Cancel never nukes your manual times because you never started a timer in the first place.

---

Part 2 — Book+Section scope for Search and Part Picker

Both surfaces get two required dropdowns. The pattern is identical, so here's the shared helper + the deltas for each.

Shared helper

Add this near Get-ConfigBooks:

```powershell
function Get-SectionsForBook {
    # Returns @{ Display; Key; FilePath } for every section of the given book.
    # Key is "Section 5" (the combined-section CSV base name).
    param([string]$BookName)

    $out = @()
    if ([string]::IsNullOrWhiteSpace($BookName)) { return ,$out }
    if (-not $script:config -or -not $script:config.Books) { return ,$out }

    $bookProp = $script:config.Books.PSObject.Properties[$BookName]
    if (-not $bookProp) { return ,$out }

    $volumesCsvPath = $bookProp.Value.VolumesToUrlCsvPath
    $bookDir = if ($volumesCsvPath) { Split-Path -Path $volumesCsvPath -Parent }
               else { Join-Path $script:config.PartsBooksDirectory $BookName }

    $combinedDir = Join-Path $bookDir "CombinedSections"
    if (-not (Test-Path $combinedDir)) { return ,$out }

    $sectionNames = @{}
    $sectionNamesFile = Join-Path $bookDir "SectionNames.txt"
    if (Test-Path $sectionNamesFile) {
        foreach ($line in (Get-Content -Path $sectionNamesFile)) {
            if ($line -match '^Section\s+(\d+)\s+(.*)$') {
                $sectionNames["Section $($matches[1])"] = $line.Trim()
            }
        }
    }

    $files = Get-ChildItem -Path $combinedDir -Filter "Section *.csv" -File |
             Sort-Object -Property @{ Expression = {
                 if ($_.BaseName -match '^Section (\d+)$') { [int]$matches[1] } else { 9999 }
             }}

    foreach ($f in $files) {
        $key = $f.BaseName
        $display = if ($sectionNames.ContainsKey($key)) { $sectionNames[$key] } else { $key }
        $out += [PSCustomObject]@{ Key = $key; Display = $display; FilePath = $f.FullName }
    }
    return ,$out
}

function Get-SearchableBooks {
    # Books that actually have a CombinedSections folder on disk.
    $books = @()
    if (-not $script:config -or -not $script:config.Books) { return ,$books }
    foreach ($bp in $script:config.Books.PSObject.Properties) {
        $volumesCsvPath = $bp.Value.VolumesToUrlCsvPath
        $bookDir = if ($volumesCsvPath) { Split-Path -Path $volumesCsvPath -Parent }
                   else { Join-Path $script:config.PartsBooksDirectory $bp.Name }
        if (Test-Path (Join-Path $bookDir "CombinedSections")) {
            $books += $bp.Name
        }
    }
    return ,$books
}
```

Search tab changes

In Setup-SearchTab, extend the top panel and add the two dropdowns on the right side of the existing three fields. Bump $topPanel.Height from 130 → 130 (fit them beside the current fields, so no height change), and add these after $script:textBoxDescription is created:

```powershell
# --- Book dropdown (right column) ---
$labelBook = New-Object System.Windows.Forms.Label
$labelBook.Text = "Book:"
$labelBook.Location = New-Object System.Drawing.Point(400, 18)
$labelBook.Size = New-Object System.Drawing.Size(80, 22)
$labelBook.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelBook.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
$topPanel.Controls.Add($labelBook)

$script:comboBook = New-Object System.Windows.Forms.ComboBox
$script:comboBook.Location = New-Object System.Drawing.Point(480, 15)
$script:comboBook.Size = New-Object System.Drawing.Size(200, 25)
$script:comboBook.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:comboBook.Font = New-Object System.Drawing.Font("Segoe UI", 9)
foreach ($b in (Get-SearchableBooks)) { [void]$script:comboBook.Items.Add($b) }
$topPanel.Controls.Add($script:comboBook)

# --- Section dropdown (right column) ---
$labelSection = New-Object System.Windows.Forms.Label
$labelSection.Text = "Section:"
$labelSection.Location = New-Object System.Drawing.Point(400, 52)
$labelSection.Size = New-Object System.Drawing.Size(80, 22)
$labelSection.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelSection.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
$topPanel.Controls.Add($labelSection)

$script:comboSection = New-Object System.Windows.Forms.ComboBox
$script:comboSection.Location = New-Object System.Drawing.Point(480, 49)
$script:comboSection.Size = New-Object System.Drawing.Size(200, 25)
$script:comboSection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:comboSection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$topPanel.Controls.Add($script:comboSection)

# Map display -> key so we can translate SelectedItem back to "Section 5"
$script:searchSectionKeyMap = @{}
```

Repopulate sections on book change, and gate the Search button:

```powershell
$script:comboBook.Add_SelectedIndexChanged({
    $script:comboSection.Items.Clear()
    $script:comboSection.SelectedIndex = -1
    $script:searchSectionKeyMap = @{}
    $bookName = "$($script:comboBook.SelectedItem)"
    if ([string]::IsNullOrWhiteSpace($bookName)) { return }
    foreach ($sec in (Get-SectionsForBook -BookName $bookName)) {
        [void]$script:comboSection.Items.Add($sec.Display)
        $script:searchSectionKeyMap[$sec.Display] = $sec.Key
    }
    if ($script:comboSection.Items.Count -gt 0) {
        $script:comboSection.SelectedIndex = 0
    }
})

$updateSearchEnabled = {
    $searchButton.Enabled = (
        $script:comboBook.SelectedIndex    -ge 0 -and
        $script:comboSection.SelectedIndex -ge 0
    )
}.GetNewClosure()

$script:comboBook.Add_SelectedIndexChanged($updateSearchEnabled)
$script:comboSection.Add_SelectedIndexChanged($updateSearchEnabled)
$searchButton.Enabled = $false
```

Then inside the Search button's Add_Click, replace the CrossRef block (foreach ($book in $config.Books.PSObject.Properties) { ... }) with a single-book/single-section scan:

```powershell
# ---- 3) CROSS REFERENCE — scoped to the selected book + section ----
$crossRefResults = @()
$selectedBookName       = "$($script:comboBook.SelectedItem)".Trim()
$selectedSectionDisplay = "$($script:comboSection.SelectedItem)".Trim()
$selectedSectionKey     = $script:searchSectionKeyMap[$selectedSectionDisplay]

if (-not [string]::IsNullOrWhiteSpace($selectedBookName) -and
    -not [string]::IsNullOrWhiteSpace($selectedSectionKey)) {

    $secs = Get-SectionsForBook -BookName $selectedBookName
    $match = $secs | Where-Object { $_.Key -eq $selectedSectionKey } | Select-Object -First 1
    if ($match) {
        try {
            $sectionData = Import-Csv -Path $match.FilePath
        } catch {
            Write-Log "CrossRef: could not read $($match.FilePath): $($_.Exception.Message)"
            $sectionData = @()
        }

        if ($nsnNoMatch -or $oemNoMatch) {
            $filtered = @()
        } else {
            $filtered = @($sectionData | Where-Object {
                $nsnOk = $true
                if ($nsnActive) {
                    $rowStock = ([string]$_.'STOCK NO.') -replace '[^0-9]', ''
                    $nsnOk = $rowStock -match $nsnPattern
                }
                $oemOk = $true
                if ($oemActive) {
                    $rowPart = ([string]$_.'PART NO.' -replace '[^A-Za-z0-9]', '').ToUpper()
                    $oemOk = ($rowPart -and ($rowPart -match $oemPattern))
                }
                $descOk = $true
                if ($descActive) {
                    $rowDesc = (([string]$_.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                    foreach ($token in $descTokens) {
                        if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                    }
                }
                $nsnOk -and $oemOk -and $descOk
            })
        }

        foreach ($item in @($filtered)) {
            $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $selectedBookName  -Force
            $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $selectedSectionDisplay -Force
            $crossRefResults += $item
        }
    }
} else {
    Write-Log "CrossRef skipped: no book/section selected."
}
Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference."
```

The remaining CrossRef populating code (building the ListView items) is unchanged — it just sees fewer rows now.

Availability and Same Day still populate as before — they aren't in a book so book+section scope is meaningless for them, and it would be annoying to hide a technician's own room behind a dropdown.

Part Picker changes

Show-PartPicker gets the same two dropdowns in its search panel. Layout: extend $searchPanel from 70 → 130 and move the results/chosen areas down by 60px.

```powershell
# --- Book dropdown ---
$lblBook = New-Object System.Windows.Forms.Label
$lblBook.Text = "Book:"
$lblBook.Location = New-Object System.Drawing.Point(14, 46)
$lblBook.Size = New-Object System.Drawing.Size(80, 22)
$lblBook.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchPanel.Controls.Add($lblBook)

$cmbBook = New-Object System.Windows.Forms.ComboBox
$cmbBook.Location = New-Object System.Drawing.Point(96, 44)
$cmbBook.Size = New-Object System.Drawing.Size(180, 24)
$cmbBook.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$searchPanel.Controls.Add($cmbBook)
foreach ($b in (Get-SearchableBooks)) { [void]$cmbBook.Items.Add($b) }

# --- Section dropdown ---
$lblSection = New-Object System.Windows.Forms.Label
$lblSection.Text = "Section:"
$lblSection.Location = New-Object System.Drawing.Point(290, 46)
$lblSection.Size = New-Object System.Drawing.Size(84, 22)
$lblSection.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchPanel.Controls.Add($lblSection)

$cmbSection = New-Object System.Windows.Forms.ComboBox
$cmbSection.Location = New-Object System.Drawing.Point(376, 44)
$cmbSection.Size = New-Object System.Drawing.Size(260, 24)
$cmbSection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$searchPanel.Controls.Add($cmbSection)

$ppSectionKeyMap = @{}
$cmbBook.Add_SelectedIndexChanged({
    $cmbSection.Items.Clear()
    $cmbSection.SelectedIndex = -1
    $script:ppSectionKeyMap = @{}
    $b = "$($cmbBook.SelectedItem)"
    if ([string]::IsNullOrWhiteSpace($b)) { return }
    foreach ($sec in (Get-SectionsForBook -BookName $b)) {
        [void]$cmbSection.Items.Add($sec.Display)
        $script:ppSectionKeyMap[$sec.Display] = $sec.Key
    }
    if ($cmbSection.Items.Count -gt 0) { $cmbSection.SelectedIndex = 0 }
}.GetNewClosure())
```

Move the Search button down to (650, 44) so it aligns with the Book/Section row. Leave the NSN/Description textboxes where they were.

Then extend Search-Parts to accept optional book+section scope and include CrossRef matches:

```powershell
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description,
        [string]$BookFilter = '',
        [string]$SectionFilter = ''
    )

    $results = @()

    # ...existing NSN pattern / descTokens / rowTester code unchanged...

    # 1. Local Parts Room (unchanged)
    # 2. Same Day Parts Room (unchanged)

    # 3. Scoped Cross Reference (new)
    if (-not [string]::IsNullOrWhiteSpace($BookFilter) -and
        -not [string]::IsNullOrWhiteSpace($SectionFilter)) {

        $crossTester = {
            param($row)
            $nsnOk = $true
            if ($nsnPattern) {
                $rowNSN = ([string]$row.'STOCK NO.') -replace '[^0-9]', ''
                $nsnOk = $rowNSN -match $nsnPattern
            }
            $descOk = $true
            if ($descTokens.Count -gt 0) {
                $rowDesc = (([string]$row.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                foreach ($t in $descTokens) {
                    if (-not $rowDesc.Contains($t)) { $descOk = $false; break }
                }
            }
            $nsnOk -and $descOk
        }

        $secs = Get-SectionsForBook -BookName $BookFilter
        $match = $secs | Where-Object { $_.Key -eq $SectionFilter } | Select-Object -First 1
        if ($match) {
            try { $sectionData = Import-Csv -Path $match.FilePath }
            catch { $sectionData = @() }

            foreach ($row in @($sectionData)) {
                if (-not (& $crossTester $row)) { continue }
                $qty = 0; [void][int]::TryParse("$($row.QTY)", [ref]$qty)
                $results += [PSCustomObject]@{
                    PartNumber  = "$($row.'STOCK NO.')"
                    Description = "$($row.'PART DESCRIPTION')"
                    Quantity    = $qty
                    Location    = "$($row.Location)"
                    OEMNumber   = "$($row.'PART NO.')"
                    Source      = "Book: $BookFilter / $SectionFilter"
                }
            }
        }
    }

    return $results
}
```

And in the picker's $searchBtn.Add_Click:

```powershell
$bookName       = "$($cmbBook.SelectedItem)".Trim()
$sectionDisplay = "$($cmbSection.SelectedItem)".Trim()
$sectionKey     = $script:ppSectionKeyMap[$sectionDisplay]

if ([string]::IsNullOrWhiteSpace($bookName) -or [string]::IsNullOrWhiteSpace($sectionKey)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Select a Parts Book and a Section first.",
        "Add Parts", "OK", "Information")
    return
}

$found = Search-Parts -NSN $nsn -Description $dsc -BookFilter $bookName -SectionFilter $sectionKey
```

Then gate the picker's Search button the same way as the Search tab ($searchBtn.Enabled = ($cmbBook.SelectedIndex -ge 0 -and $cmbSection.SelectedIndex -ge 0)).

Reference parts land in the picker with Source = "Book: <book> / Section 5" and their QTY/Location already populated by the merge step, so picking one flows into the job's parts list without any further changes.

---

Summary of what changed

· Test-JobFields — end time no longer required when timer state is running/paused
· Show-JobEditor — Start Timer — running state set before validation; silent invalid-machine-ID return fixed
· Show-JobEditor — Stop Timer — reset candidate preserves description/WO#/parts; returns Action='save' instead of falling off the edge as Action='stop'
· Show-JobTimerForm + Setup-JobLogSubTab timer-column click — removed auto-start; added explicit Start Timer button on the timer form
· Search tab — book+section dropdowns (required to enable Search); CrossRef scan scoped to one file instead of every book's every section
· Part Picker — same two required dropdowns; Search-Parts now optionally includes CrossRef parts from the chosen section

The performance side-benefit of the book/section scope is probably the biggest practical win here — Search-CrossReferenceData was iterating every CombinedSections/*.csv across every book; scoping to a single file cuts the read-everything cost to a single Import-Csv per search.