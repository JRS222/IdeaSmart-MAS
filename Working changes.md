Combined Change Checklist

Grouped into A. Bug fixes, B. Book/Section scoping, C. eMARS user list (caveated). Each item lists where to find the block, then old → new. New-only items say where to add.

---

A. Bug fixes

A1. Test eMARS button wipes the site ID

Find: $emarsTestBtn.Add_Click({ — the first line inside Setup-SettingsTab that follows $emarsTestBtn creation.

Why: the handler references $emarsSiteBox, $emarsUserBox, etc. — function locals that are gone by the time the click fires. In PS 5.1 $null.Text is "", so Save gets EmarsSite: "" and Test says "not set".

Old:

```powershell
$emarsTestBtn.Add_Click({
    try {
        # Apply the current boxes first (same as Save) so the test
        # always uses what the user is looking at right now.
        $cfg = $script:config
        $cfg.WorkTracking.EmarsOrigin  = "$($emarsOriginBox.Text)".Trim()
        $cfg.WorkTracking.EmarsUser    = "$($emarsUserBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSite    = "$($emarsSiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSubsite = "$($emarsSubsiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsFacId   = "$($emarsFacBox.Text)".Trim()
        ($cfg | ConvertTo-Json -Depth 12) | Set-Content -Path $script:configPath -Encoding UTF8
        $script:config = $cfg
        Write-Log "Test eMARS Connection: values applied and saved; running endpoint checks..."

        $results = Test-EmarsConnection
        foreach ($line in $results) { Write-Log "Test eMARS Connection: $line" }
        [System.Windows.Forms.MessageBox]::Show(
            ("eMARS connectivity check:`r`n`r`n" + ($results -join "`r`n") + "`r`n`r`n(Details are also in UI.log)"),
            "eMARS Test", "OK", "Information")
    } catch {
        Write-Log "Test eMARS Connection failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Test failed: $($_.Exception.Message)", "eMARS Test", "OK", "Error")
    }
})
```

New:

```powershell
$emarsTestBtn.Add_Click({
    try {
        $c = $script:wtSettings
        if (-not $c) { throw "Settings controls are not initialized yet." }

        $cfg = $script:config
        $cfg.WorkTracking.EmarsOrigin  = "$($c.EmarsOriginBox.Text)".Trim()
        $cfg.WorkTracking.EmarsUser    = "$($c.EmarsUserBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSite    = "$($c.EmarsSiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSubsite = "$($c.EmarsSubsiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsFacId   = "$($c.EmarsFacBox.Text)".Trim()
        ($cfg | ConvertTo-Json -Depth 12) | Set-Content -Path $script:configPath -Encoding UTF8
        $script:config = $cfg

        Write-Log "Test eMARS Connection: values applied and saved; running endpoint checks..."
        $results = Test-EmarsConnection
        foreach ($line in $results) { Write-Log "Test eMARS Connection: $line" }
        [System.Windows.Forms.MessageBox]::Show(
            ("eMARS connectivity check:`r`n`r`n" + ($results -join "`r`n") + "`r`n`r`n(Details are also in UI.log)"),
            "eMARS Test", "OK", "Information")
    } catch {
        Write-Log "Test eMARS Connection failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Test failed: $($_.Exception.Message)", "eMARS Test", "OK", "Error")
    }
})
```

Note: no auto-derivation of EmarsSite or TechnicianName from the previous thread. Different eDAC vs eMARS user IDs make the technician-name fallback wrong, and the site-ID fallback is redundant with the fixed Test button.

---

A2. Test-JobFields demands an end time on running timers

Find: function Test-JobFields {

Old:

```powershell
function Test-JobFields {
    # Returns a string[] of validation errors (empty array = valid).
    param($Job)

    $errs = @()
    if ($null -eq $Job) { return @('No job to validate.') }

    $m = Format-MachineId "$($Job.machineId)"
    if ($null -eq $m) { $errs += 'Machine ID must look like "DBCS 49".' }

    if ([string]::IsNullOrWhiteSpace($Job.startTime)) { $errs += 'Start time is required.' }
    if ([string]::IsNullOrWhiteSpace($Job.endTime))   { $errs += 'End time is required.' }

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

New:

```powershell
function Test-JobFields {
    # Returns a string[] of validation errors (empty array = valid).
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

---

A3. Start Timer validates before entering running state

Find: $startTimerBtn.Add_Click({ inside Show-JobEditor. Also fixes the silent if ($null -eq $m) { return } (invalid machine ID currently shows nothing to the user) and A4 (Stop branch discards all other editor edits).

Old:

```powershell
$startTimerBtn.Add_Click({
    # Toggle: when the job's timer is already active, this button STOPS it
    # (so a running timer can always be stopped from the editor popup).
    if ((Get-JobTimerState -Job $Job) -in @('running','paused')) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Stop this timer?`r`n`r`nElapsed time will NOT be recorded — the timer resets to idle.",
            "Stop Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Reset-JobTimer -JobId "$($Job.id)"
        $form.Tag = [PSCustomObject]@{ Action = 'stop' }
        $form.Close()
        return
    }

    $candidate = [PSCustomObject]@{
        id              = $Job.id
        kind            = $Job.kind
        machineId       = $txtMachine.Text.Trim()
        startTime       = $dtpStart.Value.ToString('HH:mm')
        endTime         = $dtpEnd.Value.ToString('HH:mm')
        hoursOverride   = $Job.hoursOverride
        description     = $txtDesc.Text.Trim()
        workOrderNo     = $txtWo.Text.Trim()
        parts           = $script:__je_parts
        createdAt       = $Job.createdAt
        sentToHistorian = $Job.sentToHistorian
        sentAt          = $Job.sentAt
        timerState              = if ($Job.PSObject.Properties['timerState'])              { $Job.timerState }              else { 'idle' }
        timerStartedAt          = if ($Job.PSObject.Properties['timerStartedAt'])          { $Job.timerStartedAt }          else { $null }
        timerPausedAt           = if ($Job.PSObject.Properties['timerPausedAt'])           { $Job.timerPausedAt }           else { $null }
        timerPausedTotalSeconds = if ($Job.PSObject.Properties['timerPausedTotalSeconds']) { $Job.timerPausedTotalSeconds } else { 0 }
        timerEndedAt            = if ($Job.PSObject.Properties['timerEndedAt'])            { $Job.timerEndedAt }            else { $null }
        timerComputedHours      = if ($Job.PSObject.Properties['timerComputedHours'])      { $Job.timerComputedHours }      else { $null }
    }

    $errs = Test-JobFields -Job $candidate
    if ($errs.Count -gt 0) { $lblVal.Text = "! " + $errs[0]; return }
    $lblVal.Text = ''

    $m = Format-MachineId $candidate.machineId
    $candidate.machineId = $m.Canonical
    if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

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

    $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate; StartTimer = $true }
    $form.Close()
})
```

New:

```powershell
$startTimerBtn.Add_Click({
    # ---- Stop branch: reset ONLY the timer, keep the user's other edits ----
    if ((Get-JobTimerState -Job $Job) -in @('running','paused')) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Stop this timer?`r`n`r`nThe elapsed time will NOT be recorded.`r`nYour other edits (description, WO#, parts) will be saved.",
            "Stop Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

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

    # ---- Start branch: enter running state FIRST, then validate ----
    $candidate = [PSCustomObject]@{
        id              = $Job.id
        kind            = $Job.kind
        machineId       = $txtMachine.Text.Trim()
        startTime       = $dtpStart.Value.ToString('HH:mm')
        endTime         = $dtpEnd.Value.ToString('HH:mm')
        hoursOverride   = $Job.hoursOverride
        description     = $txtDesc.Text.Trim()
        workOrderNo     = $txtWo.Text.Trim()
        parts           = $script:__je_parts
        createdAt       = $Job.createdAt
        sentToHistorian = $Job.sentToHistorian
        sentAt          = $Job.sentAt
        timerState              = if ($Job.PSObject.Properties['timerState'])              { $Job.timerState }              else { 'idle' }
        timerStartedAt          = if ($Job.PSObject.Properties['timerStartedAt'])          { $Job.timerStartedAt }          else { $null }
        timerPausedAt           = if ($Job.PSObject.Properties['timerPausedAt'])           { $Job.timerPausedAt }           else { $null }
        timerPausedTotalSeconds = if ($Job.PSObject.Properties['timerPausedTotalSeconds']) { $Job.timerPausedTotalSeconds } else { 0 }
        timerEndedAt            = if ($Job.PSObject.Properties['timerEndedAt'])            { $Job.timerEndedAt }            else { $null }
        timerComputedHours      = if ($Job.PSObject.Properties['timerComputedHours'])      { $Job.timerComputedHours }      else { $null }
    }

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

---

A4. Show-JobTimerForm auto-starts the timer

Find: # Opening the window always means the timer should be live inside Show-JobTimerForm.

Old:

```powershell
$job = Get-JobById -JobId $JobId
if ($null -eq $job) { return }
# Opening the window always means the timer should be live (same behavior
# as the PM task timer). If it is idle, start it now so the Pause / Finish
# / Cancel controls are all available.
if ((Get-JobTimerState -Job $job) -eq 'idle') {
    Start-JobTimer -JobId $JobId
    $job = Get-JobById -JobId $JobId
}
$kindLabel = if ("$($job.kind)" -eq 'workorder') { 'Work Order' } else { 'Reactive Call' }
```

New:

```powershell
$job = Get-JobById -JobId $JobId
if ($null -eq $job) { return }
# Do NOT auto-start. The user clicks Start Timer inside the form.
$kindLabel = if ("$($job.kind)" -eq 'workorder') { 'Work Order' } else { 'Reactive Call' }
```

---

A5. Timer-column click in Setup-JobLogSubTab also auto-starts

Find: if ((Get-JobTimerState -Job $job) -eq 'idle') { — the one directly above Write-Log "Job timer column clicked.

Old:

```powershell
if ((Get-JobTimerState -Job $job) -eq 'idle') {
    Start-JobTimer -JobId $jobId
    Refresh-WorkBudget
}
Write-Log "Job timer column clicked: id=$jobId"
Show-JobTimerForm -JobId $jobId
```

New:

```powershell
Write-Log "Job timer column clicked: id=$jobId"
Show-JobTimerForm -JobId $jobId
```

---

A6. Add explicit Start Timer button to Show-JobTimerForm

Add after the $cancelTimerBtn block (the line $form.Controls.Add($cancelTimerBtn) is the anchor — insert immediately after).

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

Then in $refreshDisplay, the default (idle) branch — find the line $lblState.Text = "Idle"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145):

Old:

```powershell
default {
    $lblState.Text = "Idle"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
    $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $false
}
```

New:

```powershell
default {
    $lblState.Text = "Idle — not started"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
    $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $false
    $startBtn.Visible = $true
}
```

Also add $startBtn.Visible = $false to each of the 'running', 'paused', and 'finished' branches of that same switch (they already set the other four button visibilities).

---

B. Book + Section scoping

B1. Add two helper functions

Insert immediately before function Search-Parts {.

```powershell
function Get-SectionsForBook {
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

---

B2. Search tab — add book + section dropdowns

Insert immediately after $topPanel.Controls.Add($script:textBoxDescription) inside Setup-SearchTab. Layout: three existing fields stay at X=15–360; Search button stays at X=380–500; the two dropdowns go to the right at X=520–780 — no downstream repositioning needed.

```powershell
# --- Book dropdown ---
$labelBook = New-Object System.Windows.Forms.Label
$labelBook.Text = "Book:"
$labelBook.Location = New-Object System.Drawing.Point(520, 18)
$labelBook.Size = New-Object System.Drawing.Size(60, 22)
$labelBook.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelBook.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
$topPanel.Controls.Add($labelBook)

$script:comboBook = New-Object System.Windows.Forms.ComboBox
$script:comboBook.Location = New-Object System.Drawing.Point(590, 15)
$script:comboBook.Size = New-Object System.Drawing.Size(190, 25)
$script:comboBook.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:comboBook.Font = New-Object System.Drawing.Font("Segoe UI", 9)
foreach ($b in (Get-SearchableBooks)) { [void]$script:comboBook.Items.Add($b) }
$topPanel.Controls.Add($script:comboBook)

# --- Section dropdown ---
$labelSection = New-Object System.Windows.Forms.Label
$labelSection.Text = "Section:"
$labelSection.Location = New-Object System.Drawing.Point(520, 52)
$labelSection.Size = New-Object System.Drawing.Size(60, 22)
$labelSection.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelSection.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
$topPanel.Controls.Add($labelSection)

$script:comboSection = New-Object System.Windows.Forms.ComboBox
$script:comboSection.Location = New-Object System.Drawing.Point(590, 49)
$script:comboSection.Size = New-Object System.Drawing.Size(190, 25)
$script:comboSection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:comboSection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$topPanel.Controls.Add($script:comboSection)

$script:searchSectionKeyMap = @{}

# Populate sections on book change.
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
```

---

B3. Search tab — gate the Search button on both dropdowns

Insert immediately after the block added in B2. (.GetNewClosure() is required because $searchButton is a function-local; without it the handler resolves to $null at click time — same class of bug as A1.)

```powershell
$updateSearchEnabled = {
    $searchButton.Enabled = (
        $script:comboBook.SelectedIndex    -ge 0 -and
        $script:comboSection.SelectedIndex -ge 0
    )
}.GetNewClosure()

$script:comboBook.Add_SelectedIndexChanged($updateSearchEnabled)
$script:comboSection.Add_SelectedIndexChanged($updateSearchEnabled)

# Pre-select first book so the tab is usable as soon as it opens.
if ($script:comboBook.Items.Count -gt 0) {
    $script:comboBook.SelectedIndex = 0
}
$searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -and $script:comboSection.SelectedIndex -ge 0)
```

---

B4. Search tab — replace the CrossRef scan with a scoped single-file scan

Find: # ---- 3) CROSS REFERENCE ---- inside the $searchButton.Add_Click handler. Replace everything from that comment line down through Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference.".

⚠️ Bonus bug fix included: the original line -NotePropertyName 'Section Name' -NotePropertyName $sectionName uses -NotePropertyName twice, which is why the Section Name column in the CrossRef ListView was likely blank/wrong. The new code uses -NotePropertyValue.

Old:

```powershell
        # ---- 3) CROSS REFERENCE ----
        $crossRefResults = @()
        if ($config.Books) {
            foreach ($book in $config.Books.PSObject.Properties) {
                $bookName = $book.Name
                $volumesCsvPath = $book.Value.VolumesToUrlCsvPath
                if ($volumesCsvPath) {
                    $bookDir = Split-Path -Path $volumesCsvPath -Parent
                } else {
                    $bookDir = Join-Path $config.PartsBooksDirectory $bookName
                }

                $combinedSectionsDir = Join-Path $bookDir "CombinedSections"
                $sectionNamesFile    = Join-Path $bookDir "SectionNames.txt"

                $sectionNameMapping = @{}
                if (Test-Path $sectionNamesFile) {
                    $sectionNames = Get-Content -Path $sectionNamesFile
                    foreach ($line in $sectionNames) {
                        if ($line -match '^Section\s+(\d+)\s+(.*)$') {
                            $sectionNumber   = $Matches[1]
                            $sectionFullName = $line.Trim()
                            $sectionNameMapping["Section $sectionNumber"] = $sectionFullName
                        }
                    }
                    Write-Log "Loaded $($sectionNameMapping.Count) section names for $bookName"
                } else {
                    Write-Log "SectionNames.txt not found for $bookName"
                }

                if (-not (Test-Path $combinedSectionsDir)) {
                    Write-Log "CombinedSections directory not found for $bookName"
                    continue
                }

                $sectionCsvFiles = Get-ChildItem -Path $combinedSectionsDir -Filter "*.csv" -File
                Write-Log "Found $($sectionCsvFiles.Count) CSV files in $bookName"

                foreach ($csvFile in $sectionCsvFiles) {
                    $sectionFileName = [IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
                    $csvFilePath = $csvFile.FullName

                    if ($sectionFileName -match '^Section\s+(\d+)$') {
                        $sectionNumber = $Matches[1]
                        if ($sectionNameMapping.ContainsKey("Section $sectionNumber")) {
                            $sectionName = $sectionNameMapping["Section $sectionNumber"]
                        } else {
                            $sectionName = "Section $sectionNumber"
                        }
                    } else {
                        $sectionName = $sectionFileName
                    }

                    try {
                        $sectionData = Import-Csv -Path $csvFilePath
                        Write-Log "Processed $($sectionData.Count) rows from $($csvFile.Name)"
                    } catch {
                        Write-Log "Failed to read CSV file $csvFilePath. Error: $_"
                        continue
                    }

                    if ($nsnNoMatch -or $oemNoMatch) {
                        $filteredSectionData = @()
                    } else {
                        $filteredSectionData = @($sectionData | Where-Object {
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

                    foreach ($item in $filteredSectionData) {
                        $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $bookName    -Force
                        $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyName $sectionName -Force
                        $crossRefResults += $item
                    }
                }
            }
        } else {
            Write-Log "No books defined in configuration."
        }

        Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference."
```

New:

```powershell
        # ---- 3) CROSS REFERENCE — scoped to selected book + section ----
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
                    $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $selectedBookName       -Force
                    $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $selectedSectionDisplay -Force
                    $crossRefResults += $item
                }
            }
        } else {
            Write-Log "CrossRef skipped: no book/section selected."
        }
        Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference."
```

The CrossRef ListView population block that follows is unchanged — it just iterates $crossRefResults.

---

B5. Extend Search-Parts with BookFilter / SectionFilter

Find: function Search-Parts {. Replace only the param and add the CrossRef scan before the return $results.

Old (top of function):

```powershell
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description
    )
```

New:

```powershell
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description,
        [string]$BookFilter = '',
        [string]$SectionFilter = ''
    )
```

Old (end of function, before return $results):

```powershell
    # 2. Same Day Parts Room
    $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    if (Test-Path $sameDayDir) {
        foreach ($f in (Get-ChildItem -Path (Join-Path $sameDayDir "*.csv") -File)) {
            $siteName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            foreach ($row in (Import-Csv -Path $f.FullName)) {
                if ([int]$row.QTY -le 0) { continue }
                if (-not (& $rowTester $row)) { continue }
                $results += [PSCustomObject]@{
                    PartNumber = $row.'Part (NSN)'
                    Description = $row.Description
                    Quantity = [int]$row.QTY
                    Location = $row.Location
                    OEMNumber = $row.'OEM 1'
                    Source = "$siteName (Same Day)"
                }
            }
        }
    }

    return $results
}
```

New (insert the CrossRef block between the Same Day block and return $results):

```powershell
    # 2. Same Day Parts Room
    $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    if (Test-Path $sameDayDir) {
        foreach ($f in (Get-ChildItem -Path (Join-Path $sameDayDir "*.csv") -File)) {
            $siteName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            foreach ($row in (Import-Csv -Path $f.FullName)) {
                if ([int]$row.QTY -le 0) { continue }
                if (-not (& $rowTester $row)) { continue }
                $results += [PSCustomObject]@{
                    PartNumber = $row.'Part (NSN)'
                    Description = $row.Description
                    Quantity = [int]$row.QTY
                    Location = $row.Location
                    OEMNumber = $row.'OEM 1'
                    Source = "$siteName (Same Day)"
                }
            }
        }
    }

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

---

B6. Part Picker — extend search panel + add dropdowns

Find: $searchPanel.Size = New-Object System.Drawing.Size(850, 70) inside Show-PartPicker.

Old:

```powershell
$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Location = New-Object System.Drawing.Point(10, 10)
$searchPanel.Size = New-Object System.Drawing.Size(850, 70)
$searchPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($searchPanel)
```

New:

```powershell
$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Location = New-Object System.Drawing.Point(10, 10)
$searchPanel.Size = New-Object System.Drawing.Size(850, 130)
$searchPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($searchPanel)
```

Insert (after $searchPanel.Controls.Add($searchBtn) — the Search button add):

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

# Search button moves down to the second row
$searchBtn.Location = New-Object System.Drawing.Point(650, 44)

$script:ppSectionKeyMap = @{}
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
    $searchBtn.Enabled = ($cmbBook.SelectedIndex -ge 0 -and $cmbSection.SelectedIndex -ge 0)
}.GetNewClosure())

$cmbSection.Add_SelectedIndexChanged({
    $searchBtn.Enabled = ($cmbBook.SelectedIndex -ge 0 -and $cmbSection.SelectedIndex -ge 0)
}.GetNewClosure())
```

Then shift every control below the panel down by 60. Find and update these four Location lines:

Line Old New
$lblResults.Location = New-Object System.Drawing.Point(14, 88) (14, 88) (14, 148)
$lvResults.Location = New-Object System.Drawing.Point(14, 110) (14, 110) (14, 170)
$addSelBtn.Location = New-Object System.Drawing.Point(14, 338) (14, 338) (14, 398)
$rmChosenBtn.Location = New-Object System.Drawing.Point(170, 338) (170, 338) (170, 398)
$lblChosen.Location = New-Object System.Drawing.Point(14, 378) (14, 378) (14, 438)
$lvChosen.Location = New-Object System.Drawing.Point(14, 400) (14, 400) (14, 460)

And bump the form height so the chosen list still fits:

Old:

```powershell
$form.Size = New-Object System.Drawing.Size(880, 700)
$form.MinimumSize = New-Object System.Drawing.Size(880, 700)
```

New:

```powershell
$form.Size = New-Object System.Drawing.Size(880, 760)
$form.MinimumSize = New-Object System.Drawing.Size(880, 760)
```

---

B7. Part Picker — pass the filters and gate the Search button

Find: $searchBtn.Add_Click({ inside Show-PartPicker.

Old:

```powershell
$searchBtn.Add_Click({
    $lvResults.Items.Clear()
    $nsn = $txtNSN.Text.Trim()
    $dsc = $txtDesc.Text.Trim()
    try {
        $found = Search-Parts -NSN $nsn -Description $dsc
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Search failed: $($_.Exception.Message)", "Error", "OK", "Error")
        return
    }
    foreach ($part in @($found)) {
        $it = New-Object System.Windows.Forms.ListViewItem("$($part.PartNumber)")
        $it.SubItems.Add("$($part.Description)") | Out-Null
        $it.SubItems.Add("$($part.Quantity)")    | Out-Null
        $it.SubItems.Add("$($part.Location)")    | Out-Null
        $it.SubItems.Add("$($part.OEMNumber)")   | Out-Null
        $it.SubItems.Add("$($part.Source)")      | Out-Null
        $it.Tag = $part
        $lvResults.Items.Add($it) | Out-Null
    }
})
```

New:

```powershell
$searchBtn.Add_Click({
    $bookName       = "$($cmbBook.SelectedItem)".Trim()
    $sectionDisplay = "$($cmbSection.SelectedItem)".Trim()
    $sectionKey     = $script:ppSectionKeyMap[$sectionDisplay]

    if ([string]::IsNullOrWhiteSpace($bookName) -or [string]::IsNullOrWhiteSpace($sectionKey)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select a Parts Book and a Section first.",
            "Add Parts", "OK", "Information")
        return
    }

    $lvResults.Items.Clear()
    $nsn = $txtNSN.Text.Trim()
    $dsc = $txtDesc.Text.Trim()
    try {
        $found = Search-Parts -NSN $nsn -Description $dsc -BookFilter $bookName -SectionFilter $sectionKey
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Search failed: $($_.Exception.Message)", "Error", "OK", "Error")
        return
    }
    foreach ($part in @($found)) {
        $it = New-Object System.Windows.Forms.ListViewItem("$($part.PartNumber)")
        $it.SubItems.Add("$($part.Description)") | Out-Null
        $it.SubItems.Add("$($part.Quantity)")    | Out-Null
        $it.SubItems.Add("$($part.Location)")    | Out-Null
        $it.SubItems.Add("$($part.OEMNumber)")   | Out-Null
        $it.SubItems.Add("$($part.Source)")      | Out-Null
        $it.Tag = $part
        $lvResults.Items.Add($it) | Out-Null
    }
}.GetNewClosure())
```

Also add the initial Enabled state right after the two Add_SelectedIndexChanged handlers from B6:

```powershell
$searchBtn.Enabled = ($cmbBook.SelectedIndex -ge 0 -and $cmbSection.SelectedIndex -ge 0)
if ($cmbBook.Items.Count -gt 0) { $cmbBook.SelectedIndex = 0 }   # triggers section population + enable
```

---

C. eMARS user list dropdown — caveated

C1. Verify the endpoint first (manual, ~5 min)

1. Open eCBM in Chrome/Edge → F12 → Network → Fetch/XHR.
2. Navigate to any user-list / technician-list / sign-off screen.
3. Look for ecbm.*.php requests returning XML — candidates likely contain user, employee, person, admin, or security.
4. Copy the URL and its query params.
5. Confirm the technician's own login can read it (not just admin). If the request 403s or returns a login page for non-admin accounts, use the fallback in C3 — a dropdown that only works for supervisors is worse than a text box.

C2. Path 1 — endpoint works: add Get-EmarsUserList + settings combo

Insert alongside the existing eMARS helpers (e.g. immediately after function Get-EmarsMachineList). Fill in the Proc / Query / field names from C1:

```powershell
function Get-EmarsUserList {
    param([string]$SiteId = '')

    if ([string]::IsNullOrWhiteSpace($SiteId)) { $SiteId = Get-EmarsSetting -Name 'EmarsSite' }
    if ([string]::IsNullOrWhiteSpace($SiteId)) { throw "eMARS Site ID is not set." }

    # TODO: replace Proc and Query with the values observed in DevTools.
    $xml = Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_user.getUsers' `
        -Query "p_site=$SiteId" `
        -Label "getUsers(site $SiteId)"

    $users = @()
    foreach ($row in (Get-EmarsRows -Doc $xml)) {
        $userId = Get-EmarsField -Row $row -Names @('USERID','USER_ID','P_USER') -Index 0
        $name   = Get-EmarsField -Row $row -Names @('USERNAME','FULLNAME','NAME') -Index 1
        if ([string]::IsNullOrWhiteSpace($userId)) { continue }
        $users += [PSCustomObject]@{
            UserId = $userId
            Name   = if ($name) { $name } else { $userId }
        }
    }
    return ,$users
}
```

Then in Setup-SettingsTab, replace $emarsUserBox (the New-Object TextBox for the eMARS User ID) with a ComboBox. Keep the variable name so the Save handler stays untouched — the ComboBox has .Text too.

Old:

```powershell
$emarsUserBox = New-Object System.Windows.Forms.TextBox
$emarsUserBox.Size = New-Object System.Drawing.Size(440, 24)
New-SettingRow -Parent $gEmars -Top 62 -LabelText "eMARS User ID (p_user):" -InputControl $emarsUserBox
```

New:

```powershell
$emarsUserBox = New-Object System.Windows.Forms.ComboBox
$emarsUserBox.Size = New-Object System.Drawing.Size(440, 24)
$emarsUserBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
New-SettingRow -Parent $gEmars -Top 62 -LabelText "eMARS User ID (p_user):" -InputControl $emarsUserBox

# Refresh the list whenever the site box changes.
# (Attach after $emarsSiteBox is created and after $script:wtSettings is set.)
$refreshEmarsUsers = {
    try {
        $script:wtSettings.EmarsUserBox.Items.Clear()
        $site = "$($script:wtSettings.EmarsSiteBox.Text)".Trim()
        if ([string]::IsNullOrWhiteSpace($site)) { return }
        foreach ($u in (Get-EmarsUserList -SiteId $site)) {
            [void]$script:wtSettings.EmarsUserBox.Items.Add("$($u.UserId)  —  $($u.Name)")
        }
        # Preserve whatever the config already stored, if it appears in the list.
        $saved = "$($script:config.WorkTracking.EmarsUser)".Trim()
        if ($saved) {
            for ($i = 0; $i -lt $script:wtSettings.EmarsUserBox.Items.Count; $i++) {
                if ("$($script:wtSettings.EmarsUserBox.Items[$i])".StartsWith("$saved ")) {
                    $script:wtSettings.EmarsUserBox.SelectedIndex = $i
                    break
                }
            }
        }
    } catch {
        Write-Log "Get-EmarsUserList failed: $($_.Exception.Message)"
    }
}.GetNewClosure()
```

Then, at the very end of Setup-SettingsTab (after Load-WtSettingsValues), add:

```powershell
$script:wtSettings.EmarsSiteBox.Add_TextChanged($refreshEmarsUsers)
& $refreshEmarsUsers
```

And in the Save handler, parse the selected item back to a bare user ID before writing:

Old:

```powershell
$emarsUser    = "$($c.EmarsUserBox.Text)".Trim()
```

New:

```powershell
$emarsUserRaw = "$($c.EmarsUserBox.Text)".Trim()
$emarsUser    = ($emarsUserRaw -split '\s+—\s+', 2)[0].Trim()
if ([string]::IsNullOrWhiteSpace($emarsUser)) { $emarsUser = $emarsUserRaw }
```

C3. Path 2 (fallback) — if the endpoint is admin-only or undiscoverable

Add a Dropdown CSVs\EmarsUsers.csv with header Site ID,User ID,Full Name. Populate it once from a supervisor account. Replace the Get-EmarsUserList call inside $refreshEmarsUsers with:

```powershell
function Get-EmarsUsersFromCsv {
    param([string]$SiteId)
    $path = Join-Path $script:config.DropdownCsvsDirectory 'EmarsUsers.csv'
    if (-not (Test-Path $path)) { return ,@() }
    $rows = Import-Csv -Path $path
    return ,@($rows | Where-Object { "$($_.'Site ID')" -eq $SiteId } |
        ForEach-Object {
            [PSCustomObject]@{ UserId = "$($_."User ID")"; Name = "$($_."Full Name")" }
        })
}
```

Otherwise the ComboBox / Save-handler changes in C2 apply identically — only the data source differs. This matches how Sites.csv, Machines.csv, and Parsed-Parts-Volumes.csv are already seeded, so it needs no new browser inspection and no permission escalation at runtime.

---

One follow-up worth knowing about

The existing Show-PartPicker search handler already references function-locals ($txtNSN, $txtDesc, $lvResults, $lvChosen, $script:__pp_chosen) without $script: or .GetNewClosure(). If Part Picker search has been working for you, PowerShell is resolving them somehow at call time — if it ever starts throwing "cannot call a method on a null-valued expression," wrap that handler in .GetNewClosure(). Everything added in this checklist uses .GetNewClosure() or $script: throughout, so the new code is not exposed to that class of bug.