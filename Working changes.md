All Timer-Related Changes in One Message

This message contains every timer feature — the two from before (breaks, jobs) plus the new one for individual PM tasks. Nothing in this message touches Excel, hyperlinks, or figure highlighting, so you can apply it on top of the Excel changes you already made.

The key design point shared by all three: timers are modeless — start one, keep working, close and reopen the window whenever. Each timer stores its own timestamps in the underlying data file (Breaks.json, Jobs.json, *.session.json) so a state survives an app restart.

---

Change 4 — Break / Lunch countdown timer

4a — Add Show-BreakTimerForm (insert just above function Show-BreakEditor)

```powershell
function Show-BreakTimerForm {
    param(
        [Parameter(Mandatory)][object]$BreakRecord,
        [Parameter(Mandatory)][int]$DefaultMinutes
    )

    $breakId   = "$($BreakRecord.id)"
    $startedAt = Get-Date
    $startHHmm = $startedAt.ToString('HH:mm')
    $totalSecs = $DefaultMinutes * 60

    # Persist start time NOW (durationMin stays 0 = "in progress")
    $rows = Get-Breaks
    if ($null -eq $rows) { $rows = @() }
    foreach ($b in $rows) { if ("$($b.id)" -eq $breakId) { $b.startTime = $startHHmm } }
    Save-Breaks -Breaks $rows | Out-Null
    if (Get-Command Refresh-BreaksList -ErrorAction SilentlyContinue) {
        try { Refresh-BreaksList } catch { }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Break Timer — $($BreakRecord.type)"
    $form.Size = New-Object System.Drawing.Size(500, 380)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = "$($BreakRecord.type)"
    $lblType.Location = New-Object System.Drawing.Point(12, 12)
    $lblType.Size = New-Object System.Drawing.Size(460, 36)
    $lblType.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblType.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblType.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblType)

    $lblCountdown = New-Object System.Windows.Forms.Label
    $lblCountdown.Text = ("{0}:00" -f $DefaultMinutes)
    $lblCountdown.Location = New-Object System.Drawing.Point(12, 52)
    $lblCountdown.Size = New-Object System.Drawing.Size(460, 110)
    $lblCountdown.Font = New-Object System.Drawing.Font("Consolas", 56, [System.Drawing.FontStyle]::Bold)
    $lblCountdown.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $lblCountdown.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblCountdown)

    $lblStarted = New-Object System.Windows.Forms.Label
    $lblStarted.Text = "Started: $($startedAt.ToString('HH:mm:ss'))   Target: $DefaultMinutes min"
    $lblStarted.Location = New-Object System.Drawing.Point(12, 168)
    $lblStarted.Size = New-Object System.Drawing.Size(460, 22)
    $lblStarted.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblStarted.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $form.Controls.Add($lblStarted)

    $lblElapsed = New-Object System.Windows.Forms.Label
    $lblElapsed.Text = "Elapsed: 0:00"
    $lblElapsed.Location = New-Object System.Drawing.Point(12, 194)
    $lblElapsed.Size = New-Object System.Drawing.Size(460, 26)
    $lblElapsed.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblElapsed.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblElapsed)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "You can keep using the app while this runs."
    $lblHint.Location = New-Object System.Drawing.Point(12, 224)
    $lblHint.Size = New-Object System.Drawing.Size(460, 20)
    $lblHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(160,170,185)
    $form.Controls.Add($lblHint)

    $alarmed = $false

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500

    $doFinish = {
        param([bool]$Log)
        $elapsedMin = ((Get-Date) - $startedAt).TotalMinutes
        $all = Get-Breaks
        if ($null -eq $all) { $all = @() }
        foreach ($b in $all) {
            if ("$($b.id)" -ne $breakId) { continue }
            if ($Log) { $b.durationMin = [int][Math]::Max(1, [Math]::Round($elapsedMin)) }
            else      { $b.startTime = ''; $b.durationMin = 0 }
        }
        Save-Breaks -Breaks $all | Out-Null
        if (Get-Command Refresh-BreaksList -ErrorAction SilentlyContinue) { try { Refresh-BreaksList } catch { } }
        if (Get-Command Refresh-WorkBudget  -ErrorAction SilentlyContinue) { try { Refresh-WorkBudget  } catch { } }
        $form.Tag = 'closing'
        $form.Close()
    }.GetNewClosure()

    $timer.Add_Tick({
        $elapsedSec = ((Get-Date) - $startedAt).TotalSeconds
        $remaining  = $totalSecs - $elapsedSec
        if ($remaining -gt 0) {
            $mm = [int]($remaining / 60); $ss = [int]($remaining % 60)
            $lblCountdown.Text = "{0}:{1:00}" -f $mm, $ss
            $lblCountdown.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
        } else {
            if (-not $alarmed) {
                $alarmed = $true
                try { [System.Media.SystemSounds]::Exclamation.Play() } catch { }
                try { $form.WindowState = 'Normal'; $form.Activate(); $form.BringToFront() } catch { }
                $form.Text = "Break Over — $($BreakRecord.type)"
            }
            $over = [Math]::Abs($remaining)
            $mm = [int]($over / 60); $ss = [int]($over % 60)
            $lblCountdown.Text = "-{0}:{1:00}" -f $mm, $ss
            $lblCountdown.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
        }
        $em = [int]($elapsedSec / 60); $es = [int]($elapsedSec % 60)
        $lblElapsed.Text = "Elapsed: {0}:{1:00}" -f $em, $es
    }.GetNewClosure())
    $timer.Start()

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"; $cancelBtn.Size = New-Object System.Drawing.Size(140, 40)
    $cancelBtn.Location = New-Object System.Drawing.Point(150, 258)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelBtn.Add_Click({
        $elapsedMin = ((Get-Date) - $startedAt).TotalMinutes
        $elapsedRounded = [int][Math]::Max(1, [Math]::Round($elapsedMin))
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Log this break ($elapsedRounded min) or discard it?`r`n`r`nYes    — Log it`r`nNo     — Discard`r`nCancel — Keep the timer running",
            "Cancel Break Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { & $doFinish $true } else { & $doFinish $false }
    }.GetNewClosure())
    $form.Controls.Add($cancelBtn)

    $finishBtn = New-Object System.Windows.Forms.Button
    $finishBtn.Text = "Finish && Log"; $finishBtn.Size = New-Object System.Drawing.Size(160, 40)
    $finishBtn.Location = New-Object System.Drawing.Point(300, 258)
    $finishBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $finishBtn.FlatAppearance.BorderSize = 0
    $finishBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $finishBtn.ForeColor = [System.Drawing.Color]::White
    $finishBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $finishBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $finishBtn.Add_Click({ & $doFinish $true }.GetNewClosure())
    $form.Controls.Add($finishBtn)

    $form.Add_FormClosing({
        param($sender, $e)
        if ($sender.Tag -ne 'closing') {
            $e.Cancel = $true
            $elapsedMin = ((Get-Date) - $startedAt).TotalMinutes
            $elapsedRounded = [int][Math]::Max(1, [Math]::Round($elapsedMin))
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "Close the timer and log this break ($elapsedRounded min)?`r`n`r`nYes    — Log and close`r`nNo     — Discard and close`r`nCancel — Keep the timer running",
                "Close Break Timer",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { & $doFinish $true } else { & $doFinish $false }
        }
        $timer.Stop()
    }.GetNewClosure())

    $form.Add_FormClosed({
        try { $timer.Stop(); $timer.Dispose() } catch { }
        if ($null -ne $script:activeBreakTimers) { [void]$script:activeBreakTimers.Remove($breakId) }
    }.GetNewClosure())

    $form.Show()   # modeless
}
```

4b — Replace Refresh-BreaksList

```powershell
function Refresh-BreaksList {
    if (-not $script:breaksListView) { return }
    $script:breaksListView.BeginUpdate()
    try {
        $script:breaksListView.Items.Clear()
        foreach ($b in @(Get-TodaysBreaks)) {
            $dur = 0
            [void][int]::TryParse("$($b.durationMin)", [ref]$dur)
            $hasStart = -not [string]::IsNullOrWhiteSpace("$($b.startTime)")
            $state = if ($dur -gt 0) { 'complete' } elseif ($hasStart) { 'active' } else { 'placeholder' }

            # Columns: [Action] [Type] [Start] [Min] [Notes]
            $it = New-Object System.Windows.Forms.ListViewItem("")
            $it.SubItems.Add("$($b.type)")        | Out-Null
            $it.SubItems.Add("$($b.startTime)")   | Out-Null
            $it.SubItems.Add("$($b.durationMin)") | Out-Null
            $it.SubItems.Add("$($b.notes)")       | Out-Null
            $it.Tag = $b

            switch ($state) {
                'placeholder' {
                    $it.SubItems[0].Text = "Start"
                    $it.UseItemStyleForSubItems = $false
                    $it.SubItems[0].ForeColor = [System.Drawing.Color]::FromArgb(39,174,96)
                    $it.SubItems[0].Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                    foreach ($i in 1..4) { $it.SubItems[$i].ForeColor = [System.Drawing.Color]::FromArgb(150, 160, 175) }
                }
                'active' {
                    $it.SubItems[0].Text = [string][char]0x25CF   # ●
                    $it.UseItemStyleForSubItems = $false
                    $it.SubItems[0].ForeColor = [System.Drawing.Color]::FromArgb(243,156,18)
                    $it.SubItems[0].Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
                    $it.SubItems[3].Text = "running"
                    $it.SubItems[3].ForeColor = [System.Drawing.Color]::FromArgb(243,156,18)
                }
            }
            $script:breaksListView.Items.Add($it) | Out-Null
        }
    } finally {
        $script:breaksListView.EndUpdate()
    }
}
```

4c — Replace Test-MissedBreaksForDate

```powershell
function Test-MissedBreaksForDate {
    param([DateTime]$Date)
    $sched = @(Get-ConfiguredBreakSchedule)
    if ($sched.Count -eq 0) { return ,@() }
    $dateStr = $Date.ToString("yyyy-MM-dd")
    $all = Get-Breaks
    if ($null -eq $all) { return ,@() }

    # Map type -> state: -1 = in progress, 0 = untouched, >0 = logged
    $logged = @{}
    foreach ($b in $all) {
        if ("$($b.date)" -ne $dateStr) { continue }
        $d = 0; [void][int]::TryParse("$($b.durationMin)", [ref]$d)
        $hasStart = -not [string]::IsNullOrWhiteSpace("$($b.startTime)")
        $logged["$($b.type)"] = if ($hasStart -and $d -le 0) { -1 } else { $d }
    }

    $missed = @()
    foreach ($item in $sched) {
        if (-not $logged.ContainsKey($item.Type)) { continue }
        if ($logged[$item.Type] -ne 0) { continue }
        $missed += [PSCustomObject]@{ Type = $item.Type; ScheduledMin = [int]$item.DefaultMinutes }
    }
    return ,$missed
}
```

4d — Add the Start column and click handler in Setup-JobLogSubTab

Find the breaks ListView column definitions and replace with:

```powershell
    $script:breaksListView = New-Object System.Windows.Forms.ListView
    $script:breaksListView.Dock = 'Fill'
    $script:breaksListView.Columns.Add("", 55)       | Out-Null   # Start column
    $script:breaksListView.Columns.Add("Type", 110)  | Out-Null
    $script:breaksListView.Columns.Add("Start", 55)  | Out-Null
    $script:breaksListView.Columns.Add("Min", 40)    | Out-Null
    $script:breaksListView.Columns.Add("Notes", 200) | Out-Null
    Set-ListViewStyle -ListView $script:breaksListView
```

Insert this click handler (anywhere after the ListView exists, e.g. right before the $script:breaksListView.Add_DoubleClick block):

```powershell
    $script:breaksListView.Add_MouseClick({
        param($sender, $e)
        if ($sender.Columns.Count -lt 1) { return }
        if ($e.X -gt $sender.Columns[0].Width) { return }
        $hit = $sender.HitTest($e.Location)
        if (-not $hit.Item) { return }
        $break = $hit.Item.Tag
        if ($null -eq $break) { return }

        $dur = 0; [void][int]::TryParse("$($break.durationMin)", [ref]$dur)
        $hasStart = -not [string]::IsNullOrWhiteSpace("$($break.startTime)")
        if ($dur -gt 0 -or $hasStart) { return }   # complete or already running

        if ($null -eq $script:activeBreakTimers) { $script:activeBreakTimers = @{} }
        $breakId = "$($break.id)"
        if ($script:activeBreakTimers.ContainsKey($breakId)) { return }

        $defaultMin = 15
        foreach ($s in @(Get-ConfiguredBreakSchedule)) {
            if ("$($s.Type)" -eq "$($break.type)") { $defaultMin = [int]$s.DefaultMinutes; break }
        }

        $script:activeBreakTimers[$breakId] = $true
        try { Show-BreakTimerForm -BreakRecord $break -DefaultMinutes $defaultMin }
        catch {
            [void]$script:activeBreakTimers.Remove($breakId)
            Write-Log "Failed to open break timer: $($_.Exception.Message)"
        }
    })
```

4e — Guard $addBreakBtn.Add_Click

Find this snippet inside the handler:

```powershell
            $d = 0
            [void][int]::TryParse("$($b.durationMin)", [ref]$d)
            if ($d -gt 0) { continue }
```

Replace with:

```powershell
            $d = 0
            [void][int]::TryParse("$($b.durationMin)", [ref]$d)
            $hasStart = -not [string]::IsNullOrWhiteSpace("$($b.startTime)")
            if ($d -gt 0 -or $hasStart) { continue }
```

---

Change 5 — Job (reactive / work order) count-up timer with pause

5a — Insert helper functions (before function Get-JobHours)

```powershell
function Get-JobById {
    param([string]$JobId)
    if ([string]::IsNullOrWhiteSpace($JobId)) { return $null }
    foreach ($j in @(Get-Jobs)) { if ("$($j.id)" -eq $JobId) { return $j } }
    return $null
}

function Get-JobTimerState {
    param($Job)
    if ($null -eq $Job) { return 'idle' }
    if ($Job.PSObject.Properties['timerState'] -and -not [string]::IsNullOrWhiteSpace("$($Job.timerState)")) {
        return "$($Job.timerState)"
    }
    return 'idle'
}

function Get-JobTimerElapsedSeconds {
    param($Job)
    if ($null -eq $Job) { return 0 }
    $state = Get-JobTimerState -Job $Job
    if ($state -notin @('running','paused','finished')) { return 0 }
    $startedAt = $null
    try { $startedAt = [DateTime]::Parse("$($Job.timerStartedAt)") } catch { return 0 }
    $pausedTotal = 0
    [void][int]::TryParse("$($Job.timerPausedTotalSeconds)", [ref]$pausedTotal)
    $now = Get-Date; $elapsedSec = 0
    switch ($state) {
        'running'  { $elapsedSec = ($now - $startedAt).TotalSeconds - $pausedTotal }
        'paused'   {
            $pausedAt = $null
            try { $pausedAt = [DateTime]::Parse("$($Job.timerPausedAt)") } catch { $pausedAt = $now }
            $elapsedSec = ($pausedAt - $startedAt).TotalSeconds - $pausedTotal
        }
        'finished' {
            $endedAt = $null
            try { $endedAt = [DateTime]::Parse("$($Job.timerEndedAt)") } catch { return 0 }
            $elapsedSec = ($endedAt - $startedAt).TotalSeconds - $pausedTotal
        }
    }
    if ($elapsedSec -lt 0) { $elapsedSec = 0 }
    return [int]$elapsedSec
}

function Format-JobTimerElapsed {
    param([int]$Seconds)
    if ($Seconds -lt 0) { $Seconds = 0 }
    $h = [int]($Seconds / 3600); $m = [int](($Seconds % 3600) / 60); $s = [int]($Seconds % 60)
    if ($h -gt 0) { return "{0}:{1:00}:{2:00}" -f $h, $m, $s }
    return "{0}:{1:00}" -f $m, $s
}

function Start-JobTimer {
    param([string]$JobId)
    $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
    $now = Get-Date
    foreach ($j in $jobs) {
        if ("$($j.id)" -ne $JobId) { continue }
        $j | Add-Member -NotePropertyName 'timerState'              -NotePropertyValue 'running'         -Force
        $j | Add-Member -NotePropertyName 'timerStartedAt'          -NotePropertyValue $now.ToString("o") -Force
        $j | Add-Member -NotePropertyName 'timerPausedAt'           -NotePropertyValue $null             -Force
        $j | Add-Member -NotePropertyName 'timerPausedTotalSeconds' -NotePropertyValue 0                 -Force
        $j | Add-Member -NotePropertyName 'timerEndedAt'            -NotePropertyValue $null             -Force
        $j | Add-Member -NotePropertyName 'timerComputedHours'      -NotePropertyValue $null             -Force
        $j.startTime = $now.ToString('HH:mm'); $j.endTime = $null
        Write-Log "Start-JobTimer: started $JobId at $($now.ToString('HH:mm:ss'))"
    }
    Save-Jobs -Jobs $jobs | Out-Null
    if (Get-Command Refresh-JobLogList -ErrorAction SilentlyContinue) { try { Refresh-JobLogList } catch { } }
}

function Pause-JobTimer {
    param([string]$JobId)
    $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
    $now = Get-Date
    foreach ($j in $jobs) {
        if ("$($j.id)" -ne $JobId) { continue }
        if ((Get-JobTimerState -Job $j) -ne 'running') { return }
        $j | Add-Member -NotePropertyName 'timerPausedAt' -NotePropertyValue $now.ToString("o") -Force
        $j | Add-Member -NotePropertyName 'timerState'    -NotePropertyValue 'paused' -Force
    }
    Save-Jobs -Jobs $jobs | Out-Null
    if (Get-Command Refresh-JobLogList -ErrorAction SilentlyContinue) { try { Refresh-JobLogList } catch { } }
}

function Resume-JobTimer {
    param([string]$JobId)
    $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
    $now = Get-Date
    foreach ($j in $jobs) {
        if ("$($j.id)" -ne $JobId) { continue }
        if ((Get-JobTimerState -Job $j) -ne 'paused') { return }
        $pausedAt = $null
        try { $pausedAt = [DateTime]::Parse("$($j.timerPausedAt)") } catch { $pausedAt = $now }
        $extraPause = [int](($now - $pausedAt).TotalSeconds)
        if ($extraPause -lt 0) { $extraPause = 0 }
        $prevTotal = 0; [void][int]::TryParse("$($j.timerPausedTotalSeconds)", [ref]$prevTotal)
        $j | Add-Member -NotePropertyName 'timerPausedTotalSeconds' -NotePropertyValue ($prevTotal + $extraPause) -Force
        $j | Add-Member -NotePropertyName 'timerPausedAt'           -NotePropertyValue $null -Force
        $j | Add-Member -NotePropertyName 'timerState'              -NotePropertyValue 'running' -Force
    }
    Save-Jobs -Jobs $jobs | Out-Null
    if (Get-Command Refresh-JobLogList -ErrorAction SilentlyContinue) { try { Refresh-JobLogList } catch { } }
}

function Finish-JobTimer {
    param([string]$JobId)
    $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
    $now = Get-Date
    foreach ($j in $jobs) {
        if ("$($j.id)" -ne $JobId) { continue }
        $state = Get-JobTimerState -Job $j
        if ($state -notin @('running','paused')) { return }
        $pausedTotal = 0; [void][int]::TryParse("$($j.timerPausedTotalSeconds)", [ref]$pausedTotal)
        if ($state -eq 'paused') {
            $pausedAt = $null
            try { $pausedAt = [DateTime]::Parse("$($j.timerPausedAt)") } catch { $pausedAt = $now }
            $extra = [int](($now - $pausedAt).TotalSeconds)
            if ($extra -gt 0) { $pausedTotal += $extra }
        }
        $startedAt = $null
        try { $startedAt = [DateTime]::Parse("$($j.timerStartedAt)") } catch { $startedAt = $now }
        $workedSeconds = ($now - $startedAt).TotalSeconds - $pausedTotal
        if ($workedSeconds -lt 0) { $workedSeconds = 0 }
        $workedHours = [Math]::Round($workedSeconds / 3600.0, 2)
        $j | Add-Member -NotePropertyName 'timerPausedTotalSeconds' -NotePropertyValue $pausedTotal -Force
        $j | Add-Member -NotePropertyName 'timerEndedAt'            -NotePropertyValue $now.ToString("o") -Force
        $j | Add-Member -NotePropertyName 'timerComputedHours'      -NotePropertyValue $workedHours -Force
        $j | Add-Member -NotePropertyName 'timerState'              -NotePropertyValue 'finished' -Force
        $j.endTime = $now.ToString('HH:mm')
    }
    Save-Jobs -Jobs $jobs | Out-Null
    if (Get-Command Refresh-JobLogList -ErrorAction SilentlyContinue) { try { Refresh-JobLogList } catch { } }
    if (Get-Command Refresh-WorkBudget  -ErrorAction SilentlyContinue) { try { Refresh-WorkBudget  } catch { } }
}

function Reset-JobTimer {
    param([string]$JobId)
    $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
    foreach ($j in $jobs) {
        if ("$($j.id)" -ne $JobId) { continue }
        foreach ($k in 'timerState','timerStartedAt','timerPausedAt','timerEndedAt','timerComputedHours') {
            $j | Add-Member -NotePropertyName $k -NotePropertyValue $null -Force
        }
        $j | Add-Member -NotePropertyName 'timerPausedTotalSeconds' -NotePropertyValue 0 -Force
        $j | Add-Member -NotePropertyName 'timerState'              -NotePropertyValue 'idle' -Force
        $j.startTime = ''; $j.endTime = ''
    }
    Save-Jobs -Jobs $jobs | Out-Null
    if (Get-Command Refresh-JobLogList -ErrorAction SilentlyContinue) { try { Refresh-JobLogList } catch { } }
    if (Get-Command Refresh-WorkBudget  -ErrorAction SilentlyContinue) { try { Refresh-WorkBudget  } catch { } }
}
```

5b — Insert Show-JobTimerForm (right after the helpers above)

```powershell
function Show-JobTimerForm {
    param([Parameter(Mandatory)][string]$JobId)
    if (-not $script:jobTimerForms) { $script:jobTimerForms = @{} }
    if ($script:jobTimerForms.ContainsKey($JobId)) {
        $existing = $script:jobTimerForms[$JobId]
        if ($existing -and -not $existing.IsDisposed) {
            $existing.WindowState = 'Normal'; $existing.Activate(); $existing.BringToFront(); return
        }
    }
    $job = Get-JobById -JobId $JobId
    if ($null -eq $job) { return }
    $kindLabel = if ("$($job.kind)" -eq 'workorder') { 'Work Order' } else { 'Reactive Call' }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Timer — $($job.machineId) ($kindLabel)"
    $form.Size = New-Object System.Drawing.Size(540, 430)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ShowInTaskbar = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "$($job.machineId) — $kindLabel"
    $lblTitle.Location = New-Object System.Drawing.Point(12, 14)
    $lblTitle.Size = New-Object System.Drawing.Size(500, 30)
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblTitle)

    $lblElapsed = New-Object System.Windows.Forms.Label
    $lblElapsed.Text = "0:00"
    $lblElapsed.Location = New-Object System.Drawing.Point(12, 48)
    $lblElapsed.Size = New-Object System.Drawing.Size(500, 100)
    $lblElapsed.Font = New-Object System.Drawing.Font("Consolas", 52, [System.Drawing.FontStyle]::Bold)
    $lblElapsed.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblElapsed)

    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Location = New-Object System.Drawing.Point(12, 150)
    $lblState.Size = New-Object System.Drawing.Size(500, 30)
    $lblState.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblState.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblState)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(12, 184)
    $lblInfo.Size = New-Object System.Drawing.Size(500, 44)
    $lblInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $lblInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblInfo)

    $pauseResumeBtn = New-Object System.Windows.Forms.Button
    $pauseResumeBtn.Text = "Pause"
    $pauseResumeBtn.Size = New-Object System.Drawing.Size(130, 44)
    $pauseResumeBtn.Location = New-Object System.Drawing.Point(115, 246)
    $pauseResumeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pauseResumeBtn.FlatAppearance.BorderSize = 0
    $pauseResumeBtn.ForeColor = [System.Drawing.Color]::White
    $pauseResumeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pauseResumeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($pauseResumeBtn)

    $finishBtn = New-Object System.Windows.Forms.Button
    $finishBtn.Text = "Finish && Log"
    $finishBtn.Size = New-Object System.Drawing.Size(140, 44)
    $finishBtn.Location = New-Object System.Drawing.Point(255, 246)
    $finishBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $finishBtn.FlatAppearance.BorderSize = 0
    $finishBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $finishBtn.ForeColor = [System.Drawing.Color]::White
    $finishBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $finishBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($finishBtn)

    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = "Reset Timer"
    $resetBtn.Size = New-Object System.Drawing.Size(140, 44)
    $resetBtn.Location = New-Object System.Drawing.Point(200, 246)
    $resetBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $resetBtn.FlatAppearance.BorderSize = 0
    $resetBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $resetBtn.ForeColor = [System.Drawing.Color]::White
    $resetBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $resetBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $resetBtn.Visible = $false
    $form.Controls.Add($resetBtn)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "You can keep using the app while this runs. Closing this window does not stop the timer."
    $lblHint.Location = New-Object System.Drawing.Point(12, 300)
    $lblHint.Size = New-Object System.Drawing.Size(500, 22)
    $lblHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(160,170,185)
    $form.Controls.Add($lblHint)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"; $closeBtn.Size = New-Object System.Drawing.Size(100, 34)
    $closeBtn.Location = New-Object System.Drawing.Point(220, 330)
    $closeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $closeBtn.Add_Click({ $form.Close() })
    $form.Controls.Add($closeBtn)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500

    $refreshDisplay = {
        $currentJob = Get-JobById -JobId $JobId
        if ($null -eq $currentJob) { $form.Close(); return }
        $state = Get-JobTimerState -Job $currentJob
        $sec = Get-JobTimerElapsedSeconds -Job $currentJob
        $lblElapsed.Text = Format-JobTimerElapsed -Seconds $sec

        switch ($state) {
            'running' {
                $lblState.Text = "Running"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $pauseResumeBtn.Text = "Pause"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(243,156,18)
                $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true; $resetBtn.Visible = $false
            }
            'paused' {
                $lblState.Text = "Paused"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(243,156,18)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
                $pauseResumeBtn.Text = "Resume"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true; $resetBtn.Visible = $false
            }
            'finished' {
                $lblState.Text = "Finished"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(39,174,96)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
                $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false; $resetBtn.Visible = $true
            }
            default {
                $lblState.Text = "Idle"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
                $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false; $resetBtn.Visible = $false
            }
        }

        $info = @()
        if ($currentJob.startTime) { $info += "Started: $($currentJob.startTime)" }
        if ($currentJob.endTime)   { $info += "Ended: $($currentJob.endTime)" }
        $pt = 0; [void][int]::TryParse("$($currentJob.timerPausedTotalSeconds)", [ref]$pt)
        if ($pt -gt 0) { $info += "Paused: $([Math]::Round($pt / 60.0, 1)) min" }
        $lblInfo.Text = ($info -join "     ")
    }.GetNewClosure()

    $pauseResumeBtn.Add_Click({
        $currentJob = Get-JobById -JobId $JobId
        if ($null -eq $currentJob) { return }
        $state = Get-JobTimerState -Job $currentJob
        if ($state -eq 'running')  { Pause-JobTimer  -JobId $JobId }
        elseif ($state -eq 'paused') { Resume-JobTimer -JobId $JobId }
        & $refreshDisplay
    }.GetNewClosure())

    $finishBtn.Add_Click({
        $currentJob = Get-JobById -JobId $JobId
        if ($null -eq $currentJob) { return }
        $state = Get-JobTimerState -Job $currentJob
        if ($state -notin @('running','paused')) { return }
        $sec = Get-JobTimerElapsedSeconds -Job $currentJob
        $hms = Format-JobTimerElapsed -Seconds $sec
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Finish this timer?`r`n`r`nWorked time: $hms`r`n`r`nThe end time will be recorded and this duration used for the job.",
            "Finish Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Finish-JobTimer -JobId $JobId
        & $refreshDisplay
    }.GetNewClosure())

    $resetBtn.Add_Click({
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Reset this timer? This clears the start time, end time, and computed hours for this job.",
            "Reset Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Reset-JobTimer -JobId $JobId
        & $refreshDisplay
    }.GetNewClosure())

    $timer.Add_Tick({ & $refreshDisplay }.GetNewClosure())
    $timer.Start()

    $form.Add_FormClosed({
        try { $timer.Stop(); $timer.Dispose() } catch { }
        if ($script:jobTimerForms) { [void]$script:jobTimerForms.Remove($JobId) }
    }.GetNewClosure())

    $script:jobTimerForms[$JobId] = $form
    & $refreshDisplay
    $form.Show()
}
```

5c — Replace Get-JobHours and New-JobObject

```powershell
function Get-JobHours {
    param($Job)
    if ($null -eq $Job) { return $null }
    if ($null -ne $Job.hoursOverride -and "$($Job.hoursOverride)" -ne '') {
        $h = 0.0
        if ([double]::TryParse("$($Job.hoursOverride)", [ref]$h)) { return $h }
    }
    if ($Job.PSObject.Properties['timerComputedHours'] -and
        $null -ne $Job.timerComputedHours -and
        "$($Job.timerComputedHours)" -ne '') {
        $h = 0.0
        if ([double]::TryParse("$($Job.timerComputedHours)", [ref]$h)) { return $h }
    }
    if ([string]::IsNullOrWhiteSpace($Job.startTime) -or [string]::IsNullOrWhiteSpace($Job.endTime)) { return $null }
    try {
        $s = [DateTime]::ParseExact($Job.startTime, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        $e = [DateTime]::ParseExact($Job.endTime,   'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($e -lt $s) { $e = $e.AddDays(1) }
        return [Math]::Round(($e - $s).TotalHours, 2)
    } catch { return $null }
}

function New-JobObject {
    param([string]$Kind)
    return [PSCustomObject]@{
        id              = [guid]::NewGuid().ToString()
        kind            = $Kind
        machineId       = ''
        startTime       = ''
        endTime         = ''
        hoursOverride   = $null
        description     = ''
        workOrderNo     = ''
        parts           = @()
        createdAt       = (Get-Date).ToString("o")
        sentToHistorian = $false
        sentAt          = $null
        timerState              = 'idle'
        timerStartedAt          = $null
        timerPausedAt           = $null
        timerPausedTotalSeconds = 0
        timerEndedAt            = $null
        timerComputedHours      = $null
    }
}
```

5d — Replace Refresh-JobLogList

```powershell
function Refresh-JobLogList {
    if (-not $script:jobLogListView) { return }
    $allJobs = Get-Jobs; if ($null -eq $allJobs) { $allJobs = @() }
    $script:wtJobs = $allJobs
    $selected = if ($script:jobLogDatePicker) { $script:jobLogDatePicker.Value.Date } else { (Get-Date).Date }

    $jobs = @()
    foreach ($job in $allJobs) {
        $jobDate = $null
        if ($job.createdAt) { try { $jobDate = ([DateTime]$job.createdAt).Date } catch { } }
        $onSelected  = ($null -ne $jobDate) -and ($jobDate -eq $selected)
        $unsent      = (-not $job.sentToHistorian)
        $timerActive = (Get-JobTimerState -Job $job) -in @('running','paused')
        if ($onSelected -or $unsent -or $timerActive) { $jobs += $job }
    }

    $script:jobLogListView.BeginUpdate()
    try {
        $script:jobLogListView.Items.Clear()
        foreach ($job in $jobs) {
            $kindText = if ($job.kind -eq 'workorder') { 'WORK ORDER' } else { 'REACTIVE' }
            $dateText = ''
            if ($job.createdAt) {
                try { $dateText = ([DateTime]$job.createdAt).ToString("yyyy-MM-dd") }
                catch { $dateText = "$($job.createdAt)".Split('T')[0] }
            }
            $hrsText = ''
            $h = Get-JobHours $job
            if ($null -ne $h) { $hrsText = $h.ToString("F2") }
            $partsCount = if ($job.parts) { @($job.parts).Count } else { 0 }
            $statusText = if ($job.sentToHistorian) { 'SENT' } else { 'DRAFT' }

            $tState = Get-JobTimerState -Job $job
            $tSec = Get-JobTimerElapsedSeconds -Job $job
            $tStr = Format-JobTimerElapsed -Seconds $tSec
            switch ($tState) {
                'running'  { $timerText = "> $tStr";   $timerColor = [System.Drawing.Color]::FromArgb(52,152,219) }
                'paused'   { $timerText = "II $tStr";  $timerColor = [System.Drawing.Color]::FromArgb(243,156,18) }
                'finished' { $timerText = "OK $tStr";  $timerColor = [System.Drawing.Color]::FromArgb(39,174,96)  }
                default    { $timerText = "START";     $timerColor = [System.Drawing.Color]::FromArgb(52,152,219) }
            }

            $item = New-Object System.Windows.Forms.ListViewItem($kindText)
            $item.SubItems.Add("$($job.machineId)")   | Out-Null
            $item.SubItems.Add($dateText)             | Out-Null
            $item.SubItems.Add("$($job.startTime)")   | Out-Null
            $item.SubItems.Add("$($job.endTime)")     | Out-Null
            $item.SubItems.Add($hrsText)              | Out-Null
            $item.SubItems.Add("$($job.workOrderNo)") | Out-Null
            $item.SubItems.Add("$partsCount")         | Out-Null
            $item.SubItems.Add("$($job.description)") | Out-Null
            $item.SubItems.Add($statusText)           | Out-Null
            $item.SubItems.Add($timerText)            | Out-Null  # index 10
            $item.Tag = $job
            $item.UseItemStyleForSubItems = $false
            $item.SubItems[10].ForeColor = $timerColor
            $item.SubItems[10].Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
            $script:jobLogListView.Items.Add($item) | Out-Null
        }
    } finally {
        $script:jobLogListView.EndUpdate()
    }

    if ($script:jobLogStatusLabel) {
        $shown  = @($jobs).Count
        $total  = @($allJobs).Count
        $drafts = @($jobs | Where-Object { -not $_.sentToHistorian }).Count
        $selStr = $selected.ToString('yyyy-MM-dd')
        $script:jobLogStatusLabel.Text = "Showing $shown of $total — $selStr — $drafts draft$(if ($drafts -ne 1) { 's' })"
    }
    Refresh-JobTotals
}
```

5e — Edits inside Setup-JobLogSubTab

Column definition — replace:

```powershell
    $script:jobLogListView.Columns.Add("Description", 320) | Out-Null
    $script:jobLogListView.Columns.Add("Status", 70)       | Out-Null
```

with:

```powershell
    $script:jobLogListView.Columns.Add("Description", 320) | Out-Null
    $script:jobLogListView.Columns.Add("Status", 70)       | Out-Null
    $script:jobLogListView.Columns.Add("Timer", 110)       | Out-Null
```

Replace the + New Reactive Call and + New Work Order button blocks so they check StartTimer:

```powershell
    $newReactiveBtn = New-Button "+ New Reactive Call" {
        $result = Show-JobEditor -Job $null -Kind 'reactive'
        if ($null -ne $result -and $result.Action -eq 'save') {
            $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
            $jobs += $result.Job
            Save-Jobs -Jobs $jobs | Out-Null
            Refresh-JobLogList; Refresh-WorkBudget
            if ($result.StartTimer) { Show-JobTimerForm -JobId "$($result.Job.id)" }
        }
    } -Style 'Primary' -Width 180 -Height 32 -TextAlign MiddleCenter

    $newWoBtn = New-Button "+ New Work Order" {
        $result = Show-JobEditor -Job $null -Kind 'workorder'
        if ($null -ne $result -and $result.Action -eq 'save') {
            $jobs = Get-Jobs; if ($null -eq $jobs) { $jobs = @() }
            $jobs += $result.Job
            Save-Jobs -Jobs $jobs | Out-Null
            Refresh-JobLogList; Refresh-WorkBudget
            if ($result.StartTimer) { Show-JobTimerForm -JobId "$($result.Job.id)" }
        }
    } -Style 'Primary' -Width 170 -Height 32 -TextAlign MiddleCenter
```

Add the Timer-column click handler (anywhere after the ListView is created):

```powershell
    $script:jobLogListView.Add_MouseClick({
        param($sender, $e)
        if ($sender.Columns.Count -lt 11) { return }
        $timerColX = 0
        for ($i = 0; $i -lt 10; $i++) { $timerColX += $sender.Columns[$i].Width }
        if ($e.X -lt $timerColX) { return }
        if ($e.X -gt ($timerColX + $sender.Columns[10].Width)) { return }
        $hit = $sender.HitTest($e.Location)
        if (-not $hit.Item) { return }
        $job = $hit.Item.Tag
        if ($null -eq $job) { return }
        $jobId = "$($job.id)"
        if ([string]::IsNullOrWhiteSpace($jobId)) { return }

        if ($script:jobTimerForms -and $script:jobTimerForms.ContainsKey($jobId)) {
            $existing = $script:jobTimerForms[$jobId]
            if ($existing -and -not $existing.IsDisposed) {
                $existing.WindowState = 'Normal'; $existing.Activate(); $existing.BringToFront(); return
            }
        }
        if ((Get-JobTimerState -Job $job) -eq 'idle') {
            Start-JobTimer -JobId $jobId
            Refresh-WorkBudget
        }
        Show-JobTimerForm -JobId $jobId
    })
```

Add live timer ticker — find the block that ends with:

```powershell
    # Initial loads
    Refresh-JobLogList
    Refresh-BreaksList
    Refresh-WorkBudget

    Write-Log "Job Log sub-tab setup completed."
```

Replace with:

```powershell
    # Initial loads
    Refresh-JobLogList
    Refresh-BreaksList
    Refresh-WorkBudget

    if (-not $script:jobTimerListTicker) {
        $script:jobTimerListTicker = New-Object System.Windows.Forms.Timer
        $script:jobTimerListTicker.Interval = 1000
        $script:jobTimerListTicker.Add_Tick({
            if (-not $script:jobLogListView) { return }
            if ($script:jobLogListView.Columns.Count -lt 11) { return }
            foreach ($item in $script:jobLogListView.Items) {
                $job = $item.Tag
                if ($null -eq $job) { continue }
                if ((Get-JobTimerState -Job $job) -ne 'running') { continue }
                $sec = Get-JobTimerElapsedSeconds -Job $job
                $item.SubItems[10].Text = "> " + (Format-JobTimerElapsed -Seconds $sec)
            }
        })
        $script:jobTimerListTicker.Start()
    }

    Write-Log "Job Log sub-tab setup completed."
```

5f — Edits inside Show-JobEditor

a) Add the Start Timer button — insert right after the $deleteBtn block:

```powershell
    $startTimerBtn = New-Object System.Windows.Forms.Button
    $startTimerBtn.Text = "Start Timer"
    $startTimerBtn.Location = New-Object System.Drawing.Point(110, 10)
    $startTimerBtn.Size = New-Object System.Drawing.Size(110, 32)
    $startTimerBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $startTimerBtn.ForeColor = [System.Drawing.Color]::White
    $startTimerBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $startTimerBtn.FlatAppearance.BorderSize = 0
    $startTimerBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $startTimerBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBar.Controls.Add($startTimerBtn)
```

b) Lock the time pickers when a timer owns them — after the existing try { $dtpStart.Value = ... } catch {} block, append:

```powershell
    $__timerStateForEditor = Get-JobTimerState -Job $Job
    if ($__timerStateForEditor -in @('running','paused','finished')) {
        $dtpStart.Enabled = $false; $dtpEnd.Enabled = $false
        $btnStartNow.Enabled = $false; $btnEndNow.Enabled = $false
        $lblHours.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
    }
```

c) Preserve timer fields in the $candidate inside $saveBtn.Add_Click — find the candidate block and append these 6 lines to it (inside the same @{ ... }):

```powershell
            timerState              = if ($Job.PSObject.Properties['timerState'])              { $Job.timerState }              else { 'idle' }
            timerStartedAt          = if ($Job.PSObject.Properties['timerStartedAt'])          { $Job.timerStartedAt }          else { $null }
            timerPausedAt           = if ($Job.PSObject.Properties['timerPausedAt'])           { $Job.timerPausedAt }           else { $null }
            timerPausedTotalSeconds = if ($Job.PSObject.Properties['timerPausedTotalSeconds']) { $Job.timerPausedTotalSeconds } else { 0 }
            timerEndedAt            = if ($Job.PSObject.Properties['timerEndedAt'])            { $Job.timerEndedAt }            else { $null }
            timerComputedHours      = if ($Job.PSObject.Properties['timerComputedHours'])      { $Job.timerComputedHours }      else { $null }
```

d) Add the Start Timer button handler — insert right after $cancelBtn.Add_Click({...}):

```powershell
    $startTimerBtn.Add_Click({
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

---

Change 6 — PM task count-up timer ⭐ NEW

Per-task timer that lives in the PM session file and overwrites CustomTimeMin on finish. Any running timer's elapsed time can be discarded (Reset) or committed (Finish).

6a — Replace Get-PmSession and Save-PmSession (preserve timer fields)

```powershell
function Get-PmSession {
    param($Machine)
    $path = Get-PmSessionPath -JsonPath $Machine.JsonPath
    $empty = @{ Selections = @{} }
    if (-not (Test-Path $path)) { return $empty }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $empty }
        $data = $raw | ConvertFrom-Json
        $selections = @{}
        if ($data.Selections) {
            foreach ($prop in $data.Selections.PSObject.Properties) {
                $entry = $prop.Value
                $selections[$prop.Name] = @{
                    Selected             = [bool]$entry.Selected
                    CustomTimeMin        = if ($null -ne $entry.CustomTimeMin) { [int]$entry.CustomTimeMin } else { $null }
                    TimerState           = if ($entry.PSObject.Properties['TimerState'])           { "$($entry.TimerState)" } else { 'idle' }
                    TimerStartedAt       = if ($entry.PSObject.Properties['TimerStartedAt'])       { $entry.TimerStartedAt }  else { $null }
                    TimerPausedAt        = if ($entry.PSObject.Properties['TimerPausedAt'])        { $entry.TimerPausedAt }   else { $null }
                    TimerPausedTotalSecs = if ($entry.PSObject.Properties['TimerPausedTotalSecs']) { [int]$entry.TimerPausedTotalSecs } else { 0 }
                    TimerEndedAt         = if ($entry.PSObject.Properties['TimerEndedAt'])         { $entry.TimerEndedAt }    else { $null }
                }
            }
        }
        return @{ Selections = $selections }
    } catch {
        Write-Log "Get-PmSession: read error: $($_.Exception.Message)"
        return $empty
    }
}

function Save-PmSession {
    param($Machine)
    if (-not $Machine -or -not $Machine.State) { return }
    $path = Get-PmSessionPath -JsonPath $Machine.JsonPath

    $selObj = [ordered]@{}
    foreach ($key in $Machine.State.Keys) {
        $e = $Machine.State[$key]
        $selObj[$key] = [PSCustomObject]@{
            Selected             = [bool]$e.Selected
            CustomTimeMin        = if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { $null }
            TimerState           = if ($e.ContainsKey('TimerState') -and $e['TimerState']) { "$($e['TimerState'])" } else { 'idle' }
            TimerStartedAt       = if ($e.ContainsKey('TimerStartedAt'))       { $e['TimerStartedAt'] }       else { $null }
            TimerPausedAt        = if ($e.ContainsKey('TimerPausedAt'))        { $e['TimerPausedAt'] }        else { $null }
            TimerPausedTotalSecs = if ($e.ContainsKey('TimerPausedTotalSecs')) { [int]$e['TimerPausedTotalSecs'] } else { 0 }
            TimerEndedAt         = if ($e.ContainsKey('TimerEndedAt'))         { $e['TimerEndedAt'] }         else { $null }
        }
    }

    $payload = [PSCustomObject]@{
        machineId     = $Machine.MachineId
        checklistDate = $Machine.ChecklistDate
        checklistNo   = $Machine.ChecklistNo
        updatedAt     = (Get-Date).ToString("o")
        selections    = $selObj
    }
    try { $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8 }
    catch { Write-Log "Save-PmSession: error: $($_.Exception.Message)" }
}
```

6b — Add PM task timer helpers (insert just above function Show-PmTaskEditor)

```powershell
function Get-PmTaskTimerState {
    param($Entry)
    if ($null -eq $Entry) { return 'idle' }
    if ($Entry.TimerState) { return "$($Entry.TimerState)" }
    return 'idle'
}

function Get-PmTaskTimerElapsedSeconds {
    param($Entry)
    if ($null -eq $Entry) { return 0 }
    $state = Get-PmTaskTimerState -Entry $Entry
    if ($state -notin @('running','paused','finished')) { return 0 }
    $startedAt = $null
    try { $startedAt = [DateTime]::Parse("$($Entry.TimerStartedAt)") } catch { return 0 }
    $pausedTotal = 0
    [void][int]::TryParse("$($Entry.TimerPausedTotalSecs)", [ref]$pausedTotal)
    $now = Get-Date; $elapsedSec = 0
    switch ($state) {
        'running'  { $elapsedSec = ($now - $startedAt).TotalSeconds - $pausedTotal }
        'paused'   {
            $pausedAt = $null
            try { $pausedAt = [DateTime]::Parse("$($Entry.TimerPausedAt)") } catch { $pausedAt = $now }
            $elapsedSec = ($pausedAt - $startedAt).TotalSeconds - $pausedTotal
        }
        'finished' {
            $endedAt = $null
            try { $endedAt = [DateTime]::Parse("$($Entry.TimerEndedAt)") } catch { return 0 }
            $elapsedSec = ($endedAt - $startedAt).TotalSeconds - $pausedTotal
        }
    }
    if ($elapsedSec -lt 0) { $elapsedSec = 0 }
    return [int]$elapsedSec
}

function Start-PmTaskTimer {
    param($Machine, $Task)
    if (-not $Machine -or -not $Task) { return }
    $taskId = "$($Task.ItemNo)"
    $existing = $Machine.State[$taskId]
    $selected  = if ($existing) { [bool]$existing.Selected } else { $false }
    $customMin = if ($existing -and $null -ne $existing.CustomTimeMin) { [int]$existing.CustomTimeMin } else { $null }
    $now = Get-Date
    $Machine.State[$taskId] = @{
        Selected             = $selected
        CustomTimeMin        = $customMin
        TimerState           = 'running'
        TimerStartedAt       = $now.ToString("o")
        TimerPausedAt        = $null
        TimerPausedTotalSecs = 0
        TimerEndedAt         = $null
    }
    Save-PmSession -Machine $Machine
    Write-Log "PM task timer started: $($Machine.MachineId) task $taskId at $($now.ToString('HH:mm:ss'))"
}

function Pause-PmTaskTimer {
    param($Machine, $Task)
    if (-not $Machine -or -not $Task) { return }
    $taskId = "$($Task.ItemNo)"
    $e = $Machine.State[$taskId]
    if (-not $e -or (Get-PmTaskTimerState -Entry $e) -ne 'running') { return }
    $e.TimerPausedAt = (Get-Date).ToString("o")
    $e.TimerState    = 'paused'
    Save-PmSession -Machine $Machine
}

function Resume-PmTaskTimer {
    param($Machine, $Task)
    if (-not $Machine -or -not $Task) { return }
    $taskId = "$($Task.ItemNo)"
    $e = $Machine.State[$taskId]
    if (-not $e -or (Get-PmTaskTimerState -Entry $e) -ne 'paused') { return }
    $now = Get-Date
    $pausedAt = $null
    try { $pausedAt = [DateTime]::Parse("$($e.TimerPausedAt)") } catch { $pausedAt = $now }
    $extra = [int](($now - $pausedAt).TotalSeconds)
    if ($extra -lt 0) { $extra = 0 }
    $prev = 0; [void][int]::TryParse("$($e.TimerPausedTotalSecs)", [ref]$prev)
    $e.TimerPausedTotalSecs = $prev + $extra
    $e.TimerPausedAt = $null
    $e.TimerState    = 'running'
    Save-PmSession -Machine $Machine
}

function Finish-PmTaskTimer {
    param($Machine, $Task)
    if (-not $Machine -or -not $Task) { return }
    $taskId = "$($Task.ItemNo)"
    $e = $Machine.State[$taskId]
    if (-not $e) { return }
    $state = Get-PmTaskTimerState -Entry $e
    if ($state -notin @('running','paused')) { return }

    $now = Get-Date
    $pausedTotal = 0
    [void][int]::TryParse("$($e.TimerPausedTotalSecs)", [ref]$pausedTotal)
    if ($state -eq 'paused') {
        $pausedAt = $null
        try { $pausedAt = [DateTime]::Parse("$($e.TimerPausedAt)") } catch { $pausedAt = $now }
        $extra = [int](($now - $pausedAt).TotalSeconds)
        if ($extra -gt 0) { $pausedTotal += $extra }
    }

    $startedAt = $null
    try { $startedAt = [DateTime]::Parse("$($e.TimerStartedAt)") } catch { $startedAt = $now }
    $workedSecs = ($now - $startedAt).TotalSeconds - $pausedTotal
    if ($workedSecs -lt 0) { $workedSecs = 0 }
    # The resulting time OVERWRITES the task's actual time (CustomTimeMin).
    $mins = [int][Math]::Max(1, [Math]::Round($workedSecs / 60.0))

    $e.CustomTimeMin        = $mins
    $e.TimerState           = 'finished'
    $e.TimerPausedAt        = $null
    $e.TimerPausedTotalSecs = $pausedTotal
    $e.TimerEndedAt         = $now.ToString("o")
    Save-PmSession -Machine $Machine
    Update-PmSelectedHoursTotal
    Write-Log "PM task timer finished: $($Machine.MachineId) task $taskId — $mins min (actual time updated)"
}

function Reset-PmTaskTimer {
    param($Machine, $Task)
    if (-not $Machine -or -not $Task) { return }
    $taskId = "$($Task.ItemNo)"
    $e = $Machine.State[$taskId]
    if (-not $e) { return }
    $e.TimerState           = 'idle'
    $e.TimerStartedAt       = $null
    $e.TimerPausedAt        = $null
    $e.TimerPausedTotalSecs = 0
    $e.TimerEndedAt         = $null
    Save-PmSession -Machine $Machine
    Update-PmSelectedHoursTotal
}
```

6c — Add Show-PmTaskTimerForm and Refresh-PmTaskRow (insert after the helpers above)

```powershell
function Refresh-PmTaskRow {
    param($Machine, $Task)
    if (-not $script:pmMachineListViews) { return }
    $lv = $script:pmMachineListViews[$Machine.MachineId]
    if (-not $lv -or $lv.IsDisposed) { return }
    $taskId = "$($Task.ItemNo)"
    foreach ($item in $lv.Items) {
        if ("$($item.Tag.ItemNo)" -ne $taskId) { continue }
        $e = $Machine.State[$taskId]
        $state = Get-PmTaskTimerState -Entry $e
        $sec = Get-PmTaskTimerElapsedSeconds -Entry $e
        $str = Format-JobTimerElapsed -Seconds $sec
        switch ($state) {
            'running'  { $item.SubItems[8].Text = "> $str";  $item.SubItems[8].ForeColor = [System.Drawing.Color]::FromArgb(52,152,219) }
            'paused'   { $item.SubItems[8].Text = "II $str"; $item.SubItems[8].ForeColor = [System.Drawing.Color]::FromArgb(243,156,18) }
            'finished' {
                $item.SubItems[8].Text = "OK $str"
                $item.SubItems[8].ForeColor = [System.Drawing.Color]::FromArgb(39,174,96)
                if ($e -and $null -ne $e.CustomTimeMin) { $item.SubItems[6].Text = "$($e.CustomTimeMin)" }
            }
            default    { $item.SubItems[8].Text = "START"; $item.SubItems[8].ForeColor = [System.Drawing.Color]::FromArgb(39,174,96) }
        }
        break
    }
    if ($lv.Parent) { Update-PmMachineTabTotals -Machine $Machine -Tab $lv.Parent }
}

function Show-PmTaskTimerForm {
    param([Parameter(Mandatory)]$Machine, [Parameter(Mandatory)]$Task)

    $taskId  = "$($Task.ItemNo)"
    $formKey = "$($Machine.MachineId)|$taskId"
    if (-not $script:pmTaskTimerForms) { $script:pmTaskTimerForms = @{} }
    if ($script:pmTaskTimerForms.ContainsKey($formKey)) {
        $existing = $script:pmTaskTimerForms[$formKey]
        if ($existing -and -not $existing.IsDisposed) {
            $existing.WindowState = 'Normal'; $existing.Activate(); $existing.BringToFront(); return
        }
    }

    $entry = $Machine.State[$taskId]
    if ((Get-PmTaskTimerState -Entry $entry) -eq 'idle') {
        Start-PmTaskTimer -Machine $Machine -Task $Task
        Refresh-PmTaskRow -Machine $Machine -Task $Task
        $entry = $Machine.State[$taskId]
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PM Task Timer — $($Task.ItemNo) $($Task.Component)"
    $form.Size = New-Object System.Drawing.Size(560, 460)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ShowInTaskbar = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "$($Task.ItemNo) — $($Task.Component)"
    $lblTitle.Location = New-Object System.Drawing.Point(12, 14)
    $lblTitle.Size = New-Object System.Drawing.Size(520, 30)
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblTitle)

    $lblMachine = New-Object System.Windows.Forms.Label
    $lblMachine.Text = "$($Machine.MachineId) — Checklist $($Machine.ChecklistNo)"
    $lblMachine.Location = New-Object System.Drawing.Point(12, 44)
    $lblMachine.Size = New-Object System.Drawing.Size(520, 20)
    $lblMachine.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblMachine.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblMachine.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $form.Controls.Add($lblMachine)

    $lblElapsed = New-Object System.Windows.Forms.Label
    $lblElapsed.Text = "0:00"
    $lblElapsed.Location = New-Object System.Drawing.Point(12, 72)
    $lblElapsed.Size = New-Object System.Drawing.Size(520, 100)
    $lblElapsed.Font = New-Object System.Drawing.Font("Consolas", 52, [System.Drawing.FontStyle]::Bold)
    $lblElapsed.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblElapsed)

    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Location = New-Object System.Drawing.Point(12, 174)
    $lblState.Size = New-Object System.Drawing.Size(520, 30)
    $lblState.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblState.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblState)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(12, 208)
    $lblInfo.Size = New-Object System.Drawing.Size(520, 44)
    $lblInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $lblInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($lblInfo)

    $pauseResumeBtn = New-Object System.Windows.Forms.Button
    $pauseResumeBtn.Text = "Pause"
    $pauseResumeBtn.Size = New-Object System.Drawing.Size(130, 44)
    $pauseResumeBtn.Location = New-Object System.Drawing.Point(120, 268)
    $pauseResumeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pauseResumeBtn.FlatAppearance.BorderSize = 0
    $pauseResumeBtn.ForeColor = [System.Drawing.Color]::White
    $pauseResumeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pauseResumeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($pauseResumeBtn)

    $finishBtn = New-Object System.Windows.Forms.Button
    $finishBtn.Text = "Finish && Log"
    $finishBtn.Size = New-Object System.Drawing.Size(140, 44)
    $finishBtn.Location = New-Object System.Drawing.Point(260, 268)
    $finishBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $finishBtn.FlatAppearance.BorderSize = 0
    $finishBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $finishBtn.ForeColor = [System.Drawing.Color]::White
    $finishBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $finishBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($finishBtn)

    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = "Reset Timer"
    $resetBtn.Size = New-Object System.Drawing.Size(140, 44)
    $resetBtn.Location = New-Object System.Drawing.Point(200, 268)
    $resetBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $resetBtn.FlatAppearance.BorderSize = 0
    $resetBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $resetBtn.ForeColor = [System.Drawing.Color]::White
    $resetBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $resetBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $resetBtn.Visible = $false
    $form.Controls.Add($resetBtn)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "You can keep using the app while this runs. Finish overwrites this task's actual time."
    $lblHint.Location = New-Object System.Drawing.Point(12, 322)
    $lblHint.Size = New-Object System.Drawing.Size(520, 22)
    $lblHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(160,170,185)
    $form.Controls.Add($lblHint)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Size = New-Object System.Drawing.Size(100, 34)
    $closeBtn.Location = New-Object System.Drawing.Point(230, 356)
    $closeBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $closeBtn.Add_Click({ $form.Close() })
    $form.Controls.Add($closeBtn)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500

    $refreshDisplay = {
        $curEntry = $Machine.State[$taskId]
        $state = Get-PmTaskTimerState -Entry $curEntry
        $sec = Get-PmTaskTimerElapsedSeconds -Entry $curEntry
        $lblElapsed.Text = Format-JobTimerElapsed -Seconds $sec

        switch ($state) {
            'running' {
                $lblState.Text = "Running"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $pauseResumeBtn.Text = "Pause"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(243,156,18)
                $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true; $resetBtn.Visible = $false
            }
            'paused' {
                $lblState.Text = "Paused"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(243,156,18)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
                $pauseResumeBtn.Text = "Resume"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true; $resetBtn.Visible = $false
            }
            'finished' {
                $lblState.Text = "Finished"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(39,174,96)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
                $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false; $resetBtn.Visible = $true
            }
            default {
                $lblState.Text = "Idle"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
                $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
                $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false; $resetBtn.Visible = $false
            }
        }

        $info = @()
        if ($curEntry -and $curEntry.TimerStartedAt) {
            try { $info += "Started: $(([DateTime]::Parse("$($curEntry.TimerStartedAt)")).ToString('HH:mm:ss'))" } catch { }
        }
        if ($curEntry -and $null -ne $curEntry.CustomTimeMin) { $info += "Actual: $($curEntry.CustomTimeMin) min" }
        $pt = 0; if ($curEntry) { [void][int]::TryParse("$($curEntry.TimerPausedTotalSecs)", [ref]$pt) }
        if ($pt -gt 0) { $info += "Paused: $([Math]::Round($pt / 60.0, 1)) min" }
        $lblInfo.Text = ($info -join "     ")
    }.GetNewClosure()

    $pauseResumeBtn.Add_Click({
        $curEntry = $Machine.State[$taskId]
        $state = Get-PmTaskTimerState -Entry $curEntry
        if ($state -eq 'running')  { Pause-PmTaskTimer  -Machine $Machine -Task $Task }
        elseif ($state -eq 'paused') { Resume-PmTaskTimer -Machine $Machine -Task $Task }
        Refresh-PmTaskRow -Machine $Machine -Task $Task
        & $refreshDisplay
    }.GetNewClosure())

    $finishBtn.Add_Click({
        $curEntry = $Machine.State[$taskId]
        $state = Get-PmTaskTimerState -Entry $curEntry
        if ($state -notin @('running','paused')) { return }
        $sec = Get-PmTaskTimerElapsedSeconds -Entry $curEntry
        $hms = Format-JobTimerElapsed -Seconds $sec
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Finish this timer?`r`n`r`nWorked time: $hms`r`n`r`nThis will OVERWRITE the task's actual time (CustomTimeMin) with the rounded minutes.",
            "Finish PM Task Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Finish-PmTaskTimer -Machine $Machine -Task $Task
        Refresh-PmTaskRow -Machine $Machine -Task $Task
        & $refreshDisplay
    }.GetNewClosure())

    $resetBtn.Add_Click({
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Reset this timer? The actual time stays at its current value.",
            "Reset PM Task Timer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Reset-PmTaskTimer -Machine $Machine -Task $Task
        Refresh-PmTaskRow -Machine $Machine -Task $Task
        & $refreshDisplay
    }.GetNewClosure())

    $timer.Add_Tick({ & $refreshDisplay }.GetNewClosure())
    $timer.Start()

    $form.Add_FormClosed({
        try { $timer.Stop(); $timer.Dispose() } catch { }
        if ($script:pmTaskTimerForms) { [void]$script:pmTaskTimerForms.Remove($formKey) }
    }.GetNewClosure())

    $script:pmTaskTimerForms[$formKey] = $form
    & $refreshDisplay
    $form.Show()
}
```

6d — Preserve timer state in Handle-PmCheckboxChanged and Handle-PmItemDoubleClick

Replace Handle-PmCheckboxChanged:

```powershell
function Handle-PmCheckboxChanged {
    param(
        [System.Windows.Forms.ListView]$Sender,
        [System.Windows.Forms.ItemCheckedEventArgs]$Event
    )
    if ($script:pmSuppressEvents) { return }
    $machine = $Sender.Tag
    if (-not $machine -or -not $machine.State) { return }
    $item = $Event.Item
    if (-not $item) { return }
    $task = $item.Tag
    if (-not $task) { return }

    $taskId = "$($task.ItemNo)"
    $entry = $machine.State[$taskId]
    if ($null -eq $entry) {
        $entry = @{
            Selected             = $false
            CustomTimeMin        = $null
            TimerState           = 'idle'
            TimerStartedAt       = $null
            TimerPausedAt        = $null
            TimerPausedTotalSecs = 0
            TimerEndedAt         = $null
        }
        $machine.State[$taskId] = $entry
    }
    $entry.Selected = [bool]$item.Checked
    Save-PmSession -Machine $machine
    Update-PmMachineTabTotals -Machine $machine -Tab $Sender.Parent
}
```

Replace Handle-PmItemDoubleClick:

```powershell
function Handle-PmItemDoubleClick {
    param(
        [System.Windows.Forms.ListView]$Sender,
        [System.EventArgs]$Event
    )
    if ($script:pmSuppressEvents) { return }
    $machine = $Sender.Tag
    if (-not $machine -or -not $machine.State) { return }
    $sel = $Sender.SelectedItems
    if ($sel.Count -eq 0) { return }
    $item = $sel[0]
    $task = $item.Tag
    if (-not $task) { return }

    $result = Show-PmTaskEditor -Machine $machine -Task $task
    if ($null -eq $result) { return }

    $taskId = "$($task.ItemNo)"
    $entry = $machine.State[$taskId]
    if ($null -eq $entry) {
        $entry = @{
            Selected             = $false
            CustomTimeMin        = $null
            TimerState           = 'idle'
            TimerStartedAt       = $null
            TimerPausedAt        = $null
            TimerPausedTotalSecs = 0
            TimerEndedAt         = $null
        }
        $machine.State[$taskId] = $entry
    }
    $entry.Selected      = [bool]$result.Selected
    $entry.CustomTimeMin = $result.CustomTimeMin
    Save-PmSession -Machine $machine

    $item.Checked = [bool]$result.Selected
    $displayMin = if ($null -ne $result.CustomTimeMin) { [int]$result.CustomTimeMin } else { [int]$task.EstTimeMin }
    $item.SubItems[6].Text = "$displayMin"

    $rowColor = Get-PmTaskRowColor -Task $task -Today (Get-Date).Date
    Set-PmTaskRowColors -Item $item -RowColor $rowColor -TimeIsOverridden ($null -ne $result.CustomTimeMin)
    Update-PmMachineTabTotals -Machine $machine -Tab $Sender.Parent
}
```

6e — Edits inside Refresh-PmMachineTabs

a) Add the Timer column and populate it — find the task ListView construction:

```powershell
        $lv.Columns.Add("", 30)          | Out-Null
        $lv.Columns.Add("Item", 70)      | Out-Null
        $lv.Columns.Add("Component", 240)| Out-Null
        $lv.Columns.Add("Power", 70)     | Out-Null
        $lv.Columns.Add("Skill", 60)     | Out-Null
        $lv.Columns.Add("Est", 50)       | Out-Null
        $lv.Columns.Add("Time", 60)      | Out-Null
        $lv.Columns.Add("Due", 90)       | Out-Null
```

Replace with (add one more column):

```powershell
        $lv.Columns.Add("", 30)          | Out-Null
        $lv.Columns.Add("Item", 70)      | Out-Null
        $lv.Columns.Add("Component", 240)| Out-Null
        $lv.Columns.Add("Power", 70)     | Out-Null
        $lv.Columns.Add("Skill", 60)     | Out-Null
        $lv.Columns.Add("Est", 50)       | Out-Null
        $lv.Columns.Add("Time", 60)      | Out-Null
        $lv.Columns.Add("Due", 90)       | Out-Null
        $lv.Columns.Add("Timer", 110)    | Out-Null   # index 8
```

b) Populate the Timer sub-item for each task — find the foreach ($task in @($m.Parsed.Tasks)) loop and, right after $it.SubItems.Add("$($task.DueDate)") | Out-Null, insert:

```powershell
            # Timer column
            $tState = Get-PmTaskTimerState -Entry $e
            $tSec = Get-PmTaskTimerElapsedSeconds -Entry $e
            $tStr = Format-JobTimerElapsed -Seconds $tSec
            switch ($tState) {
                'running'  { $timerText = "> $tStr";  $timerColor = [System.Drawing.Color]::FromArgb(52,152,219) }
                'paused'   { $timerText = "II $tStr"; $timerColor = [System.Drawing.Color]::FromArgb(243,156,18) }
                'finished' { $timerText = "OK $tStr"; $timerColor = [System.Drawing.Color]::FromArgb(39,174,96) }
                default    { $timerText = "START";    $timerColor = [System.Drawing.Color]::FromArgb(39,174,96) }
            }
            $it.SubItems.Add($timerText) | Out-Null
```

c) Style the Timer cell and register the click handler — find where the ListView is added to the tab ($tab.Controls.Add($lv)) and, before that line, insert:

```powershell
        $it.UseItemStyleForSubItems = $false
        $it.SubItems[8].ForeColor = $timerColor
        $it.SubItems[8].Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)

        # Store the ListView so timers can refresh specific rows
        if (-not $script:pmMachineListViews) { $script:pmMachineListViews = @{} }
        $script:pmMachineListViews[$m.MachineId] = $lv

        $lv.Add_MouseClick({
            param($sender, $e)
            if ($sender.Columns.Count -lt 9) { return }
            $timerColX = 0
            for ($i = 0; $i -lt 8; $i++) { $timerColX += $sender.Columns[$i].Width }
            if ($e.X -lt $timerColX) { return }
            if ($e.X -gt ($timerColX + $sender.Columns[8].Width)) { return }
            $hit = $sender.HitTest($e.Location)
            if (-not $hit.Item) { return }
            $task = $hit.Item.Tag
            $machine = $sender.Tag
            if (-not $task -or -not $machine) { return }
            Show-PmTaskTimerForm -Machine $machine -Task $task
        })
```

6f — Live ticker in Setup-PmChecklistSubTab

Find the block that ends the setup:

```powershell
    $script:pmMachines = Load-PmMachines
    Refresh-PmMachineTabs

    Write-Log "PM Checklist sub-tab setup completed."
```

Replace with:

```powershell
    $script:pmMachines = Load-PmMachines
    Refresh-PmMachineTabs

    if (-not $script:pmTaskTimerTicker) {
        $script:pmTaskTimerTicker = New-Object System.Windows.Forms.Timer
        $script:pmTaskTimerTicker.Interval = 1000
        $script:pmTaskTimerTicker.Add_Tick({
            if (-not $script:pmMachineListViews) { return }
            foreach ($machineId in $script:pmMachineListViews.Keys) {
                $lv = $script:pmMachineListViews[$machineId]
                if (-not $lv -or $lv.IsDisposed) { continue }
                $machine = $lv.Tag
                if (-not $machine -or -not $machine.State) { continue }
                foreach ($item in $lv.Items) {
                    $task = $item.Tag
                    if (-not $task) { continue }
                    $e = $machine.State["$($task.ItemNo)"]
                    if ((Get-PmTaskTimerState -Entry $e) -ne 'running') { continue }
                    $sec = Get-PmTaskTimerElapsedSeconds -Entry $e
                    $item.SubItems[8].Text = "> " + (Format-JobTimerElapsed -Seconds $sec)
                }
            }
        })
        $script:pmTaskTimerTicker.Start()
    }

    Write-Log "PM Checklist sub-tab setup completed."
```

---

Summary

Timer Where it starts Where it ends Live display On finish
Break (Change 4) Ghosted break row's "Start" cell in the breaks list "Finish & Log" or Cancel dialog Runs modeless; row shows ●; countdown turns red with overtime Breaks.json gets startTime + durationMin
Job (Change 5) Timer column in job list, or "Start Timer" in job editor "Finish & Log" in timer window Runs modeless; job list row shows > / II / OK H:MM:SS; form updates 2 Hz Jobs.json gets timerComputedHours, endTime; Get-JobHours prefers it
PM task (Change 6) Timer column in task list (auto-starts on first click) "Finish & Log" in timer window Runs modeless; task row shows > / II / OK H:MM:SS; ticker updates 1 Hz Session file's CustomTimeMin overwritten by rounded minutes; task list "Time" column updates

Overlapping and pausing: break and job timers cannot be paused (breaks are discrete, jobs just count up). PM task timer supports Pause/Resume — paused time is excluded from the final CustomTimeMin overwrite.

Persistence: all three timers survive an app restart. Breaks.json stores startTime; Jobs.json stores six timer* fields; *.session.json (per PM machine) stores six Timer* fields per task.

Re-opening a running timer: clicking its Timer/Start cell brings the existing window to the front instead of starting a new one.