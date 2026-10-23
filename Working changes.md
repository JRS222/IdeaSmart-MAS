Here are the four blocks that need to change. For each I show the exact current text from your script, followed by the exact replacement. All other callers of Test-JobFields (Send-JobToHistorian, Send-AllReadyJobs) stay unchanged — that's intentional, so a running timer still can't be sent.

---

Block 1 — Test-JobFields

Current

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

Replace with

```powershell
function Test-JobFields {
    # Returns a string[] of validation errors (empty array = valid).
    # -AllowRunningTimer: permits a live timer (no endTime yet).  Used by the
    # editor's Save and Start/Finish buttons.  NOT used by the send-to-historian
    # paths, so a still-running job is still correctly skipped there.
    param(
        $Job,
        [switch]$AllowRunningTimer
    )

    $errs = @()
    if ($null -eq $Job) { return @('No job to validate.') }

    $m = Format-MachineId "$($Job.machineId)"
    if ($null -eq $m) { $errs += 'Machine ID must look like "DBCS 49".' }

    $timerState   = Get-JobTimerState -Job $Job
    $timerRunning = $timerState -in @('running', 'paused')

    if ([string]::IsNullOrWhiteSpace($Job.startTime)) {
        $errs += 'Start time is required.'
    }

    if ([string]::IsNullOrWhiteSpace($Job.endTime)) {
        # Only enforce endTime when the caller doesn't explicitly allow a
        # running timer.  This is the actual fix for the count-up workflow.
        if (-not ($AllowRunningTimer -and $timerRunning)) {
            $errs += 'End time is required.'
        }
    } else {
        $h = Get-JobHours $Job
        if ($null -ne $h -and $h -le 0) {
            $errs += 'End time must be after start time.'
        }
    }

    $h = Get-JobHours $Job
    $needsDetail = ($null -ne $h -and $h -gt 0.5) -or (@($Job.parts).Count -gt 0)
    if ($needsDetail) {
        if ([string]::IsNullOrWhiteSpace($Job.description)) { $errs += 'Description required (>30 min or parts used).' }
        if ([string]::IsNullOrWhiteSpace($Job.workOrderNo)) { $errs += 'Work order # required (>30 min or parts used).' }
    }

    return ,$errs
}
```

---

Block 2 — Button relabel in Show-JobEditor

Current

```powershell
    # When the timer is active, the Start Timer button becomes Stop Timer so the
    # timer can always be stopped right here in the editor.
    if ($__timerStateForEditor -in @('running','paused')) {
        $startTimerBtn.Text = "Stop Timer"
        $startTimerBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    }
```

Replace with

```powershell
    # When the timer is active, the Start Timer button becomes Finish Timer so
    # the timer can be properly closed (end time recorded, hours computed)
    # from the editor.  Discarding is done via Cancel Timer on the popup.
    if ($__timerStateForEditor -in @('running','paused')) {
        $startTimerBtn.Text = "Finish Timer"
        $startTimerBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    }
```

Colour flips red → green because the action is now productive, not destructive.

---

Block 3 — $startTimerBtn.Add_Click in Show-JobEditor

Current

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

Replace with

```powershell
    $startTimerBtn.Add_Click({
        # If the timer is live, this button FINISHES it: records end time,
        # computes hours, sets timerState='finished'.  startTime is preserved.
        # Use Cancel Timer on the timer popup to discard.
        if ((Get-JobTimerState -Job $Job) -in @('running','paused')) {
            $sec = Get-JobTimerElapsedSeconds -Job $Job
            $hms = Format-JobTimerElapsed -Seconds $sec
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "Finish this timer?`r`n`r`nWorked time: $hms`r`n`r`nThe end time will be recorded and the computed hours used for this job.",
                "Finish Timer",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            Finish-JobTimer -JobId "$($Job.id)"
            $form.Tag = [PSCustomObject]@{ Action = 'stop' }
            $form.Close()
            return
        }

        # Start branch: build the candidate in its POST-start state (running,
        # no endTime) BEFORE validating so Test-JobFields doesn't reject a
        # count-up timer whose endTime is legitimately not yet known.
        $now = Get-Date
        $candidate = [PSCustomObject]@{
            id              = $Job.id
            kind            = $Job.kind
            machineId       = $txtMachine.Text.Trim()
            startTime       = $now.ToString('HH:mm')
            endTime         = $null
            hoursOverride   = $Job.hoursOverride
            description     = $txtDesc.Text.Trim()
            workOrderNo     = $txtWo.Text.Trim()
            parts           = $script:__je_parts
            createdAt       = $Job.createdAt
            sentToHistorian = $Job.sentToHistorian
            sentAt          = $Job.sentAt
            timerState              = 'running'
            timerStartedAt          = $now.ToString("o")
            timerPausedAt           = $null
            timerPausedTotalSeconds = 0
            timerEndedAt            = $null
            timerComputedHours      = $null
        }

        $errs = Test-JobFields -Job $candidate -AllowRunningTimer
        if ($errs.Count -gt 0) { $lblVal.Text = "! " + $errs[0]; return }
        $lblVal.Text = ''

        $m = Format-MachineId $candidate.machineId
        $candidate.machineId = $m.Canonical
        if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

        $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate; StartTimer = $true }
        $form.Close()
    })
```

---

Block 4 — $saveBtn.Add_Click in Show-JobEditor

Current

```powershell
    $saveBtn.Add_Click({
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
        if ($errs.Count -gt 0) {
            $lblVal.Text = "! " + $errs[0]
            return
        }
        $lblVal.Text = ''

        $m = Format-MachineId $candidate.machineId
        $candidate.machineId = $m.Canonical
        if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

        $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate }
        $form.Close()
    })
```

Replace with

```powershell
    $saveBtn.Add_Click({
        # When a timer is live, the datetime pickers have been disabled and
        # endTime is legitimately unknown.  Preserve the real startTime and
        # leave endTime null so the picker's "now" fallback doesn't fabricate
        # a bogus end time on every Save.
        $timerActiveForSave = (Get-JobTimerState -Job $Job) -in @('running','paused')

        $candidate = [PSCustomObject]@{
            id              = $Job.id
            kind            = $Job.kind
            machineId       = $txtMachine.Text.Trim()
            startTime       = if ($timerActiveForSave -and -not [string]::IsNullOrWhiteSpace($Job.startTime)) {
                                  $Job.startTime
                              } else {
                                  $dtpStart.Value.ToString('HH:mm')
                              }
            endTime         = if ($timerActiveForSave) { $null } else { $dtpEnd.Value.ToString('HH:mm') }
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

        $errs = Test-JobFields -Job $candidate -AllowRunningTimer
        if ($errs.Count -gt 0) {
            $lblVal.Text = "! " + $errs[0]
            return
        }
        $lblVal.Text = ''

        $m = Format-MachineId $candidate.machineId
        $candidate.machineId = $m.Canonical
        if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

        $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate }
        $form.Close()
    })
```

---

What you get after applying all four

Action Before After
Click timer column in the list to start Works Works (unchanged)
Open editor while timer running → click Stop Timer Wipes startTime, endTime, computed hours Click Finish Timer → records endTime + hours, startTime preserved
Click Start Timer in editor Intermittently fails validation when pickers landed on the same minute Always succeeds; candidate is built in its running state before validating
Click Save in editor while a timer is running Overwrites the picker's fallback "now" into endTime (pollutes totals) Leaves endTime null; startTime preserved
Cancel Timer on the timer popup Discards (by design, with confirmation) Unchanged
Send all ready / Send Checklist with a running timer Skips the running job (correct) Still skips (unchanged) — Send-JobToHistorian doesn't pass -AllowRunningTimer

The three send/report callers of Test-JobFields — Send-JobToHistorian, Send-AllReadyJobs, and any future use that omits the switch — retain the old "must have an end time" behaviour, so you can't accidentally push an unfinished timer to the Historian.