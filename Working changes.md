Diagnosis

Issue 1 — Closing the timer window closes the whole app

The tick handler runs every 500 ms and updates labels ($lblElapsed.Text = …). When you click the X, WinForms starts disposing the form before FormClosed fires. A tick that was already queued can still fire between FormClosing and FormClosed, try to set .Text on an already-disposed label, and throw ObjectDisposedException. That exception escapes into the WinForms message loop, the default Application.ThreadException handler shows the "unhandled exception — Quit / Continue" dialog, and clicking Quit calls Application.Exit() — taking the main form with it.

Two missing safeguards make this possible:

1. The timer is stopped in FormClosed (too late) instead of FormClosing.
2. $refreshDisplay never checks $form.IsDisposed before touching controls.

Issue 2 — No way to stop a running job timer

The form only offers Pause (temporarily halts elapsed time) and Finish & Log (records the time). There is no "cancel / abandon" path — once a timer is running or paused, the only way out is to finish it (recording hours you don't want) and then Reset. Reset is only visible in the finished state.

The fix for both is below. Apply the same set of changes to Show-JobTimerForm and Show-PmTaskTimerForm.

---

Fix

Change 1 — Stop the timer before the form disposes

Find (near the end of Show-JobTimerForm):

```powershell
$timer.Add_Tick({ & $refreshDisplay }.GetNewClosure())
$timer.Start()

$form.Add_FormClosed({
    try { $timer.Stop(); $timer.Dispose() } catch { }
    if ($script:jobTimerForms) { [void]$script:jobTimerForms.Remove($JobId) }
}.GetNewClosure())
```

Replace with:

```powershell
$timer.Add_Tick({ & $refreshDisplay }.GetNewClosure())
$timer.Start()

$form.Add_FormClosing({
    try { $timer.Stop() } catch { }
}.GetNewClosure())

$form.Add_FormClosed({
    try { $timer.Dispose() } catch { }
    if ($script:jobTimerForms) { [void]$script:jobTimerForms.Remove($JobId) }
}.GetNewClosure())
```

For the PM task timer, change the removal line to $script:pmTaskTimerForms.Remove($formKey).

Change 2 — Guard the tick handler against disposed controls

Find the top of $refreshDisplay in both timer forms:

```powershell
$refreshDisplay = {
    $currentJob = Get-JobById -JobId $JobId
    ...
```

Change to:

```powershell
$refreshDisplay = {
    if ($form.IsDisposed) { return }
    $currentJob = Get-JobById -JobId $JobId
    ...
```

(For the PM task timer: if ($form.IsDisposed) { return } followed by $curEntry = $Machine.State[$taskId].)

That alone kills the "app closes" bug. The remaining changes add the missing Cancel path.

Change 3 — Add a Cancel Timer button (job timer)

Insert this after $resetBtn is created (just before the $lblHint label):

```powershell
$cancelTimerBtn = New-Object System.Windows.Forms.Button
$cancelTimerBtn.Text = "Cancel Timer"
$cancelTimerBtn.Size = New-Object System.Drawing.Size(130, 44)
$cancelTimerBtn.Location = New-Object System.Drawing.Point(395, 246)
$cancelTimerBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cancelTimerBtn.FlatAppearance.BorderSize = 0
$cancelTimerBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
$cancelTimerBtn.ForeColor = [System.Drawing.Color]::White
$cancelTimerBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$cancelTimerBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$cancelTimerBtn.Visible = $false
$form.Controls.Add($cancelTimerBtn)

$cancelTimerBtn.Add_Click({
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Cancel this timer?`r`n`r`nThe timer will reset to idle. Elapsed time will NOT be recorded.",
        "Cancel Timer",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Reset-JobTimer -JobId $JobId
    & $refreshDisplay
}.GetNewClosure())
```

Change 4 — Show/hide Cancel based on state

Inside $refreshDisplay’s switch ($state) block, add $cancelTimerBtn.Visible:

```powershell
'running' {
    $lblState.Text = "Running"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $pauseResumeBtn.Text = "Pause"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(243,156,18)
    $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $true
}
'paused' {
    $lblState.Text = "Paused"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(243,156,18)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
    $pauseResumeBtn.Text = "Resume"; $pauseResumeBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $pauseResumeBtn.Visible = $true; $finishBtn.Visible = $true
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $true
}
'finished' {
    $lblState.Text = "Finished"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false
    $resetBtn.Visible = $true; $cancelTimerBtn.Visible = $false
}
default {
    $lblState.Text = "Idle"; $lblState.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $lblElapsed.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
    $pauseResumeBtn.Visible = $false; $finishBtn.Visible = $false
    $resetBtn.Visible = $false; $cancelTimerBtn.Visible = $false
}
```

Change 5 — Same treatment for the PM task timer

Inside Show-PmTaskTimerForm, use the same pattern:

· $cancelTimerBtn positioned at (395, 268) size (150, 44) (the form is 560 wide).
· Click handler calls Reset-PmTaskTimer -Machine $Machine -Task $Task instead of Reset-JobTimer.
· Update the four switch branches in its $refreshDisplay to toggle $cancelTimerBtn.Visible as above.

The important part is the same: with Reset-PmTaskTimer already defined in the script, the Cancel path is a one-liner for both forms.

---

Why this works

Bug Root cause Fix
Program exits when timer window closes Queued tick fires between FormClosing and FormClosed; ObjectDisposedException bubbles to Application.ThreadException → Quit → Application.Exit() Stop timer in FormClosing and bail out of $refreshDisplay if $form.IsDisposed
Can't stop a running timer Form only had Pause (temporary) and Finish (logs hours) Added explicit Cancel Timer button (only visible in running / paused) that calls Reset-JobTimer / Reset-PmTaskTimer

---

Safety-net fallback (optional)

If you still see unexplained exits on a specific machine, the root cause is almost always the same message-loop issue, but for extra insurance you can add this once at the very top of the script, before Show-MainForm:

```powershell
[System.Windows.Forms.Application]::SetUnhandledExceptionMode(
    [System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($sender, $e)
    Write-Log "WinForms unhandled: $($e.Exception.Message)"
    $e.Exception.ToString() | Out-File -Append -FilePath "UI.log"
    [System.Windows.Forms.MessageBox]::Show(
        "A UI error occurred: $($e.Exception.Message)`r`n`r`nSee UI.log.",
        "UI Error", "OK", "Error")
})
```

With CatchException, WinForms stops treating the exception as fatal and just displays your handler's message — the app keeps running even if a rogue handler throws. Combined with the primary fix, this makes the app essentially bullet-proof against the close cascade.