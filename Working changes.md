Two syntax errors, both in Search-Parts. They came from the B5 patch you applied — the copy/paste left a stray ) and an orphaned }.

Fix 1 — line 1379, extra close paren

Find this (line ~1374–1379):

```powershell
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description,
        [string]$BookFilter = '',
        [string]$SectionFilter = ''
    ))
```

Change )) to ):

```powershell
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description,
        [string]$BookFilter = '',
        [string]$SectionFilter = ''
    )
```

That's the Unexpected token ')' at 1379.

Fix 2 — line ~1499–1500, orphaned close brace

Find the end of Search-Parts:

```powershell
    return $results
}
}
# Data retrieval logic
function Search-CrossReferenceData {
```

Delete the second } so it reads:

```powershell
    return $results
}
# Data retrieval logic
function Search-CrossReferenceData {
```

That's the Unexpected token '}' at 1499/1500. The Missing closing '}' at line 1373 is PowerShell's report location — it can't match the brace because of the )) a few lines below, so it flags the function keyword.

Verify

Quick check before running the app:

```powershell
$ast = [System.Management.Automation.Language.Parser]::ParseFile("$PWD\UI-Script.ps1", [ref]$null, [ref]$errs)
if ($errs) { $errs | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" } } else { "OK" }
```

Should print OK.

---

Two runtime issues you'll hit right after the parse fix

Not parse errors, but both block startup once the script parses:

A. $searchButton referenced before it exists (in Setup-SearchTab)

The $updateSearchEnabled closure at line ~1334 references $searchButton, but $searchButton isn't created until ~line 1360. .GetNewClosure() captures $null, then the pre-select below fires the event and writes $null.Enabled = ..., which errors.

Fix: move the enable-logic + pre-select block to after $searchButton is created (i.e. after $topPanel.Controls.Add($searchButton)). Move this entire chunk down:

```powershell
    $updateSearchEnabled = {
        $searchButton.Enabled = (
            $script:comboBook.SelectedIndex    -ge 0 -and
            $script:comboSection.SelectedIndex -ge 0
        )
    }.GetNewClosure()

    $script:comboBook.Add_SelectedIndexChanged($updateSearchEnabled)
    $script:comboSection.Add_SelectedIndexChanged($updateSearchEnabled)

    if ($script:comboBook.Items.Count -gt 0) {
        $script:comboBook.SelectedIndex = 0
    }
    $searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -and $script:comboSection.SelectedIndex -ge 0)
```

…so it sits between $topPanel.Controls.Add($searchButton) and the tooltip creation.

B. $startBtn handler in Show-PmTaskTimerForm calls the wrong timer

The Start button you added in Show-PmTaskTimerForm (line ~4130ish) has:

```powershell
$startBtn.Add_Click({
    try {
        Start-JobTimer -JobId $JobId        # <-- wrong function, wrong variable
        Refresh-JobLogList                  # <-- not relevant here
        Refresh-WorkBudget
        & $refreshDisplay
    } catch { Write-Log "Start Timer error: $($_.Exception.Message)" }
}.GetNewClosure())
```

$JobId and Start-JobTimer belong to the Job timer, not the PM task timer. $JobId is undefined here → silent null error at click time.

Replace with:

```powershell
$startBtn.Add_Click({
    try {
        $curEntry = $Machine.State[$taskId]
        if ((Get-PmTaskTimerState -Entry $curEntry) -eq 'idle') {
            Start-PmTaskTimer -Machine $Machine -Task $Task
            Refresh-PmTaskRow -Machine $Machine -Task $Task
        }
        & $refreshDisplay
    } catch { Write-Log "PM Start Timer error: $($_.Exception.Message)" }
}.GetNewClosure())
```

$Machine, $Task, $taskId, and $refreshDisplay are all already captured correctly in that form — the only wrong lines are the three inside the try block.

The analogous Start-button handler in Show-JobTimerForm (if you're adding it there too per A6) is correct as written — it uses $JobId, Start-JobTimer, Refresh-JobLogList, Refresh-WorkBudget, all of which exist in that scope.