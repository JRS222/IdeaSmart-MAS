Two things to fix before the harness will run, and one warning

---

Part 1 — The harness (Test harness for emars-fetch-layer.ps1)

Three edits, all simple path swaps. Open the file in Notepad or VSCode.

Line 6 — point at wherever you have the probe output from yesterday:

```powershell
$probeDir = 'C:\Users\YYQTP0\Desktop\emars-pm-probe-20261024-141529'
```

If the folder name is different (later timestamp), match it exactly — the harness reads 05-getAllChecklists-ss01.xml from inside that folder. If your probe folder doesn't have that filename, check its contents first:

```powershell
Get-ChildItem 'C:\Users\YYQTP0\Desktop\emars-pm-probe-*' | Where-Object { $_.Name -like '*getAllChecklists*' } | Select FullName
```

It needs to say 05-getAllChecklists-ss01.xml. If it's 05-getAllChecklists.xml (no -ss01) or different, either rename the file or change the filename in the harness on line 20.

Line 87 — point at where you have the layer file:

```powershell
. "C:\Users\YYQTP0\Desktop\emars-fetch-layer.ps1"
```

Line ~8 — the archive output goes to $env:TEMP which is fine, no change needed.

Then run it:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\YYQTP0\Desktop\Test harness for emars-fetch-layer.ps1"
```

What to expect if it works:

```
PASS: machine list has 36 rows (got 36)
PASS: machine DBCS 16 resolves (ptr=827433 class=CK)
PASS: site name from list
PASS: open checklist found (trackno=37900603)
... (30 more PASS lines)
ALL TESTS PASSED
Archive written to: C:\Users\YYQTP0\AppData\Local\Temp\pm_emars_test_<guid>
```

What failures mean:

Failure Meaning
Cannot find path ... emars-fetch-layer.ps1 Line 87 path is wrong
Cannot find path ... 05-getAllChecklists-ss01.xml Line 6 path is wrong, OR that XML was never saved yesterday
machine list has 36 rows (got 0) The XML file is missing/empty — regenerate it via the probe
Any other FAIL: The layer has a real bug; paste the failing line back

The harness is offline — every Invoke-PmHttpGet call returns canned XML from disk. It does not contact eMARS. So ALL TESTS PASSED proves the layer's logic is correct, but says nothing about whether eMARS responds from YYQTP0.

---

Part 2 — The second file is not a test

Show-PmFetchEmarsDialog (the whole second block) is a production UI function, not a test. Running it standalone will fail immediately — it references $script:pmDatePicker, Get-WeeklyWorksheet, Get-EmarsSetting, Set-ListViewStyle, New-Button, Get-EmarsMachineList, Import-PmChecklistFromEmars — all of which only exist inside UI-Script.ps1 (and the fetch layer).

If you try to run it in a fresh PowerShell session you'll get The term 'Get-EmarsSetting' is not recognized on line 1.

What to do with it: don't run it. Save it, and when you're ready to integrate, paste both Get-WorksheetPmRows and Show-PmFetchEmarsDialog into UI-Script.ps1 (anywhere before Setup-PmChecklistSubTab), and add the button that calls it. That's a separate step.

---

Part 3 — What the harness doesn't tell you (and how to find out)

Because the harness is offline, "ALL TESTS PASSED" is not the same as "works from my desk". If you want to know whether eMARS actually responds from YYQTP0 — that's a completely different five-minute check:

```powershell
$base = 'http://emarssu3.eng.usps.gov/pemars'
$site = '345'

# Live test — no harness, no canned data
"=== Live eMARS smoke test ==="

try {
    $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 15 `
        -Uri "$base/ecbm.ecbm_report_j.getAllChecklists?pSiteID=$site&pSubsiteID=01&pAcronym=DBCS"
    $rows = ([regex]::Matches($r.Content, '<ROW>')).Count
    "getAllChecklists: OK  ($($r.Content.Length) chars, $rows row(s))"
} catch {
    "getAllChecklists: FAILED — $($_.Exception.Message)"
}

try {
    $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 15 `
        -Uri "$base/ecbm.ecbm_report_j.getOpenlist?p_machineid=827433&pSiteID=$site"
    $rows = ([regex]::Matches($r.Content, '<ROW>')).Count
    "getOpenlist:      OK  ($($r.Content.Length) chars, $rows row(s))"
} catch {
    "getOpenlist:      FAILED — $($_.Exception.Message)"
}

try {
    $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 15 `
        -Uri "$base/ecbm.ecbm_checklist.getAChecklist?p_machinetype=DBCS&p_classcode=CK&p_trackno=37900603&p_machineid=827433&p_site=$site"
    $rows = ([regex]::Matches($r.Content, '<ROW>')).Count
    "getAChecklist:    OK  ($($r.Content.Length) chars, $rows row(s))"
} catch {
    "getAChecklist:    FAILED — $($_.Exception.Message)"
}
```

If all three say OK with non-zero rows, the approach is confirmed against live eMARS from your workstation — the harness's green output then means "logic correct + data path confirmed". If any says FAILED, paste the error back — that tells us about firewall/URL differences between JR's machine and yours, which is exactly the thing the harness can't check.

getAChecklist with p_trackno=37900603 may return 0 rows if that specific checklist has closed since yesterday — that's not a failure, just stale data. getOpenlist will tell you what trackno is open now for DBCS 16.

---

What to send back

Just the harness output. Either:

· ALL TESTS PASSED → harness confirms the layer is sound, integration is safe
· A FAIL: line → paste it, tells me exactly what to fix
· A Cannot find path error → paste it, tells me which path edit is still wrong

If you also run the live smoke test and want to close the loop, paste its output too. But get the harness running first — that's the fast feedback, and it doesn't need the network at all.