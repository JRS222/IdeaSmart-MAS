#Requires -Version 5.1
<#
.SYNOPSIS
    Fetches ECBM PM checklists directly from the Oracle procedures,
    bypassing the module shell.
.PARAMETER Acronym
    Machine type (e.g. DBCS, CIOSS, APPS)
.PARAMETER Site
    eMARS site ID
.PARAMETER Subsite
    ECBM subsite ID (leave blank to try without)
#>
param(
    [string]$Acronym = 'DBCS',
    [string]$Site    = '345',
    [string]$Subsite = '',
    [string]$Origin  = 'http://emarssu3.eng.usps.gov',
    [string]$OutDir  = (Join-Path $env:USERPROFILE ("Desktop\ecbm-fetch-" + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

function Say {
    param([string]$m, [string]$c = 'Gray')
    Write-Host $m -ForegroundColor $c
}

function Fetch {
    param([string]$url)
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 25 -ErrorAction Stop
        return [PSCustomObject]@{ Ok = $true; Status = $r.StatusCode; Body = $r.Content; Url = $url }
    } catch {
        return [PSCustomObject]@{
            Ok = $false; Status = -1; Body = ''
            Url = $url; Error = $_.Exception.Message
        }
    }
}

function Show-Snippet {
    param([string]$s, [int]$n = 400)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $len = [Math]::Min($n, $s.Length)
    return $s.Substring(0, $len)
}

$base = "$Origin/pemars"

Say "============================================================" 'Cyan'
Say "  ECBM CHECKLIST FETCH — DIRECT PROC TEST" 'Cyan'
Say "============================================================" 'Cyan'
Say ""
Say ("  Site:    {0}" -f $Site)
Say ("  Acronym: {0}" -f $Acronym)
Say ("  Subsite: {0}" -f $(if ($Subsite) { $Subsite } else { '(blank)' }))
Say ("  Output:  {0}" -f $OutDir)
Say ""

# ---------------------------------------------------------------
# 1. getAllChecklists
# ---------------------------------------------------------------
Say "--- 1. getAllChecklists ---" 'Cyan'
$url1 = "$base/ecbm.ecbm_report_j.getAllChecklists?pSiteID=$Site&pSubsiteID=$Subsite&pAcronym=$Acronym"
Say ("  GET {0}" -f $url1)
$r1 = Fetch $url1

if (-not $r1.Ok) {
    Say ("  FAILED: {0}" -f $r1.Error) 'Red'
    Say '  Trying without subsite param...' 'Yellow'
    $url1b = "$base/ecbm.ecbm_report_j.getAllChecklists?pSiteID=$Site&pAcronym=$Acronym"
    Say ("  GET {0}" -f $url1b)
    $r1 = Fetch $url1b
    if ($r1.Ok) { $url1 = $url1b }
}

if ($r1.Ok) {
    Say ("  status: {0}  length: {1}" -f $r1.Status, $r1.Body.Length) 'Green'
    $r1.Body | Set-Content (Join-Path $OutDir '01-getAllChecklists.xml') -Encoding UTF8

    $machines = [regex]::Matches($r1.Body, '(?is)<ROW>(.*?)</ROW>')
    Say ("  {0} machine row(s)" -f $machines.Count) 'Green'

    $machineList = @()
    foreach ($m in $machines) {
        $fields = [regex]::Matches($m.Groups[1].Value, '(?is)<([A-Z0-9_]+)>(.*?)</\1>')
        $row = [ordered]@{}
        foreach ($f in $fields) { $row[$f.Groups[1].Value] = $f.Groups[2].Value }
        $machineList += [PSCustomObject]$row
    }

    # Show first few machines with whatever column names exist
    if ($machineList.Count -gt 0) {
        Say "  Sample machine rows (first 3):" 'Cyan'
        $cols = $machineList[0].PSObject.Properties.Name
        Say ("    Columns: {0}" -f ($cols -join ', '))
        foreach ($row in ($machineList | Select-Object -First 3)) {
            $vals = ($cols | ForEach-Object { "$_=$($row.$_)" }) -join ' | '
            Say ("    {0}" -f $vals) 'DarkGray'
        }
    }
} else {
    Say ("  FAILED: {0}" -f $r1.Error) 'Red'
    $machineList = @()
}

# ---------------------------------------------------------------
# 2. getChecklists2 for each machine
# ---------------------------------------------------------------
Say ""
Say "--- 2. getChecklists2 (per machine) ---" 'Cyan'

$allChecklists = @()

if ($machineList.Count -gt 0) {
    # Try to find the machine ptr field — could be named differently
    $sample = $machineList[0]
    $ptrField = $null
    foreach ($name in $sample.PSObject.Properties.Name) {
        if ($name -match '(?i)ptr|id|machine') { $ptrField = $name; break }
    }
    if (-not $ptrField) { $ptrField = $sample.PSObject.Properties.Name[0] }

    Say ("  Using field '{0}' as machine ptr" -f $ptrField) 'Gray'

    $tried = 0
    foreach ($machine in $machineList) {
        if ($tried -ge 5) { Say '  (stopping after 5 machines)' 'DarkGray'; break }
        $tried++

        $ptr = $machine.$ptrField
        if ([string]::IsNullOrWhiteSpace($ptr)) { continue }

        $url2 = "$base/ecbm.ecbm_report_j.getChecklists2?p_machineid=$ptr&pSiteID=$Site"
        Say ("  GET {0}" -f $url2)
        $r2 = Fetch $url2

        if ($r2.Ok) {
            $rows2 = [regex]::Matches($r2.Body, '(?is)<ROW>(.*?)</ROW>')
            Say ("    {0} checklist(s)" -f $rows2.Count) 'Green'

            foreach ($m2 in $rows2) {
                $fields2 = [regex]::Matches($m2.Groups[1].Value, '(?is)<([A-Z0-9_]+)>(.*?)</\1>')
                $row2 = [ordered]@{}
                foreach ($f2 in $fields2) { $row2[$f2.Groups[1].Value] = $f2.Groups[2].Value }
                $row2['_machinePtr'] = $ptr
                $row2['_acronym']    = $Acronym
                $allChecklists += [PSCustomObject]$row2
            }
        } else {
            Say ("    FAILED: {0}" -f $r2.Error) 'DarkGray'
        }
    }
} else {
    Say '  (skipped — no machines from step 1)' 'DarkGray'
}

if ($allChecklists.Count -gt 0) {
    Say ""
    Say ("  Total checklists found across all machines: {0}" -f $allChecklists.Count) 'Green'
    Say "  Sample checklist rows (first 5):" 'Cyan'
    $ckcols = $allChecklists[0].PSObject.Properties.Name
    Say ("    Columns: {0}" -f ($ckcols -join ', '))
    foreach ($row in ($allChecklists | Select-Object -First 5)) {
        $vals = ($ckcols | ForEach-Object { "$_=$($row.$_)" }) -join ' | '
        Say ("    {0}" -f $vals) 'DarkGray'
    }
}

# ---------------------------------------------------------------
# 3. getChecklistDetailsForIframe — the actual popup HTML
# ---------------------------------------------------------------
Say ""
Say "--- 3. getChecklistDetailsForIframe (the actual checklist HTML) ---" 'Cyan'

if ($allChecklists.Count -eq 0) {
    Say '  SKIP: no checklists found from step 2.' 'Yellow'
    Say ''
    Say '  Workaround: try a direct fetch if you know a checklist number:' 'Yellow'
    Say '  $detailUrl = "$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=<YOUR_CHECKLIST_NO>&p_site=345&p_fac=345&p_mtype=DBCS&p_mno=<equip>&p_classcode=<class>"'
} else {
    # Build field names from a sample checklist row — try to identify the checklist number field
    $sample = $allChecklists[0]
    $checklistNoField = $null
    foreach ($name in $sample.PSObject.Properties.Name) {
        if ($name -match '(?i)checklist|trackno|track_no|checklist_no|checklistno') {
            $checklistNoField = $name; break
        }
    }

    if (-not $checklistNoField) {
        Say "  Could not identify checklist number field." 'Yellow'
        Say ("  Available: {0}" -f ($sample.PSObject.Properties.Name -join ', ')) 'Yellow'
    } else {
        Say ("  Using field '{0}' as checklist number" -f $checklistNoField) 'Gray'

        # Try to fetch first 3 checklists
        $attempts = 0
        foreach ($ck in $allChecklists) {
            if ($attempts -ge 3) { break }
            $attempts++

            $trackno = $ck.$checklistNoField
            if ([string]::IsNullOrWhiteSpace($trackno)) { continue }

            # The JS also passes p_fac, p_mtype, p_mno, p_classcode. We'll try:
            #  p_fac = site (345) — best guess; falls back to empty if needed
            #  p_mtype, p_mno, p_classcode = whatever we have in the checklist row
            $fac      = $Site
            $mtype    = $Acronym
            $mno      = if ($ck.PSObject.Properties['machineno']) { $ck.machineno }
                        elseif ($ck.PSObject.Properties['MachineNo']) { $ck.MachineNo }
                        else { '' }
            $classcd  = if ($ck.PSObject.Properties['classcode']) { $ck.classcode }
                        elseif ($ck.PSObject.Properties['ClassCode']) { $ck.ClassCode }
                        else { '' }

            $detailUrl = "$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe" +
                         "?p_trackno=$trackno" +
                         "&p_site=$Site" +
                         "&p_fac=$fac" +
                         "&p_mtype=$mtype" +
                         "&p_mno=$mno" +
                         "&p_classcode=$classcd"

            Say ("  GET {0}" -f $detailUrl)
            $r3 = Fetch $detailUrl

            if ($r3.Ok) {
                $isChecklist = ($r3.Body -match '(?i)Checklist No\.' -and $r3.Body -match '(?i)Task Statement')
                $suffix = if ($isChecklist) { ' CHECKLIST ✓' } else { ' (not a checklist)' }
                $color  = if ($isChecklist) { 'Green' } else { 'Yellow' }
                Say ("    status: {0}  length: {1}{2}" -f $r3.Status, $r3.Body.Length, $suffix) $color

                $safeName = "03-checklist-$trackno.html"
                $r3.Body | Set-Content (Join-Path $OutDir $safeName) -Encoding UTF8
                Say ("    saved: {0}" -f $safeName) 'DarkGray'

                if ($isChecklist) {
                    Say ""
                    Say "  ✓ SUCCESS — checklist HTML fetched cookieless!" 'Green'
                    Say "  Next: import this into the app with Import-PmChecklist" 'Green'
                    break
                }
            } else {
                Say ("    FAILED: {0}" -f $r3.Error) 'DarkGray'
            }
        }
    }
}

# ---------------------------------------------------------------
# Summary
# ---------------------------------------------------------------
Say ""
Say "============================================================" 'Cyan'
Say "  DONE" 'Cyan'
Say "============================================================" 'Cyan'
Say ("  Files saved to: {0}" -f $OutDir)
Say ""
Say '  What to send back:' 'Cyan'
Say '    - 01-getAllChecklists.xml (first 500 chars of console output is fine)'
Say '    - The sample "Machine rows" and "checklist rows" printed above'
Say '    - Whether step 3 says CHECKLIST ✓ or (not a checklist)'