#Requires -Version 5.1
<#
.SYNOPSIS
    eMARS PM checklist discovery — single cohesive probe.
.DESCRIPTION
    Runs every test we've built up so far plus new probes for the specific
    "Task Statement and Instruction" checklist format your parser expects.

    Read-only. No writes to eMARS. Safe to re-run.
.PARAMETER UserId
    eMARS user ID (from the p_user= param in your home page URL)
.PARAMETER SiteId
    eMARS site ID  (from the p_site= param)
.PARAMETER Acronym
    Machine type to probe (DBCS/CIOSS/APPS/...)
.PARAMETER MachinePtr
    machineid of one DBCS machine (default: 827433 = DBCS 16)
.PARAMETER TrackNo
    A known checklist number to probe against
.PARAMETER OutDir
    Where to save output (default: Desktop\emars-pm-probe-<timestamp>)
#>
[CmdletBinding()]
param(
    [string]$UserId     = '1601089',
    [string]$SiteId     = '345',
    [string]$Acronym    = 'DBCS',
    [string]$MachinePtr = '827433',
    [string]$TrackNo    = '37900603',
    [string]$Origin     = 'http://emarssu3.eng.usps.gov',
    [string]$OutDir     = (Join-Path $env:USERPROFILE ("Desktop\emars-pm-probe-" + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------------------------
# Setup
# ---------------------------------------------------------------
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$reportPath = Join-Path $OutDir 'PROBE-REPORT.txt'
$script:lines = New-Object System.Collections.ArrayList
$script:findings = New-Object System.Collections.ArrayList

$base = "$Origin/pemars"

# Format markers — if any of these appear in a response, it's the checklist format your parser expects
$magicMarkers = @(
    'Task Statement and Instruction'
    'Est for All Open Tasks'
    'Est for Tasks Due Today'
    'Power Off'
    'Power On'
    'Min Skill'
    'Bypass Code'
)

function Say {
    param([string]$msg = '', [string]$color = 'Gray')
    Write-Host $msg -ForegroundColor $color
    [void]$script:lines.Add($msg)
}

function Save-Blob {
    param([string]$name, [string]$content)
    if ([string]::IsNullOrEmpty($content)) { return }
    $p = Join-Path $OutDir $name
    [System.IO.File]::WriteAllText($p, $content, (New-Object System.Text.UTF8Encoding($false)))
}

function Test-Markers {
    param([string]$content)
    if ([string]::IsNullOrEmpty($content)) { return $null }
    $hit = @()
    foreach ($m in $magicMarkers) {
        if ($content -match [regex]::Escape($m)) { $hit += $m }
    }
    if ($hit.Count -eq 0) { return $null }
    return $hit
}

function Fetch {
    param(
        [Parameter(Mandatory)][string]$url,
        [string]$label = '',
        [switch]$Quiet
    )
    $obj = [PSCustomObject]@{
        Url     = $url
        Ok      = $false
        Status  = -1
        Length  = 0
        Content = $null
        Error   = $null
        Markers = $null
        Label   = $label
    }
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
        $obj.Ok      = $true
        $obj.Status  = $r.StatusCode
        $obj.Length  = $r.Content.Length
        $obj.Content = $r.Content
        $obj.Markers = Test-Markers -content $r.Content
        if (-not $Quiet) {
            if ($obj.Markers) {
                Say ("        OK  {0,7:N0} chars  ***MARKER:{1}***" -f $obj.Length, ($obj.Markers -join '|')) 'Green'
            } else {
                Say ("        OK  {0,7:N0} chars" -f $obj.Length) 'Gray'
            }
        }
    } catch {
        $obj.Error = $_.Exception.Message
        if (-not $Quiet) { Say ("        FAIL  {0}" -f $obj.Error) 'DarkGray' }
    }
    return $obj
}

function Scan-Js-For-Urls {
    param([string]$js, [string]$label)
    if ([string]::IsNullOrEmpty($js)) { return }

    $patterns = @(
        '(?:ecbm|em_[a-z_]+)\.[a-z_]+_[a-z_]+\.[a-zA-Z0-9_]+'
        'getChecklist[a-zA-Z0-9_]*'
        'getPM[a-zA-Z0-9_]*'
        'printChecklist[a-zA-Z0-9_]*'
        'getTask[a-zA-Z0-9_]*'
        'openSwilWin\s*\(\s*[''"][A-Z0-9_]+[''"]'
        'p_checklist[a-zA-Z0-9_]*'
        'p_trackno'
    )

    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($pat in $patterns) {
        $ms = [regex]::Matches($js, $pat, 'IgnoreCase')
        foreach ($m in $ms) { $hits.Add($m.Value) }
    }

    if ($hits.Count -eq 0) {
        Say ("  [{0}] no URL patterns" -f $label) 'DarkGray'
        return
    }

    Say ("  [{0}] {1} distinct hit(s):" -f $label, ($hits | Sort-Object -Unique).Count) 'Cyan'
    $hits | Sort-Object -Unique | ForEach-Object {
        Say ("    {0}" -f $_) 'DarkGray'
    }

    # Also dump lines mentioning checklist-related tokens
    $lines = $js -split "`r?`n"
    $i = 0
    $n = 0
    Say ("  [{0}] lines mentioning checklist/PM/task URLs:" -f $label) 'Cyan'
    foreach ($ln in $lines) {
        $i++
        if ($ln -match '(?i)checklist|pm_|task.*(?:url|proc|details)|print') {
            Say ("    [{0,5}] {1}" -f $i, $ln.Trim()) 'DarkGray'
            $n++
            if ($n -ge 20) {
                Say ("    ... more suppressed (see saved file)") 'DarkGray'
                break
            }
        }
    }
    if ($n -eq 0) { Say '    (none)' 'DarkGray' }
}

# ---------------------------------------------------------------
# Banner
# ---------------------------------------------------------------
Say '============================================================' 'Cyan'
Say '  eMARS PM CHECKLIST PROBE — CONSOLIDATED' 'Cyan'
Say '============================================================' 'Cyan'
Say ''
Say ("  Origin:     {0}" -f $Origin)
Say ("  UserId:     {0}" -f $UserId)
Say ("  SiteId:     {0}" -f $SiteId)
Say ("  Acronym:    {0}" -f $Acronym)
Say ("  MachinePtr: {0}" -f $MachinePtr)
Say ("  TrackNo:    {0}" -f $TrackNo)
Say ("  Output:     {0}" -f $OutDir)
Say ''

# ---------------------------------------------------------------
# Section 1: Baseline sanity
# ---------------------------------------------------------------
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 1 — Baseline sanity (cookieless home page)' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$homeUrl = "$base/vm_hompage_prc?p_user=$UserId&p_site=$SiteId"
Say ("  GET {0}" -f $homeUrl)
$home = Fetch -url $homeUrl -label 'home'
if ($home.Ok) {
    Save-Blob '01-home.html' $home.Content
    $hasPmBtn = $home.Content -match 'EM_PREV_MAINT_001'
    Say ("        has PM button: {0}" -f $hasPmBtn) $(if ($hasPmBtn) { 'Green' } else { 'Yellow' })
} else {
    Say '        ABORT: cookieless access not working. Check site/user IDs.' 'Red'
}

# ---------------------------------------------------------------
# Section 2: eCBM chain verification
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 2 — eCBM chain verification' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

Say '  2a. getAllChecklists'
$url2a = "$base/ecbm.ecbm_report_j.getAllChecklists?pSiteID=$SiteId&pAcronym=$Acronym"
Say ("      GET {0}" -f $url2a)
$r2a = Fetch -url $url2a -label 'getAllChecklists'
if ($r2a.Ok) { Save-Blob '02a-getAllChecklists.xml' $r2a.Content }

Say '  2b. getChecklists2'
$url2b = "$base/ecbm.ecbm_report_j.getChecklists2?p_machineid=$MachinePtr&pSiteID=$SiteId"
Say ("      GET {0}" -f $url2b)
$r2b = Fetch -url $url2b -label 'getChecklists2'
if ($r2b.Ok) { Save-Blob '02b-getChecklists2.xml' $r2b.Content }

Say '  2c. getChecklistDetailsForIframe (baseline — should be the WRONG format)'
$url2c = "$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=$TrackNo&p_site=$SiteId&p_fac=$SiteId&p_mtype=$Acronym&p_mno=16&p_classcode=CK"
Say ("      GET {0}" -f $url2c)
$r2c = Fetch -url $url2c -label 'getChecklistDetailsForIframe'
if ($r2c.Ok) {
    Save-Blob '02c-getChecklistDetailsForIframe.html' $r2c.Content
    if ($r2c.Markers) {
        Say ("        MARKER HIT: {0}" -f ($r2c.Markers -join '|')) 'Green'
        [void]$script:findings.Add([PSCustomObject]@{ Kind = 'URL'; Url = $url2c; Markers = $r2c.Markers; Note = 'getChecklistDetailsForIframe' })
    } else {
        Say '        (wrong format — MMO Item No / Run Hours style, as expected)' 'Yellow'
    }
}

# ---------------------------------------------------------------
# Section 3: Fetch related JS files
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 3 — Fetch related JS files' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$jsCandidates = @(
    'ECBM_MAIN_001.js'
    'ECBM_CHECKLISTS.js'
    'ECBM_TASK_DETAILS.js'
    'ECBM_WORK_001.js'
    'ECBM_PM_001.js'
    'ECBM_PM_CHECKLIST_001.js'
    'EM_WORK_MAINT_001.js'
    'EM_REACT_MAINT_001.js'
    'EM_MAINT_REPORTS_001.js'
    'em_dialog.js'
    'vmdialogs.js'
)

$loadedJs = @{}
foreach ($js in $jsCandidates) {
    $url = "$Origin/emars-script/$js"
    Say ("  GET {0}" -f $url)
    $r = Fetch -url $url -label $js -Quiet
    if ($r.Ok) {
        Say ("        OK  {0,7:N0} chars" -f $r.Length) 'Green'
        Save-Blob ("03-$js") $r.Content
        $loadedJs[$js] = $r.Content
    } else {
        Say '        miss' 'DarkGray'
    }
}

# ---------------------------------------------------------------
# Section 4: Scan loaded JS for URL patterns
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 4 — Scan JS for URL patterns' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

foreach ($name in $loadedJs.Keys) {
    Scan-Js-For-Urls -js $loadedJs[$name] -label $name
}

# ---------------------------------------------------------------
# Section 5: Probe candidate URLs
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 5 — Probe candidate URLs for the RIGHT format' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'
Say '  (marker = {0})' 'Cyan'
foreach ($m in $magicMarkers) { Say ("    {0}" -f $m) 'DarkGray' }
Say ''

$candidates = @()

# --- eCBM task detail variants ---
$candidates += @(
    @{ N='ecbm getTaskDetails'; U="$base/ecbm.ecbm_report_j.getTaskDetails?p_trackno=$TrackNo&p_site=$SiteId&p_taskid=0" }
    @{ N='ecbm getTaskDetail';  U="$base/ecbm.ecbm_report_j.getTaskDetail?p_trackno=$TrackNo&p_site=$SiteId&p_taskid=0" }
    @{ N='ecbm getChecklistTasks'; U="$base/ecbm.ecbm_report_j.getChecklistTasks?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='ecbm getPMChecklist'; U="$base/ecbm.ecbm_report_j.getPMChecklist?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='ecbm getChecklistDetails'; U="$base/ecbm.ecbm_report_j.getChecklistDetails?p_trackno=$TrackNo&p_site=$SiteId&p_fac=$SiteId&p_mtype=$Acronym&p_mno=16&p_classcode=CK" }
    @{ N='ecbm getChecklistForPrint'; U="$base/ecbm.ecbm_report_j.getChecklistForPrint?p_trackno=$TrackNo&p_site=$SiteId&p_fac=$SiteId&p_mtype=$Acronym&p_mno=16&p_classcode=CK" }
    @{ N='ecbm printChecklist'; U="$base/ecbm.ecbm_report_j.printChecklist?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='ecbm getChecklistPM'; U="$base/ecbm.ecbm_report_j.getChecklistPM?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='ecbm getChecklistPMDetails'; U="$base/ecbm.ecbm_report_j.getChecklistPMDetails?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='ecbm getPmTaskDetails'; U="$base/ecbm.ecbm_report_j.getPmTaskDetails?p_trackno=$TrackNo&p_site=$SiteId&p_taskid=0" }
)

# --- PM module print variants ---
$candidates += @(
    @{ N='pm print_pm_checklist'; U="$base/em_prevent_maint_pkg.print_pm_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='pm em_print_checklist'; U="$base/em_prevent_maint_pkg.em_print_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='pm get_pm_checklist';   U="$base/em_prevent_maint_pkg.get_pm_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='pm getPMChecklist';     U="$base/em_prevent_maint_pkg.getPMChecklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='pm em_get_checklist';   U="$base/em_prevent_maint_pkg.em_get_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='pm em_print_pm_checklist'; U="$base/em_prevent_maint_pkg.em_print_pm_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
)

# --- Work module print variants ---
$candidates += @(
    @{ N='work print_checklist'; U="$base/em_work_pkg.print_checklist?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='work em_print_checklist'; U="$base/em_work_pkg.em_print_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='work get_checklist_details'; U="$base/em_work_pkg.get_checklist_details?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='work em_get_checklist'; U="$base/em_work_pkg.em_get_checklist?p_trackno=$TrackNo&p_site_id=$SiteId" }
    @{ N='work em_checklist_details'; U="$base/em_work_pkg.em_checklist_details?p_trackno=$TrackNo&p_site_id=$SiteId" }
)

# --- Report endpoint variants ---
$xmlKey = "<KEYLIST><TRACKNO>$TrackNo</TRACKNO></KEYLIST>"
$xmlEnc = [System.Uri]::EscapeDataString($xmlKey)
$candidates += @(
    @{ N='report em_show_report'; U="$base/em_show_report?pReportID=PM_CHECKLIST&pUser=$UserId&pSite=$SiteId&pParamXML=$xmlEnc" }
    @{ N='report em_show_report v2'; U="$base/em_show_report?pReportID=CHECKLIST&pUser=$UserId&pSite=$SiteId&pParamXML=$xmlEnc" }
    @{ N='report em_show_report v3'; U="$base/em_show_report?pReportID=MAINT_CHECKLIST&pUser=$UserId&pSite=$SiteId&pParamXML=$xmlEnc" }
    @{ N='report em_show_report v4'; U="$base/em_show_report?pReportID=WO_CHECKLIST&pUser=$UserId&pSite=$SiteId&pParamXML=$xmlEnc" }
    @{ N='report em_show_report v5'; U="$base/em_show_report?pReportID=PM_CHECKLIST_PRINT&pUser=$UserId&pSite=$SiteId&pParamXML=$xmlEnc" }
)

# --- Try module shells for task details (may or may not be gated) ---
$candidates += @(
    @{ N='shell ECBM_TASK_DETAILS'; U="$base/!vm_main_pkg.main?pSiteID=$SiteId&pUser=$UserId&pModule=ECBM_TASK_DETAILS" }
    @{ N='shell ECBM_TASK_DETAILS+taskid'; U="$base/!vm_main_pkg.main?pSiteID=$SiteId&pUser=$UserId&pModule=ECBM_TASK_DETAILS&p_taskid=0" }
)

# --- Try getChecklistDetailsForIframe with different param combos ---
$candidates += @(
    @{ N='iframe no-classcode'; U="$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=$TrackNo&p_site=$SiteId&p_fac=$SiteId&p_mtype=$Acronym&p_mno=16" }
    @{ N='iframe minimal'; U="$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=$TrackNo&p_site=$SiteId" }
    @{ N='iframe with user'; U="$base/ecbm.ecbm_report_j.getChecklistDetailsForIframe?p_trackno=$TrackNo&p_site=$SiteId&p_fac=$SiteId&p_mtype=$Acronym&p_mno=16&p_classcode=CK&p_user=$UserId" }
)

$index = 0
foreach ($c in $candidates) {
    $index++
    Say ("  [{0,2}] {1}" -f $index, $c.N) 'Cyan'
    Say ("        {0}" -f $c.U) 'DarkGray'
    $r = Fetch -url $c.U -label $c.N

    if ($r.Ok) {
        $safeName = "05-{0:00}-{1}.txt" -f $index, ($c.N -replace '[^\w]', '_')
        Save-Blob $safeName $r.Content

        if ($r.Markers) {
            Say ("        *** MATCH: {0} ***" -f ($r.Markers -join '|')) 'Green'
            [void]$script:findings.Add([PSCustomObject]@{
                Kind    = 'URL'
                Url     = $c.U
                Markers = $r.Markers
                Note    = $c.N
            })
        }
    }
}

# ---------------------------------------------------------------
# Section 6: Also try ECBM_TASK_DETAILS.js URLs directly
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'SECTION 6 — Probe URLs found inside ECBM_TASK_DETAILS.js' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

if ($loadedJs.ContainsKey('ECBM_TASK_DETAILS.js')) {
    $tdJs = $loadedJs['ECBM_TASK_DETAILS.js']

    # Look for getServerPath() + 'something'
    $urlBuilders = [regex]::Matches($tdJs, "(?i)getServerPath\s*\(\s*\)\s*\+\s*['""]((?:ecbm|em_)[a-zA-Z0-9_.?=&+]+)")
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($m in $urlBuilders) { $found.Add($m.Groups[1].Value) }

    if ($found.Count -eq 0) {
        Say '  No getServerPath()+ URLs found in ECBM_TASK_DETAILS.js' 'Yellow'

        # Alternative: look for objHTTP.open patterns
        $opens = [regex]::Matches($tdJs, "(?i)objHTTP\.open\s*\(\s*['""][A-Z]+['""]\s*,\s*getServerPath\(\)\s*\+\s*['""]([^'""]+)")
        foreach ($m in $opens) { $found.Add($m.Groups[1].Value) }
    }

    if ($found.Count -gt 0) {
        Say ("  Found {0} candidate URL(s):" -f $found.Count) 'Green'
        foreach ($u in ($found | Sort-Object -Unique)) {
            Say ("    {0}" -f $u) 'DarkGray'
            $fullUrl = "$base/$u"
            Say ("    GET {0}" -f $fullUrl)
            $r = Fetch -url $fullUrl -label "TASK_DETAILS: $u"
            if ($r.Ok) {
                if ($r.Markers) {
                    Say ("        *** MATCH: {0} ***" -f ($r.Markers -join '|')) 'Green'
                    [void]$script:findings.Add([PSCustomObject]@{
                        Kind    = 'URL'
                        Url     = $fullUrl
                        Markers = $r.Markers
                        Note    = "from ECBM_TASK_DETAILS.js"
                    })
                }
            }
        }
    }
} else {
    Say '  ECBM_TASK_DETAILS.js not loaded — skipping' 'Yellow'
}

# ---------------------------------------------------------------
# Section 7: Summary
# ---------------------------------------------------------------
Say ''
Say '============================================================' 'Cyan'
Say '  PROBE COMPLETE — SUMMARY' 'Cyan'
Say '============================================================' 'Cyan'
Say ''

if ($script:findings.Count -gt 0) {
    Say '  ✓ FOUND URLs returning the EXPECTED checklist format:' 'Green'
    Say ''
    foreach ($f in $script:findings) {
        Say ("    [{0}]" -f $f.Note) 'Green'
        Say ("         {0}" -f $f.Url) 'Gray'
        Say ("         Markers: {0}" -f ($f.Markers -join ', ')) 'Gray'
        Say ''
    }
    Say '  → Paste back: this report + the matching file from the output folder.' 'Cyan'
    Say '  → Next step: wire the winning URL into a Fetch button in UI-Script.ps1' 'Cyan'
} else {
    Say '  ✗ No URL returned the expected format markers.' 'Yellow'
    Say ''
    Say '  This means the format your parser expects is NOT served by any' 'Yellow'
    Say '  of the URLs the app exposes directly. The most likely explanations:' 'Yellow'
    Say ''
    Say '    1. The format you paste comes from a UI flow we haven'"'"'t traced,' 'Yellow'
    Say '       e.g. clicking a specific task inside eCBM opens a deeper detail view.' 'Yellow'
    Say '    2. The format is generated client-side by IE from data we can fetch.' 'Yellow'
    Say '    3. It comes from a report endpoint that requires login state.' 'Yellow'
    Say ''
    Say '  Next steps:' 'Cyan'
    Say '    a) In IE, reproduce the manual popup one more time.' 'Cyan'
    Say '       When it opens: right-click → View Source → Ctrl+A, Ctrl+C' 'Cyan'
    Say '       → paste the first ~500 chars here.' 'Cyan'
    Say '    b) OR in the popup: F12 → Network → clear → click the button again' 'Cyan'
    Say '       → look for the request whose response contains "Task Statement"' 'Cyan'
    Say '       → paste that URL.' 'Cyan'
    Say ''
    Say '  Meanwhile: the clipboard watcher remains the working solution.' 'Cyan'
    Say '  It works today. Ship it, iterate on full automation in parallel.' 'Cyan'
}

Say ''
Say ("  Output folder: {0}" -f $OutDir) 'Cyan'
Say ("  Full report:   {0}" -f $reportPath) 'Cyan'
Say '============================================================' 'Cyan'

# ---------------------------------------------------------------
# Write report
# ---------------------------------------------------------------
$script:lines | Set-Content -Path $reportPath -Encoding UTF8

Write-Host ''
Write-Host "Saved to: $OutDir" -ForegroundColor Cyan
Write-Host ''
Get-ChildItem -Path $OutDir -File | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-50}  {1,10:N0} bytes" -f $_.Name, $_.Length) -ForegroundColor DarkGray
}