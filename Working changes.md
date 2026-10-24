#Requires -Version 5.1
<#
.SYNOPSIS
    eCBM PM checklist discovery probe — self-contained, read-only.
.DESCRIPTION
    Fetches the ECBM_CHECKLISTS sub-module JS (or falls back to vmdialogs.js)
    and scans for the PM checklist URL pattern. Reports what it finds and
    saves any downloaded files to a timestamped folder on the Desktop.

    Uses only cookieless requests — the eMARS session is entirely carried
    in the pUser/pSiteID query-string parameters, so no login is needed.
.NOTES
    Re-runnable. No writes to eMARS. Safe.
#>
[CmdletBinding()]
param(
    [string]$Origin  = 'http://emarssu3.eng.usps.gov',
    [string]$OutDir  = (Join-Path $env:USERPROFILE ("Desktop\ecbm-probe-" + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------------------------
# Setup
# ---------------------------------------------------------------
if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$reportPath = Join-Path $OutDir 'PROBE-REPORT.txt'

$script:lines = New-Object System.Collections.ArrayList

function Say {
    param([string]$msg = '', [string]$color = 'Gray')
    Write-Host $msg -ForegroundColor $color
    [void]$script:lines.Add($msg)
}

function Try-Fetch-Js {
    param([string]$url, [string]$saveAs = $null)

    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
        if ($saveAs) {
            $path = Join-Path $OutDir $saveAs
            [System.IO.File]::WriteAllText($path, $r.Content, (New-Object System.Text.UTF8Encoding($false)))
            Say ("        saved: {0} ({1:N0} chars)" -f $saveAs, $r.Content.Length) 'DarkGray'
        }
        return $r.Content
    } catch {
        Say ("        FAILED: {0}" -f $_.Exception.Message) 'DarkGray'
        return $null
    }
}

function Scan-Content {
    param([string]$content, [string]$label)

    if ([string]::IsNullOrWhiteSpace($content)) {
        Say '        (no content to scan)' 'DarkGray'
        return
    }

    $patterns = @(
        'em_[a-z_]*checklist[a-z_]*\.[a-zA-Z0-9_]+'
        'em_[a-z_]*ecbm[a-z_]*\.[a-zA-Z0-9_]+'
        'ecbm_[a-z_]*\.[a-zA-Z0-9_]+'
        'p_checklist[a-zA-Z_]*'
        'p_equip[a-zA-Z_]*'
        'p_wo\b'
        'openSwilWin\s*\([^)]+\)'
        'getInnerHTML\s*\([^)]+\)'
        'window\.open\s*\([^)]+\)'
    )

    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($pat in $patterns) {
        $ms = [regex]::Matches($content, $pat, 'IgnoreCase')
        foreach ($m in $ms) {
            $hits.Add($m.Value.Trim())
        }
    }

    if ($hits.Count -eq 0) {
        Say ("        [{0}] no URL patterns matched" -f $label) 'DarkGray'
    } else {
        Say ("        [{0}] {1} distinct URL hit(s):" -f $label, ($hits | Sort-Object -Unique).Count) 'Green'
        $hits | Sort-Object -Unique | ForEach-Object {
            Say ("          {0}" -f $_) 'DarkGray'
        }
    }

    # Also dump the lines that mention "checklist" — this is where the truth lives.
    $lines = $content -split "`r?`n"
    $checklistLines = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($ln in $lines) {
        $i++
        if ($ln -match '(?i)checklist') {
            [void]$checklistLines.Add([PSCustomObject]@{ LineNo = $i; Text = $ln.Trim() })
        }
    }
    if ($checklistLines.Count -gt 0) {
        Say ("        [{0}] {1} line(s) mention 'checklist':" -f $label, $checklistLines.Count) 'Cyan'
        $shown = 0
        foreach ($cl in $checklistLines) {
            Say ("          [{0,5}] {1}" -f $cl.LineNo, $cl.Text) 'DarkGray'
            $shown++
            if ($shown -ge 30) {
                Say ("          ... ({0} more, see saved file)" -f ($checklistLines.Count - 30)) 'DarkGray'
                break
            }
        }
    } else {
        Say ("        [{0}] no line mentions 'checklist'" -f $label) 'DarkGray'
    }
}

# ---------------------------------------------------------------
# Banner
# ---------------------------------------------------------------
Say '============================================================' 'Cyan'
Say '  eCBM PM CHECKLIST — PROBE (cookieless)' 'Cyan'
Say '============================================================' 'Cyan'
Say ''
Say ("  Origin: {0}" -f $Origin)
Say ("  Output: {0}" -f $OutDir)
Say ''

# ---------------------------------------------------------------
# Step 1: fetch ECBM_CHECKLISTS.js
# ---------------------------------------------------------------
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 1 — Fetch /emars-script/ECBM_CHECKLISTS.js' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$jsUrl = "$Origin/emars-script/ECBM_CHECKLISTS.js"
Say ("  GET {0}" -f $jsUrl)
$ecbmJs = Try-Fetch-Js -url $jsUrl -saveAs 'ECBM_CHECKLISTS.js'

if ($null -ne $ecbmJs) {
    Say ("        status: 200   length: {0:N0}" -f $ecbmJs.Length) 'Green'
    Scan-Content -content $ecbmJs -label 'ECBM_CHECKLISTS.js'
} else {
    Say '        ECBM_CHECKLISTS.js not found — will try alternatives.' 'Yellow'
}

# ---------------------------------------------------------------
# Step 2: try alternative checklist-related JS filenames
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 2 — Try alternative sub-module JS filenames' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$altNames = @(
    'ECBM_CHECKLIST_001.js'
    'ECBM_PM_001.js'
    'ECBM_PM_CHECKLIST_001.js'
    'ECBM_WO_001.js'
    'ECBM_EQUIP_001.js'
    'ECBM_WORK_001.js'
    'ECBM_REPORT_001.js'
    'ECBM_MAIN.js'
)

$foundAlt = New-Object System.Collections.ArrayList
foreach ($name in $altNames) {
    $url = "$Origin/emars-script/$name"
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        if ($r.StatusCode -eq 200 -and $r.Content.Length -gt 200) {
            Say ("        FOUND  {0}  ({1:N0} chars)" -f $name, $r.Content.Length) 'Green'
            $savePath = Join-Path $OutDir $name
            [System.IO.File]::WriteAllText($savePath, $r.Content, (New-Object System.Text.UTF8Encoding($false)))
            [void]$foundAlt.Add([PSCustomObject]@{ Name = $name; Content = $r.Content })
        } else {
            Say ("        skip   {0}  (status={1}, len={2})" -f $name, $r.StatusCode, $r.Content.Length) 'DarkGray'
        }
    } catch {
        Say ("        miss   {0}" -f $name) 'DarkGray'
    }
}

if ($foundAlt.Count -eq 0) {
    Say '        No alternative checklist JS found.' 'Yellow'
} else {
    Say ''
    foreach ($f in $foundAlt) {
        Scan-Content -content $f.Content -label $f.Name
    }
}

# ---------------------------------------------------------------
# Step 3: fetch vmdialogs.js to inspect openSwilWin
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 3 — Inspect vmdialogs.js (openSwilWin signature)' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$vdUrl = "$Origin/emars-script/vmdialogs.js"
Say ("  GET {0}" -f $vdUrl)
$vdJs = Try-Fetch-Js -url $vdUrl -saveAs 'vmdialogs.js'

if ($null -ne $vdJs) {
    Say ("        status: 200   length: {0:N0}" -f $vdJs.Length) 'Green'

    # Print just the openSwilWin function (bracket-matched)
    $rx = [regex]'(?is)function\s+openSwilWin\s*\([^)]*\)\s*\{'
    $m = $rx.Match($vdJs)
    if ($m.Success) {
        $start = $m.Index
        $depth = 0
        $i = $m.Index + $m.Length - 1
        $end = -1
        while ($i -lt $vdJs.Length) {
            $ch = $vdJs[$i]
            if ($ch -eq '{') { $depth++ }
            elseif ($ch -eq '}') {
                $depth--
                if ($depth -eq 0) { $end = $i; break }
            }
            $i++
        }
        if ($end -gt $start) {
            $fn = $vdJs.Substring($start, $end - $start + 1)
            Say ''
            Say '        openSwilWin() source:' 'Cyan'
            foreach ($line in ($fn -split "`r?`n")) {
                Say ("          {0}" -f $line) 'DarkGray'
            }
        } else {
            Say '        openSwilWin() found but could not locate closing brace.' 'Yellow'
        }
    } else {
        Say '        openSwilWin() not found in vmdialogs.js' 'Yellow'
    }

    # Also grab the generic window-open / url-builder functions
    Say ''
    Say '        Other URL-related functions found:' 'Cyan'
    $fnRx = [regex]'(?im)^\s*function\s+(\w+)\s*\(([^)]*)\)'
    $fnHits = New-Object System.Collections.ArrayList
    foreach ($fm in $fnRx.Matches($vdJs)) {
        $fname = $fm.Groups[1].Value
        if ($fname -match '(?i)open|win|swil|url|http|fetch|xml|get|call') {
            [void]$fnHits.Add($fname)
        }
    }
    if ($fnHits.Count -gt 0) {
        $fnHits | Sort-Object -Unique | ForEach-Object {
            Say ("          {0}()" -f $_) 'DarkGray'
        }
    } else {
        Say '        (none)' 'DarkGray'
    }
}

# ---------------------------------------------------------------
# Step 4: fetch em_dialog.js for getServerPath / similar
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 4 — Inspect em_dialog.js (URL builders)' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$edUrl = "$Origin/emars-script/em_dialog.js"
Say ("  GET {0}" -f $edUrl)
$edJs = Try-Fetch-Js -url $edUrl -saveAs 'em_dialog.js'

if ($null -ne $edJs) {
    Say ("        status: 200   length: {0:N0}" -f $edJs.Length) 'Green'

    $fnRx = [regex]'(?im)^\s*function\s+(\w+)\s*\(([^)]*)\)'
    $interesting = New-Object System.Collections.ArrayList
    foreach ($fm in $fnRx.Matches($edJs)) {
        $fname = $fm.Groups[1].Value
        if ($fname -match '(?i)serverpath|open|swil|window|url|module|modal|inner') {
            [void]$interesting.Add($fname)
        }
    }
    if ($interesting.Count -gt 0) {
        Say '        URL/module functions in em_dialog.js:' 'Cyan'
        $interesting | Sort-Object -Unique | ForEach-Object {
            Say ("          {0}()" -f $_) 'DarkGray'
        }
    }
}

# ---------------------------------------------------------------
# Step 5: build guess URL for the module shell
# ---------------------------------------------------------------
Say ''
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 5 — Probe module shell URLs' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$user = '1601089'
$site = '345'
$base = "$Origin/pemars"
$shellProbes = @(
    @{ Module = 'ECBM_CHECKLISTS';   Label = 'ECBM_CHECKLISTS' }
    @{ Module = 'ECBM_CHECKLIST_001'; Label = 'ECBM_CHECKLIST_001' }
    @{ Module = 'ECBM_PM_001';       Label = 'ECBM_PM_001' }
)

foreach ($probe in $shellProbes) {
    $url = "$base/!vm_main_pkg.main?pSiteID=$site&pUser=$user&pModule=$($probe.Module)"
    Say ("  GET {0}" -f $url)
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
        $isHome = $r.Content -match 'vm_hompage_prc\.js'
        $len = $r.Content.Length
        if ($isHome) {
            Say ("        len={0} — HOME SHELL (server refused)" -f $len) 'Yellow'
        } else {
            Say ("        len={0} — NOT home shell — saving!" -f $len) 'Green'
            $saveName = "shell-$($probe.Module).html"
            $savePath = Join-Path $OutDir $saveName
            [System.IO.File]::WriteAllText($savePath, $r.Content, (New-Object System.Text.UTF8Encoding($false)))
            Say ("        saved: {0}" -f $saveName) 'DarkGray'
            Scan-Content -content $r.Content -label $saveName
        }
    } catch {
        Say ("        FAILED: {0}" -f $_.Exception.Message) 'DarkGray'
    }
}

# ---------------------------------------------------------------
# Summary
# ---------------------------------------------------------------
Say ''
Say '============================================================' 'Cyan'
Say '  PROBE COMPLETE' 'Cyan'
Say '============================================================' 'Cyan'
Say ''
Say ("  Output folder: {0}" -f $OutDir)
Say ("  Full report:   {0}" -f $reportPath)
Say ''
Say '  What to send back:' 'Cyan'
Say '    1. This entire PROBE-REPORT.txt (paste it in the reply)'
Say '    2. Any file whose name starts with "ECBM_" from the output folder'
Say ''
Say '  If STEP 1 found a real ECBM_CHECKLISTS.js, that file is the answer.' 'Cyan'
Say '  If it didn'"'"'t, STEP 3 will show how openSwilWin builds the module URL,' 'Cyan'
Say '  and we can reconstruct the checklist URL from that pattern.' 'Cyan'
Say ''
Say '============================================================' 'Cyan'

# Write the report
$script:lines | Set-Content -Path $reportPath -Encoding UTF8

Write-Host ''
Write-Host 'Saved:' -ForegroundColor Cyan
Write-Host "  $reportPath" -ForegroundColor Yellow
Write-Host ''
Write-Host 'Files in output folder:' -ForegroundColor Cyan
Get-ChildItem -Path $OutDir -File | ForEach-Object {
    Write-Host ("  {0}  ({1:N0} bytes)" -f $_.Name, $_.Length) -ForegroundColor DarkGray
}