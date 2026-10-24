#Requires -Version 5.1
<#
.SYNOPSIS
    eMARS PM checklist discovery test — self-contained, read-only.
.DESCRIPTION
    Tests four things against eMARS:
      1. Cookieless home-page access (baseline)
      2. Login-audit registration (browser does this on load)
      3. PM module fetch after audit
      4. Common PM JS filenames
    Then scans everything for PM/checklist URL patterns and writes a report.
.PARAMETER UserId
    eMARS user ID (p_user value from the home-page URL)
.PARAMETER SiteId
    eMARS site ID (p_site value from the home-page URL)
.PARAMETER BaseUrl
    eMARS web root (default is the production URL you've been using)
.NOTES
    No credentials required. No writes to eMARS. Safe to re-run.
#>
[CmdletBinding()]
param(
    [string]$UserId  = '1601089',
    [string]$SiteId  = '345',
    [string]$BaseUrl = 'http://emarssu3.eng.usps.gov/pemars'
)

# ---------------------------------------------------------------
# Setup
# ---------------------------------------------------------------
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

$outDir = Join-Path $env:USERPROFILE ("Desktop\emars-test-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$reportPath = Join-Path $outDir 'TEST-REPORT.txt'

$script:lines = New-Object System.Collections.ArrayList

function Say {
    param([string]$msg = '', [string]$color = 'Gray')
    Write-Host $msg -ForegroundColor $color
    [void]$script:lines.Add($msg)
}

function Save-Blob {
    param([string]$name, [string]$content)
    if ($null -eq $content) { return }
    $path = Join-Path $outDir $name
    [System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))
    Say ("        saved {0} ({1:N0} chars)" -f $name, $content.Length) 'DarkGray'
}

function Try-Get {
    param([string]$url, [string]$label = '')
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        return $r
    } catch {
        Say ("        GET FAILED: {0} — {1}" -f $url, $_.Exception.Message) 'Red'
        return $null
    }
}

# ---------------------------------------------------------------
# Banner
# ---------------------------------------------------------------
Say '============================================================' 'Cyan'
Say '  eMARS PM CHECKLIST — DISCOVERY TEST' 'Cyan'
Say '============================================================' 'Cyan'
Say ""
Say ("  User ID:  {0}" -f $UserId)
Say ("  Site ID:  {0}" -f $SiteId)
Say ("  Base URL: {0}" -f $BaseUrl)
Say ("  Output:   {0}" -f $outDir)
Say ""

$results = [ordered]@{
    HomePage        = 'NOT RUN'
    LoginAudit      = 'NOT RUN'
    PmModule        = 'NOT RUN'
    PmModuleIsHome  = 'NOT CHECKED'
    PmJsFiles       = 'NOT RUN'
    PmPatterns      = 'NONE'
}

# ---------------------------------------------------------------
# Step 1: home page baseline
# ---------------------------------------------------------------
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 1 — Home page (cookieless baseline)' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$homeUrl = "$BaseUrl/vm_hompage_prc?p_user=$UserId&p_site=$SiteId"
Say ("  GET {0}" -f $homeUrl)
$homeResp = Try-Get -url $homeUrl

if ($homeResp) {
    Save-Blob '01-home.html' $homeResp.Content
    $hasPmBtn = ($homeResp.Content -match 'EM_PREV_MAINT')
    Say ("        status: {0}   length: {1:N0}   has PM button: {2}" -f `
         $homeResp.StatusCode, $homeResp.Content.Length, $hasPmBtn) 'Green'
    if ($hasPmBtn) {
        $results.HomePage = "OK ($($homeResp.Content.Length) chars)"
    } else {
        $results.HomePage = "WARN — no EM_PREV_MAINT (IDs may be stale)"
    }
} else {
    $results.HomePage = 'FAILED'
}

# ---------------------------------------------------------------
# Step 2: login-audit registration
# ---------------------------------------------------------------
Say ""
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 2 — Login audit (what the browser does on load)' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$auditUrl = "$BaseUrl/emapp.em_login_audit_pkg.CheckLogin_InOut?p_user_id=$UserId&p_value=1"
Say ("  POST {0}" -f $auditUrl)

$auditResp = $null
try {
    $auditResp = Invoke-WebRequest -Uri $auditUrl -Method POST -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
    Say ("        status: {0}   length: {1:N0}" -f $auditResp.StatusCode, $auditResp.Content.Length) 'Green'
    Save-Blob '02-audit-response.txt' $auditResp.Content
    $results.LoginAudit = "OK ($($auditResp.StatusCode))"
} catch {
    Say ("        FAILED: {0}" -f $_.Exception.Message) 'Red'
    $results.LoginAudit = "FAILED — $($_.Exception.Message)"
}

# ---------------------------------------------------------------
# Step 3: PM module fetch
# ---------------------------------------------------------------
Say ""
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 3 — PM module fetch' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$pmUrl = "$BaseUrl/!vm_main_pkg.main?pSiteID=$SiteId&pUser=$UserId&pModule=EM_PREV_MAINT_001"
Say ("  GET {0}" -f $pmUrl)
$pmResp = Try-Get -url $pmUrl

if ($pmResp) {
    Save-Blob '03-pm-module.html' $pmResp.Content
    $isHome     = ($pmResp.Content -match 'vm_hompage_prc\.js')
    $hasPmWord  = ($pmResp.Content -match '(?i)PM|Checklist|PREV_MAINT|pm_pkg')
    $isLogin    = ($pmResp.Content -match '(?i)EMARS Login Screen|em_user_login_pkg')
    $title      = [regex]::Match($pmResp.Content, '(?is)<title>(.*?)</title>').Groups[1].Value

    Say ("        status: {0}   length: {1:N0}" -f $pmResp.StatusCode, $pmResp.Content.Length) 'Green'
    Say ("        title:    {0}" -f $title)
    Say ("        is home?  {0}   has PM words? {1}   is login? {2}" -f `
         (-not $isHome), $hasPmWord, $isLogin) `
         $(if (-not $isHome) { 'Green' } else { 'Yellow' })

    if ($isLogin) {
        $results.PmModule = 'RETURNED LOGIN PAGE'
        $results.PmModuleIsHome = 'n/a'
    } elseif ($isHome) {
        $results.PmModule = "OK ($($pmResp.Content.Length) chars) — but is HOME SHELL"
        $results.PmModuleIsHome = 'YES (server refused — needs real browser)'
    } elseif ($hasPmWord) {
        $results.PmModule = "OK ($($pmResp.Content.Length) chars) — looks like PM module"
        $results.PmModuleIsHome = 'NO (good!)'
    } else {
        $results.PmModule = "OK ($($pmResp.Content.Length) chars) — unknown shell"
        $results.PmModuleIsHome = 'NO'
    }
} else {
    $results.PmModule = 'FAILED'
}

# ---------------------------------------------------------------
# Step 4: try common PM JS filenames
# ---------------------------------------------------------------
Say ""
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 4 — Common PM JavaScript filenames' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$origin = ([Uri]$BaseUrl).GetLeftPart([UriPartial]::Authority)

$jsCandidates = @(
    'EM_PREV_MAINT_001.js',
    'EM_PREV_MAINT.js',
    'em_pm_pkg.js',
    'em_pm.js',
    'vm_prev_maint_prc.js',
    'em_preventive_maint.js',
    'em_pm_checklist.js',
    'pm_checklist.js'
)

$foundJs = @()
foreach ($js in $jsCandidates) {
    $url = "$origin/emars-script/$js"
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        if ($r.StatusCode -eq 200 -and $r.Content.Length -gt 100 -and
            $r.Content -notmatch '(?i)404|not found') {
            Say ("        FOUND  {0}  ({1:N0} chars)" -f $js, $r.Content.Length) 'Green'
            Save-Blob ("04-" + $js) $r.Content
            $foundJs += $js
        } else {
            Say ("        miss   {0}" -f $js) 'DarkGray'
        }
    } catch {
        Say ("        miss   {0}" -f $js) 'DarkGray'
    }
}

if ($foundJs.Count -gt 0) {
    $results.PmJsFiles = "FOUND $($foundJs.Count): $($foundJs -join ', ')"
} else {
    $results.PmJsFiles = 'NONE — check IE View Source for the script path'
}

# ---------------------------------------------------------------
# Step 5: pattern scan across everything captured
# ---------------------------------------------------------------
Say ""
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 5 — Pattern scan across all captured files' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$patterns = [ordered]@{
    'em_pm_pkg call'        = 'em_pm_pkg\.[a-zA-Z0-9_]+'
    'pkg.PM procs'          = 'em_[a-z_]*pm[a-z_]*_pkg\.[a-zA-Z0-9_]+'
    'PREV_MAINT module'     = 'EM_PREV_MAINT_[A-Z0-9_]+'
    'checklist param'       = 'p_checklist[a-zA-Z_]*'
    'work-order param'      = 'p_wo\b'
    'PM sql/proc'           = '[a-zA-Z0-9_]*pm_[a-zA-Z0-9_]+'
    'Checklist No. text'    = 'Checklist No\.'
}

$allFiles = @(Get-ChildItem -Path $outDir -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -notmatch '^TEST-REPORT' })

$findings = [ordered]@{}
foreach ($patName in $patterns.Keys) {
    $rx  = $patterns[$patName]
    $hit = @()
    foreach ($f in $allFiles) {
        $c = Get-Content -Path $f.FullName -Raw -ErrorAction SilentlyContinue
        if (-not $c) { continue }
        foreach ($m in [regex]::Matches($c, $rx, 'IgnoreCase')) {
            $hit += "$($m.Value)  [in $($f.Name)]"
        }
    }
    if ($hit.Count -gt 0) {
        $unique = @($hit | Sort-Object -Unique)
        $findings[$patName] = $unique
        Say ("  {0,-22} : {1} distinct hit(s)" -f $patName, $unique.Count) 'Green'
        foreach ($h in ($unique | Select-Object -First 8)) {
            Say ("      {0}" -f $h) 'DarkGray'
        }
        if ($unique.Count -gt 8) {
            Say ("      ... and {0} more" -f ($unique.Count - 8)) 'DarkGray'
        }
    } else {
        Say ("  {0,-22} : (none)" -f $patName) 'DarkGray'
    }
}

if ($findings.Count -gt 0) {
    $results.PmPatterns = "FOUND $($findings.Count) pattern category/ies"
} else {
    $results.PmPatterns = 'NONE'
}

# ---------------------------------------------------------------
# Also scan any HTML files loaded from JS to find script src tags
# ---------------------------------------------------------------
Say ""
Say '------------------------------------------------------------' 'Cyan'
Say 'STEP 6 — <script src=> tags in captured files' 'Cyan'
Say '------------------------------------------------------------' 'Cyan'

$allScriptSrc = @()
foreach ($f in ($allFiles | Where-Object { $_.Name -match '\.html$' })) {
    $c = Get-Content -Path $f.FullName -Raw -ErrorAction SilentlyContinue
    if (-not $c) { continue }
    foreach ($m in [regex]::Matches($c, '(?is)<script\b[^>]*\bsrc\s*=\s*"([^"]+)"')) {
        $s = $m.Groups[1].Value
        if ($s -notin $allScriptSrc) { $allScriptSrc += $s }
    }
}

if ($allScriptSrc.Count -gt 0) {
    Say ("  {0} script tag(s) found:" -f $allScriptSrc.Count) 'Green'
    foreach ($s in $allScriptSrc) {
        Say ("      {0}" -f $s) 'DarkGray'
    }
    Say ""
    Say "  >>> If any of those look PM/checklist-related, save the JS file from IE"
    Say "      (View Source -> click the JS link) and paste its content back. <<<"
} else {
    Say '  (no script tags captured)' 'DarkGray'
}

# ---------------------------------------------------------------
# Summary
# ---------------------------------------------------------------
Say ""
Say '============================================================' 'Cyan'
Say '  RESULTS SUMMARY' 'Cyan'
Say '============================================================' 'Cyan'

foreach ($k in $results.Keys) {
    $v = $results[$k]
    $color = if ($v -match '^OK|^FOUND') { 'Green' }
             elseif ($v -match 'NOT RUN|NOT CHECKED|NOT SET') { 'DarkGray' }
             else { 'Yellow' }
    Say ("  {0,-15} : {1}" -f $k, $v) $color
}

Say ""
Say '============================================================' 'Cyan'
Say '  NEXT STEPS' 'Cyan'
Say '============================================================' 'Cyan'

if ($results.PmModule -match 'looks like PM module') {
    Say '  ✓ PM module loaded!' 'Green'
    Say '  → Paste back: 03-pm-module.html' 'Green'
} elseif ($results.PmJsFiles -match '^FOUND') {
    Say '  ✓ PM JavaScript found!' 'Green'
    Say '  → Paste back the 04-*.js files' 'Green'
} elseif ($results.PmModule -match 'is HOME SHELL') {
    Say '  The server served the home shell instead of the PM module.' 'Yellow'
    Say '  → Open PM module in IE, do File > Save As (Web Page, HTML only),' 'Yellow'
    Say '    save to Desktop, and paste that file back.' 'Yellow'
} else {
    Say '  → Open PM module in IE, press Ctrl+U (View Source), Ctrl+A, Ctrl+C,' 'Yellow'
    Say '    and paste the source back.' 'Yellow'
}

Say ""
Say ("  All files saved to: {0}" -f $outDir) 'Cyan'
Say ("  Full report:        {0}" -f $reportPath) 'Cyan'
Say '============================================================' 'Cyan'

# Write the report
$script:lines | Set-Content -Path $reportPath -Encoding UTF8

Write-Host ""
Write-Host "Report written. Attach or paste:" -ForegroundColor Cyan
Write-Host "  $reportPath" -ForegroundColor Yellow
Write-Host ""