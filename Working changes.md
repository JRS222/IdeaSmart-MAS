#Requires -Version 5.1
$origin = 'http://emarssu3.eng.usps.gov'
$outDir = "$env:USERPROFILE\Desktop\ecbm-probe-v2"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

function Fetch-Js {
    param([string]$url, [string]$saveAs)
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
        [System.IO.File]::WriteAllText((Join-Path $outDir $saveAs), $r.Content, (New-Object System.Text.UTF8Encoding($false)))
        "  FOUND  $saveAs  ($($r.Content.Length) chars)"
        return $r.Content
    } catch {
        "  miss   $saveAs  ($($_.Exception.Message))"
        return $null
    }
}

function Dump-OpenSwilWin {
    param([string]$js, [string]$label)
    if ([string]::IsNullOrWhiteSpace($js)) { return }

    # Bracket-matched extraction of the openSwilWin function body
    $m = [regex]::Match($js, '(?is)function\s+openSwilWin\s*\([^)]*\)\s*\{')
    if (-not $m.Success) {
        "  [$label] openSwilWin() not found in this file"
        return
    }

    $depth = 0
    $i = $m.Index + $m.Length - 1
    $end = -1
    while ($i -lt $js.Length) {
        $ch = $js[$i]
        if ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) { $end = $i; break }
        }
        $i++
    }

    if ($end -lt 0) {
        "  [$label] openSwilWin() found but braces unbalanced"
        return
    }

    $body = $js.Substring($m.Index, $end - $m.Index + 1)
    "  [$label] openSwilWin() source:"
    foreach ($line in ($body -split "`r?`n")) {
        "      $line"
    }
}

function Dump-ChecklistLines {
    param([string]$js, [string]$label)
    if ([string]::IsNullOrWhiteSpace($js)) { return }

    $lines = $js -split "`r?`n"
    $i = 0
    $n = 0
    foreach ($ln in $lines) {
        $i++
        if ($ln -match '(?i)ECBM_CHECKLISTS|ECBM_CHECKLIST|p_checklist|checklist') {
            "  [$label L$i] $($ln.Trim())"
            $n++
            if ($n -ge 40) {
                "  [$label] ... more lines suppressed; see saved file"
                break
            }
        }
    }
    if ($n -eq 0) {
        "  [$label] no checklist-related lines"
    }
}

"============================================================"
"  eCBM CHECKLIST — PROBE v2"
"  Output: $outDir"
"============================================================"

# ---- 1. vmdialogs.js (where openSwilWin should live) ----
""
"--- 1. vmdialogs.js ---"
$vd = Fetch-Js -url "$origin/emars-script/vmdialogs.js" -saveAs 'vmdialogs.js'
Dump-OpenSwilWin -js $vd -label 'vmdialogs.js'

# ---- 2. em_dialog.js (backup location) ----
""
"--- 2. em_dialog.js ---"
$ed = Fetch-Js -url "$origin/emars-script/em_dialog.js" -saveAs 'em_dialog.js'
Dump-OpenSwilWin -js $ed -label 'em_dialog.js'

# ---- 3. ECBM_CHECKLISTS.js (the sub-module itself) ----
""
"--- 3. ECBM_CHECKLISTS.js ---"
$cl = Fetch-Js -url "$origin/emars-script/ECBM_CHECKLISTS.js" -saveAs 'ECBM_CHECKLISTS.js'
if ($cl) {
    Dump-ChecklistLines -js $cl -label 'ECBM_CHECKLISTS.js'
}

# ---- 4. Scan the ECBM main JS for every checklist dispatch ----
""
"--- 4. ECBM_MAIN_001.js checklist dispatch lines ---"
try {
    $main = (Invoke-WebRequest -Uri "$origin/emars-script/ECBM_MAIN_001.js" -UseBasicParsing -TimeoutSec 20).Content
    Dump-ChecklistLines -js $main -label 'ECBM_MAIN_001.js'
} catch {
    "  miss ECBM_MAIN_001.js"
}

""
"============================================================"
"  Files saved to: $outDir"
"  Paste back:"
"    - The openSwilWin() body from vmdialogs.js or em_dialog.js"
"    - Any ECBM_CHECKLISTS.js lines that mention checklist or URLs"
"============================================================"