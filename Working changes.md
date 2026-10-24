$origin  = 'http://emarssu3.eng.usps.gov'
$base    = "$origin/pemars"
$user    = '1601089'
$site    = '345'
$outDir  = "$env:USERPROFILE\Desktop\ecbm-probe"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# 1. Fetch the eCBM module JS
$jsUrl = "$origin/emars-script/ECBM_MAIN_001.js"
"GET $jsUrl"
$js = Invoke-WebRequest -UseBasicParsing -Uri $jsUrl -TimeoutSec 20
"  status: $($js.StatusCode)  length: $($js.Content.Length)"
$js.Content | Set-Content "$outDir\01-ECBM_MAIN_001.js" -Encoding UTF8

# 2. Scan it for anything PM/checklist-related
"`nScanning for PM/checklist patterns..."
$patterns = @(
    'em_[a-z_]*pm[a-z_]*_pkg\.[a-zA-Z0-9_]+',
    'em_ecbm[a-z_]*\.[a-zA-Z0-9_]+',
    'p_checklist[a-zA-Z_]*',
    'checklist[a-zA-Z_]*',
    'p_wo\b',
    'printfig|print_check|print_checklist',
    'window\.open\([^)]+\)'
)
$hits = @{}
foreach ($pat in $patterns) {
    foreach ($m in [regex]::Matches($js.Content, $pat, 'IgnoreCase')) {
        $key = $m.Value
        if ($hits.ContainsKey($key)) { $hits[$key]++ } else { $hits[$key] = 1 }
    }
}
"  $($hits.Count) distinct hit(s):"
$hits.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    "    x{0,-3}  {1}" -f $_.Value, $_.Key
}

# 3. Try common eCBM submodule JS filenames
"`nTrying eCBM submodule JS files..."
foreach ($n in @('ECBM_MAIN_001','ECBM_CHECKLIST_001','ECBM_PM_001',
                 'ECBM_PM_CHECKLIST_001','ECBM_WO_001','ECBM_EQUIP_001')) {
    $u = "$origin/emars-script/$n.js"
    try {
        $r = Invoke-WebRequest -UseBasicParsing -Uri $u -TimeoutSec 10
        if ($r.StatusCode -eq 200 -and $r.Content.Length -gt 200) {
            "  FOUND  $n.js  ($($r.Content.Length) chars)"
            $r.Content | Set-Content "$outDir\02-$n.js" -Encoding UTF8
        }
    } catch { "  miss   $n.js" }
}

# 4. Fetch the eCBM module shell itself (from the !vm_main_pkg.main route)
"`nFetching eCBM module shell..."
$shellUrl = "$base/!vm_main_pkg.main?pSiteID=$site&pUser=$user&pModule=ECBM_MAIN_001"
"GET $shellUrl"
try {
    $shell = Invoke-WebRequest -UseBasicParsing -Uri $shellUrl -TimeoutSec 20
    "  status: $($shell.StatusCode)  length: $($shell.Content.Length)"
    "  looks like home shell? $($shell.Content -match 'vm_hompage_prc\.js')"
    $shell.Content | Set-Content "$outDir\03-ecbm-shell.html" -Encoding UTF8
} catch {
    "  FAILED: $($_.Exception.Message)"
}

"`nSaved to: $outDir"
Get-ChildItem $outDir | ForEach-Object { "  $($_.Name)  ($($_.Length) bytes)" }