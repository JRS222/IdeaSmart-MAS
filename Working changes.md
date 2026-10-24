$origin = 'http://emarssu3.eng.usps.gov'
$outDir = "$env:USERPROFILE\Desktop\ecbm-checklist-probe"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$jsUrl = "$origin/emars-script/ECBM_CHECKLISTS.js"
"GET $jsUrl"
$js = Invoke-WebRequest -UseBasicParsing -Uri $jsUrl -TimeoutSec 20
"  status: $($js.StatusCode)   length: $($js.Content.Length)"
$js.Content | Set-Content "$outDir\ECBM_CHECKLISTS.js" -Encoding UTF8

# Scan for the URL pattern
"`nURL candidates:"
$patterns = @(
    'em_[a-z_]*checklist[a-z_]*\.[a-zA-Z0-9_]+',
    'ecbm\.[a-zA-Z0-9_]+',
    'p_checklist[a-zA-Z_]*',
    'p_equip[a-zA-Z_]*',
    'p_wo\b',
    'openSwilWin\([^)]+\)',
    'getInnerHTML\([^)]+\)'
)
foreach ($pat in $patterns) {
    foreach ($m in [regex]::Matches($js.Content, $pat, 'IgnoreCase')) {
        "  {0}" -f $m.Value
    }
} | Sort-Object -Unique