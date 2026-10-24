$user   = '1601089'
$site   = '345'
$base   = 'http://emarssu3.eng.usps.gov/pemars'

# 1. Confirm cookieless home-page access (proves the earlier finding)
$home = Invoke-WebRequest -UseBasicParsing -Uri "$base/vm_hompage_prc?p_user=$user&p_site=$site"
"home length: $($home.Content.Length)  (expect ~5-6k, contains 'EM_PREV_MAINT_001' if OK)"
"is auth'd:   $($home.Content -match 'EM_PREV_MAINT_001')"

# 2. Fetch the PM module shell
$pm = Invoke-WebRequest -UseBasicParsing -Uri "$base/!vm_main_pkg.main?pUser=$user&pSiteID=$site&pModule=EM_PREV_MAINT_001"
"pm   length: $($pm.Content.Length)"
"is auth'd:   $($pm.Content -match 'PM|CHECKLIST|MAINT')"
$pm.Content | Set-Content -Path "$env:TEMP\pm_module.html" -Encoding UTF8
"pm body -> $env:TEMP\pm_module.html"