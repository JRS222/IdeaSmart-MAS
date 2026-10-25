# Test harness for emars-fetch-layer.ps1
# Simulates the eCBM endpoints with captured/synthesized XML and
# verifies the conversion + archive pipeline without any network or UI.
$ErrorActionPreference = 'Stop'

$probeDir = 'C:\Users\JR\Downloads\IdeaSmart-MAS-main\IdeaSmart-MAS-main\emars-pm-probe-20261024-141529'
$archive  = Join-Path $env:TEMP ("pm_emars_test_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $archive -Force | Out-Null

# ---- stubs for UI-Script.ps1 dependencies -------------------------------
$script:config = [PSCustomObject]@{
    RootDirectory = $env:TEMP
    WorkTracking  = [PSCustomObject]@{
        EmarsOrigin  = 'http://emarssu3.eng.usps.gov'
        EmarsUser    = '1601089'
        EmarsSite    = '345'
        EmarsSubsite = '01'
        EmarsFacId   = ''
    }
}

function Write-Log { param([string]$m) Write-Host "LOG: $m" }
function Save-DebugHtml { param([string]$Html, [string]$Tag) Write-Host "DUMP: $Tag" }
function Get-PmArchiveDir { return $archive }

$script:getAchecklistVariant = 'named'   # or 'positional'

function Invoke-PmHttpGet {
    param([string]$Url, [int]$TimeoutSec = 30)
    Write-Host "GET $Url"

    if ($Url -match 'getAllChecklists') {
        return [System.IO.File]::ReadAllText((Join-Path $probeDir '05-getAllChecklists-ss01.xml'))
    }
    if ($Url -match 'getOpenlist') {
        if ($script:noOpen) { return '<?xml version="1.0"?><ROWSET></ROWSET>' }
        return @'
<?xml version="1.0"?>
<ROWSET>
   <ROW>
    <TRACKNO>37900603</TRACKNO>
    <X1>827433</X1>
    <X2>DBCS</X2>
    <X3>CK</X3>
    <OPEN_YN>1</OPEN_YN>
    <X5></X5>
    <X6></X6>
    <DATE_CREATED>10/24/2026</DATE_CREATED>
   </ROW>
</ROWSET>
'@
    }
    if ($Url -match 'getAChecklist') {
        if ($Url -match 'p_trackno=99999') { return '<?xml version="1.0"?><ROWSET></ROWSET>' }
        if ($Url -match 'p_trackno=88888') { return '<HTML><BODY>not a checklist</BODY></HTML>' }
        if ($script:getAchecklistVariant -eq 'named') {
            return @'
<?xml version="1.0"?>
<ROWSET>
   <ROW>
    <STEPID>101</STEPID>
    <COMPONENT>Reader/Stacker</COMPONENT>
    <STEP>&#60;p&#62;&#60;b&#62;Vacuum&#47;clean&#32;top&#32;of&#32;Reader&#32;and&#32;Stacker&#32;Modules.&#60;&#47;b&#62;&#60;&#47;p&#62;</STEP>
    <X3></X3>
    <NOTE></NOTE>
    <MINUTES>23</MINUTES>
    <X6></X6>
    <SKILLDESC>MM</SKILLDESC>
    <POWERONOFF>OFF</POWERONOFF>
    <PREROUTE></PREROUTE>
    <DATE_DUE>10/24/2026</DATE_DUE>
    <STATUS>pending</STATUS>
    <D2L>0</D2L>
    <ROUTETYPE>Preventive Maintenance</ROUTETYPE>
    <PERUNIT>F</PERUNIT>
    <ITEMNO>1</ITEMNO>
    <CAT>VACUUM/CLEAN</CAT>
    <CRITICAL></CRITICAL>
   </ROW>
   <ROW>
    <STEPID>102</STEPID>
    <COMPONENT>System</COMPONENT>
    <STEP>&#60;p&#62;&#60;b&#62;Run&#32;system.&#60;&#47;b&#62;&#60;&#47;p&#62;</STEP>
    <X3></X3>
    <NOTE></NOTE>
    <MINUTES>4</MINUTES>
    <X6></X6>
    <SKILLDESC>ET</SKILLDESC>
    <POWERONOFF>ON</POWERONOFF>
    <PREROUTE></PREROUTE>
    <DATE_DUE>10/25/2026</DATE_DUE>
    <STATUS>pending</STATUS>
    <D2L>1</D2L>
    <ROUTETYPE>Preventive Maintenance</ROUTETYPE>
    <PERUNIT>F</PERUNIT>
    <ITEMNO>2</ITEMNO>
    <CAT>COMPUTERS</CAT>
    <CRITICAL>Y</CRITICAL>
   </ROW>
   <ROW>
    <STEPID>103</STEPID>
    <COMPONENT>Conveyor</COMPONENT>
    <STEP>&#60;p&#62;&#60;b&#62;Grease&#32;bearings.&#60;&#47;b&#62;&#60;&#47;p&#62;</STEP>
    <X3></X3>
    <NOTE></NOTE>
    <MINUTES>10</MINUTES>
    <X6></X6>
    <SKILLDESC>MM</SKILLDESC>
    <POWERONOFF>OFF</POWERONOFF>
    <PREROUTE></PREROUTE>
    <DATE_DUE>10/20/2026</DATE_DUE>
    <STATUS>completed</STATUS>
    <D2L>5</D2L>
    <ROUTETYPE>Preventive Maintenance</ROUTETYPE>
    <PERUNIT>F</PERUNIT>
    <ITEMNO>3</ITEMNO>
    <CAT>BEARINGS</CAT>
    <CRITICAL></CRITICAL>
   </ROW>
</ROWSET>
'@
        } else {
            return @'
<?xml version="1.0"?>
<ROWSET>
   <ROW>
    <C0>101</C0><C1>Reader/Stacker</C1>
    <C2>&#60;p&#62;&#60;b&#62;Vacuum&#47;clean&#32;top&#32;of&#32;Reader&#32;and&#32;Stacker&#32;Modules.&#60;&#47;b&#62;&#60;&#47;p&#62;</C2>
    <C3></C3><C4></C4><C5>23</C5><C6></C6><C7>MM</C7><C8>OFF</C8><C9></C9>
    <C10>10/24/2026</C10><C11>pending</C11><C12>0</C12><C13>Preventive Maintenance</C13>
    <C14>F</C14><C15>1</C15><C16>VACUUM/CLEAN</C16><C17></C17>
   </ROW>
   <ROW>
    <C0>102</C0><C1>System</C1>
    <C2>&#60;p&#62;&#60;b&#62;Run&#32;system.&#60;&#47;b&#62;&#60;&#47;p&#62;</C2>
    <C3></C3><C4></C4><C5>4</C5><C6></C6><C7>ET</C7><C8>ON</C8><C9></C9>
    <C10>10/25/2026</C10><C11>pending</C11><C12>1</C12><C13>Preventive Maintenance</C13>
    <C14>F</C14><C15>2</C15><C16>COMPUTERS</C16><C17>Y</C17>
   </ROW>
   <ROW>
    <C0>103</C0><C1>Conveyor</C1>
    <C2>&#60;p&#62;&#60;b&#62;Grease&#32;bearings.&#60;&#47;b&#62;&#60;&#47;p&#62;</C2>
    <C3></C3><C4></C4><C5>10</C5><C6></C6><C7>MM</C7><C8>OFF</C8><C9></C9>
    <C10>10/20/2026</C10><C11>completed</C11><C12>5</C12><C13>Preventive Maintenance</C13>
    <C14>F</C14><C15>3</C15><C16>BEARINGS</C16><C17></C17>
   </ROW>
</ROWSET>
'@
        }
    }
    if ($Url -match 'getBypassCodes') {
        return @'
<?xml version="1.0"?>
<ROWSET>
   <ROW><CODE>1</CODE><DESC>One</DESC></ROW>
   <ROW><CODE>2</CODE><DESC>Two</DESC></ROW>
   <ROW><CODE>7</CODE><DESC>Seven</DESC></ROW>
   <ROW><CODE>W</CODE><DESC>W</DESC></ROW>
</ROWSET>
'@
    }
    if ($Url -match 'getMachineDesc') {
        return 'DBCS — Input/Output Subsystem, MS-298'
    }
    throw "Unstubbed URL: $Url"
}

# ---- load the layer ------------------------------------------------------
. "C:\Users\JR\Documents\deepseek-harness\default-workspace\emars-fetch-layer.ps1"

$failures = @()

function Assert {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { Write-Host "PASS: $Msg" -ForegroundColor Green }
    else { Write-Host "FAIL: $Msg" -ForegroundColor Red; $script:failures += $Msg }
}

# ---- 1. machine list ------------------------------------------------------
$machines = Get-EmarsMachineList -Acronym 'DBCS'
Assert ($machines.Count -eq 36) "machine list has 36 rows (got $($machines.Count))"
$m16 = $machines | Where-Object { $_.No -eq '16' } | Select-Object -First 1
Assert ($null -ne $m16 -and $m16.Ptr -eq '827433' -and $m16.ClassCode -eq 'CK') "machine DBCS 16 resolves (ptr=$($m16.Ptr) class=$($m16.ClassCode))"
Assert ($m16.SiteName -eq 'PITTSBURGH PDC') "site name from list"

# ---- 2. open checklist -----------------------------------------------------
$open = Get-EmarsOpenChecklist -MachinePtr '827433' -Acronym 'DBCS' -ClassCode 'CK'
Assert ($null -ne $open -and $open.TrackNo -eq '37900603') "open checklist found (trackno=$($open.TrackNo))"
Assert ($open.CreatedDate -eq '10/24/2026') "created date read"

# ---- 3. date normalizer -----------------------------------------------------
Assert ((ConvertTo-PmDateStamp -Text '10/24/2026') -eq '2026-10-24') "date normalize M/d/yyyy"
Assert ((ConvertTo-PmDateStamp -Text 'not a date') -eq $null) "date normalize garbage -> null"

# ---- 4. converter: named variant ---------------------------------------------
$script:getAchecklistVariant = 'named'
$xmlA  = Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_checklist.getAChecklist' -Query 'x=1' -Label 'test'
$rowsA = Get-EmarsRows -Doc $xmlA
$convA = ConvertFrom-EmarsChecklistXml -Rows $rowsA -Acronym 'DBCS' -MachinePtr '827433' `
    -MachineNo '16' -ClassCode 'CK' -SiteName 'PITTSBURGH PDC' -TrackNo '37900603' `
    -CreatedDate '10/24/2026' -MachineDesc 'DBCS I/O'

Assert (@($convA.Tasks).Count -eq 3) "named variant: 3 tasks (got $(@($convA.Tasks).Count))"
$t1 = @($convA.Tasks)[0]
Assert ($t1.PowerState -eq 'Power Off' -and $t1.ItemNo -eq '1' -and $t1.EstTimeMin -eq 23) "task1 power/item/min"
Assert ($t1.Component -eq 'Reader/Stacker - VACUUM/CLEAN') "task1 component-category join (got '$($t1.Component)')"
Assert ($t1.MinSkillLevel -eq 'MM' -and $t1.DueDate -eq '10/24/2026') "task1 skill/due"
Assert ($t1.InstructionHtml -match '<b>Vacuum') "task1 instruction HTML decoded (got '$($t1.InstructionHtml)')"
Assert ($t1.InstructionText -eq 'Vacuum/clean top of Reader and Stacker Modules.') "task1 instruction text (got '$($t1.InstructionText)')"
Assert ($t1.BypassCode -eq '1,2,7,W') "task1 bypass codes joined (got '$($t1.BypassCode)')"
$t2 = @($convA.Tasks)[1]
Assert ($t2.PowerState -eq 'Power On') "task2 power on"
Assert ($convA.Meta.EstOpenHours -eq 0.45 -and $convA.Meta.EstDueTodayHours -eq 0.45) "est totals (open=$($convA.Meta.EstOpenHours) due=$($convA.Meta.EstDueTodayHours))"
Assert ($convA.Meta.ChecklistDate -eq '2026-10-24') "meta checklist date normalized"

# ---- 5. converter: positional variant (should match named) --------------------
$script:getAchecklistVariant = 'positional'
$xmlB  = Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_checklist.getAChecklist' -Query 'x=1' -Label 'test'
$rowsB = Get-EmarsRows -Doc $xmlB
$convB = ConvertFrom-EmarsChecklistXml -Rows $rowsB -Acronym 'DBCS' -MachinePtr '827433' `
    -MachineNo '16' -ClassCode 'CK' -SiteName 'PITTSBURGH PDC' -TrackNo '37900603' `
    -CreatedDate '10/24/2026' -MachineDesc 'DBCS I/O'

$jsonA = ($convA | ConvertTo-Json -Depth 12)
$jsonB = ($convB | ConvertTo-Json -Depth 12)
Assert ($jsonA -eq $jsonB) "positional fallback yields identical output"

# ---- 6. end-to-end import (archives files) ------------------------------------
$script:getAchecklistVariant = 'named'
$summary = Import-PmChecklistFromEmars -Machine $m16
Assert ($summary.MachineId -eq 'DBCS 16') "summary machine id (got '$($summary.MachineId)')"
Assert ($summary.TaskCount -eq 3) "summary task count"
Assert ($summary.DateStamp -eq '2026-10-24') "summary date stamp"
Assert (Test-Path $summary.HtmlPath) "html archived"
Assert (Test-Path $summary.JsonPath) "json archived"

$json = Get-Content -Path $summary.JsonPath -Raw | ConvertFrom-Json
Assert ($json.Meta.Acronym -eq 'DBCS' -and $json.Meta.ChecklistNo -eq '37900603') "archived json meta"
Assert (@($json.Tasks).Count -eq 3) "archived json tasks"

# ---- 6b. worksheet Reg-WO No resolves directly -------------------------------
$script:getAchecklistVariant = 'named'
$sDirect = Import-PmChecklistFromEmars -Machine $m16 -TrackNo '37900603' -WorksheetChecklistNo '37860213'
Assert ($sDirect.ChecklistNo -eq '37900603') "direct trackno fetch uses worksheet number (got '$($sDirect.ChecklistNo)')"

# ---- 6c. worksheet Reg-WO No is a WO pointer (empty) -> openlist fallback ----
$sFallback = Import-PmChecklistFromEmars -Machine $m16 -TrackNo '99999' -WorksheetChecklistNo '000'
Assert ($sFallback.ChecklistNo -eq '37900603' -and $sFallback.DateStamp -eq '2026-10-24') "WO-pointer number falls back to open checklist (got '$($sFallback.ChecklistNo)'/'$($sFallback.DateStamp)')"

# ---- 6d. worksheet number errors out (non-XML) -> openlist fallback ----------
$sErr = Import-PmChecklistFromEmars -Machine $m16 -TrackNo '88888' -WorksheetChecklistNo '000'
Assert ($sErr.ChecklistNo -eq '37900603') "non-XML trackno falls back to open checklist"

# ---- 6e. nothing open at all -> clear messages --------------------------------
$script:noOpen = $true
$threw000 = $false
try { Import-PmChecklistFromEmars -Machine $m16 -TrackNo '99999' -WorksheetChecklistNo '000' | Out-Null }
catch { $threw000 = ($_.Exception.Message -match 'Checklist No\. 000') }
Assert $threw000 "not-created case reports 'Checklist No. 000'"
$threwNone = $false
try { Import-PmChecklistFromEmars -Machine $m16 -TrackNo '99999' -WorksheetChecklistNo '' | Out-Null }
catch { $threwNone = ($_.Exception.Message -match 'No open checklist') }
Assert $threwNone "no-open case reports 'No open checklist'"
$script:noOpen = $false

# ---- 7. generated HTML satisfies the paste parser's key regexes ----------------
# The parser runs these regexes on innerText (tags stripped, entities decoded).
$html = [System.IO.File]::ReadAllText($summary.HtmlPath)
$text = $html -replace '(?i)<br\s*/?>', ' '
$text = $text -replace '(?s)<[^>]+>', ''
$text = [System.Net.WebUtility]::HtmlDecode($text)
$flat = ($text -replace '\s+', ' ').Trim()
Assert ($flat -match 'Checklist No\.\s*:\s*(\d+)' -and $Matches[1] -eq '37900603') "html: Checklist No. marker (got '$flat')"
Assert ($flat -match 'Skill\(s\)\s*:\s*([^,]+)') "html: Skill(s) marker"
Assert ($flat -match 'Time\s*:\s*([^)]+?)\s*\)') "html: Time ) marker"
Assert ($flat -match 'Equipment Acronym') "html: Equipment Acronym header"
Assert ($flat -match 'Site Name') "html: Site Name header"
Assert ($flat -match 'for All Open Tasks:\s*([\d.]+)') "html: All Open Tasks total"
Assert ($flat -match 'for Tasks Due Today:\s*([\d.]+)') "html: Tasks Due Today total"
Assert ($flat -match 'Item No' -and $flat -match 'Task Statement and Instruction') "html: task table header"
$rowCount = ([regex]::Matches($html, '<td width="5%" align="right">1</td>')).Count
Assert ($rowCount -ge 1) "html: task row 1 present"

Write-Host ''
if ($script:failures.Count -eq 0) {
    Write-Host "ALL TESTS PASSED" -ForegroundColor Green
} else {
    Write-Host "$($script:failures.Count) FAILURE(S):" -ForegroundColor Red
    $script:failures | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
}
Write-Host "Archive written to: $archive"
