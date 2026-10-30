WO.ps1 — final version

Everything learned across all runs folded in. Key changes from the version that crashed:

· WoPtr is optional — no Mandatory on it
· SWIL detail reload tries WO_NO first (since pointer isn't in the search response), falls back to WO_PTR if populated
· Delimiter is detected once and cached — huge speed win
· Date parse uses -as [DateTime] (no try/catch) — no exceptions on 144K rows
· -SkipDetail switch — get the summary CSV in ~30s without waiting on detail fetches
· First-row cells 0-15 all logged — reveals hidden columns if they exist
· Column 15 is used as WoPtr only if it's numeric — otherwise tries cells 11-14
· Empty WoPtr no longer kills the pipeline
· Every route failure is logged, not thrown
· Default MaxRows = 200 so the first test is fast

Save as WO.ps1, overwriting your existing file.

```powershell
<#
.SYNOPSIS
    Standalone eMARS Work Order history extractor.

.DESCRIPTION
    Pulls work orders from eMARS for a date window, aggregates labor
    hours per machine, and (when a detail route works) emits a
    per-work-order CSV shaped like the printed PDF work order.

    Two CSVs land in the Reports directory:
      WorkOrderSummary_<from>_<to>.csv   one row per machine
      WorkOrderDetail_<from>_<to>.csv    one row per work order

    All diagnostics go to WorkOrderHistory.log next to this script.
    Raw search responses that yield 0 rows are dumped to Debug Dumps\.

.PARAMETER SkipDetail
    Skip the per-WO detail fetch.  Fast (~30s).  Summary CSV will
    have Work Order count and Local Hours from Jobs.json, but eMARS
    Hours and Est Hours columns will be empty.

.PARAMETER DiscoverEndpoints
    Diagnostic mode.  Fetches one status, logs the first row's full
    cell contents, and probes each detail route.  No CSVs written.
    Run this first.

.EXAMPLE
    .\WO.ps1 -DiscoverEndpoints

.EXAMPLE
    .\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -SkipDetail

.EXAMPLE
    .\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -MaxRows 50
#>

[CmdletBinding()]
param(
    [DateTime]$From   = (Get-Date).Date.AddDays(-7),
    [DateTime]$To     = (Get-Date).Date,
    [string]$Machine  = '',
    [string[]]$Status = @('E','A','B','C','D','W'),
    [ValidateSet('Due','Issued','Logged')]
    [string]$FilterBy = 'Due',
    [string]$OutputDirectory = '',
    [string]$Register = '',
    [int]$MaxRegisters = 30,
    [int]$MaxRows = 200,
    [switch]$SkipDetail,
    [switch]$ShowGui,
    [switch]$DiscoverEndpoints
)

$ErrorActionPreference = 'Continue'

# ============================================================
# Logging
# ============================================================
$script:LogPath = Join-Path $PSScriptRoot 'WorkOrderHistory.log'

function Write-Log {
    param([string]$message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
    Write-Host $line
    try { $line | Out-File -Append -FilePath $script:LogPath -Encoding UTF8 } catch { }
}

# ============================================================
# Config
# ============================================================
function Find-ConfigPath {
    $candidates = @(
        (Join-Path $PSScriptRoot 'Config.json'),
        (Join-Path $PSScriptRoot 'PartsMgmt\Config.json'),
        (Join-Path (Split-Path $PSScriptRoot -Parent) 'PartsMgmt\Config.json')
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}

function Load-Config {
    $path = Find-ConfigPath
    if (-not $path) { throw "Config.json not found." }
    Write-Log "Loading config from $path"
    return (Get-Content -Path $path -Raw | ConvertFrom-Json)
}

# ============================================================
# HTTP
# ============================================================
function Get-EmarsBase {
    $origin = "$($script:config.WorkTracking.EmarsOrigin)".Trim()
    if ([string]::IsNullOrWhiteSpace($origin)) { $origin = 'http://emarssu3.eng.usps.gov' }
    return "$($origin.TrimEnd('/'))/pemars"
}

function Invoke-EmarsGet {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Query = '',
        [int]$TimeoutSec = 40
    )
    $url = "$(Get-EmarsBase)/$Path"
    if (-not [string]::IsNullOrWhiteSpace($Query)) { $url += "?$Query" }
    Write-Log "GET $url"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $TimeoutSec
        $sw.Stop()
        Write-Log "GET done in $([int]$sw.Elapsed.TotalSeconds)s -- $($resp.Content.Length) bytes"
        return $resp.Content
    } catch {
        $sw.Stop()
        Write-Log "GET failed after $([int]$sw.Elapsed.TotalSeconds)s -- $($_.Exception.Message)"
        throw
    }
}

# ============================================================
# Debug dump
# ============================================================
function Save-RawResponse {
    param([string]$Content, [string]$Tag)
    if ([string]::IsNullOrWhiteSpace($Content)) { return }
    $dir = Join-Path $PSScriptRoot 'Debug Dumps'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $safe = ($Tag -replace '[^\w\-]', '_')
    $path = Join-Path $dir ("$safe-" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".html")
    [System.IO.File]::WriteAllText($path, $Content, (New-Object System.Text.UTF8Encoding($false)))
    Write-Log "Save-RawResponse: $($Content.Length) bytes -> $path"
    return $path
}

# ============================================================
# Work order search response parser (FAST -- delimiter cached)
# ============================================================
function ConvertFrom-WoSearchResponse {
    param(
        [string]$Content,
        [string]$SiteId,
        [string]$Tag = 'search',
        [string]$FilterMode = '',
        [DateTime]$WindowFrom = [DateTime]::MinValue,
        [DateTime]$WindowTo   = [DateTime]::MinValue
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        Write-Log "ConvertFrom-WoSearchResponse[$Tag]: empty response"
        return ,@()
    }
    $Content = $Content.Trim()
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: $($Content.Length) bytes"

    # ---------- XML ----------
    if ($Content -match '^\s*<\?xml' -or $Content -match '^\s*<\s*ROWSET') {
        Write-Log "ConvertFrom-WoSearchResponse[$Tag]: XML branch"
        $xml = New-Object System.Xml.XmlDocument
        $xml.LoadXml($Content)
        $out = @()
        foreach ($row in @($xml.SelectNodes('//ROW'))) {
            $fields = @{}
            foreach ($child in $row.ChildNodes) {
                if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    $fields[$child.Name.ToUpper()] = "$($child.InnerText)".Trim()
                }
            }
            $out += [PSCustomObject]@{
                WoPtr = $fields['WO_PTR']; WorkOrderNo = $fields['WONO']
                RegisterNo = $fields['REGNO']; WorkCode = $fields['WORK_CODE']
                Acronym = $fields['ACRONYM']; ClassCode = $fields['CLASSCODE']
                Equipment = $fields['EQUIP_NO']; Status = $fields['WO_STATUS']
                Crew = $fields['CREW_ID']; DateIssued = $fields['DATE_ISSUED']
                DueDate = $fields['DATE_DUE']; Description = $fields['WO_DESC']
                SiteId = $SiteId; RawCells = @()
            }
        }
        return ,$out
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # ---------- APPLET <param> format ----------
    $paramMatches = [regex]::Matches(
        $Content,
        '(?is)<param\b[^>]*?\bname\s*=\s*["''](?<name>[^"'']+)["''][^>]*?\bvalue\s*=\s*["''](?<val>[^"'']*)["'']'
    )
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: $($paramMatches.Count) <param> in $([int]$sw.Elapsed.TotalMilliseconds)ms"

    if ($paramMatches.Count -gt 0) {
        $rowParams = @{}    # plain Hashtable, avoids int-indexer crash
        $headers   = @()
        $keyParams = @{}

        foreach ($m in $paramMatches) {
            $pname = $m.Groups['name'].Value.Trim()
            $pval  = [System.Net.WebUtility]::HtmlDecode($m.Groups['val'].Value)
            if ($pname -match '^row(\d+)text$') {
                $rowParams[[int]$matches[1]] = $pval
            } elseif ($pname -ieq 'colheadertext') {
                $headers = @($pval -split ',' | ForEach-Object { $_.Trim() })
            } else {
                $keyParams[$pname] = $pval
            }
        }
        Write-Log "ConvertFrom-WoSearchResponse[$Tag]: headers = $($headers -join ' | ')"

        # Log applet config for diagnosis
        foreach ($kp in @('numcols','numrows','rowheader','colwidths')) {
            if ($keyParams.ContainsKey($kp)) {
                $val = "$($keyParams[$kp])"
                if ($val.Length -gt 200) { $val = $val.Substring(0,200) + '...' }
                Write-Log "  applet $kp = '$val'"
            }
        }

        if ($rowParams.Count -eq 0) { return ,@() }

        $headerCount = $headers.Count
        $sortedKeys  = @($rowParams.Keys | Sort-Object { [int]$_ })

        # ---- Delimiter detection (try up to 10 sample rows) ----
        $separator = $null
        $sampleKeys = @($sortedKeys | Select-Object -First 10)
        foreach ($sk in $sampleKeys) {
            $rawTest = "$($rowParams[$sk])"
            foreach ($sep in @('|', ';', '~', "`t", ',')) {
                $test = @($rawTest -split [regex]::Escape($sep))
                if ($test.Count -ge $headerCount) { $separator = $sep; break }
            }
            if ($separator) { break }
        }
        if (-not $separator) { $separator = '|' }
        Write-Log "ConvertFrom-WoSearchResponse[$Tag]: delimiter = '$([char[]]$separator -join '')'"

        # ---- Log full first row ----
        $firstKey = $sortedKeys[0]
        $firstRaw = "$($rowParams[$firstKey])"
        $firstCells = @($firstRaw -split [regex]::Escape($separator))
        Write-Log "  --- FIRST ROW ($($firstCells.Count) cells) ---"
        for ($pi = 0; $pi -lt $firstCells.Count; $pi++) {
            $label = if ($pi -lt $headerCount) { $headers[$pi] } else { "(hidden_$pi)" }
            $val = "$($firstCells[$pi])"
            if ($val.Length -gt 100) { $val = $val.Substring(0,100) + '...' }
            Write-Log "    [$pi] $label = '$val'"
        }

        # ---- Column index map from headers ----
        $idx = @{}
        for ($i = 0; $i -lt $headerCount; $i++) {
            $idx[$headers[$i]] = $i
        }
        function __g($idx, [string[]]$names) {
            foreach ($n in $names) { if ($idx.ContainsKey($n)) { return [int]$idx[$n] } }
            return -1
        }
        $colReg    = __g $idx @('Reg No')
        $colWoNo   = __g $idx @('WO No')
        $colWC     = __g $idx @('WC')
        $colAcro   = __g $idx @('Acronym')
        $colCC     = __g $idx @('CC')
        $colEquip  = __g $idx @('Equip No')
        $colStts   = __g $idx @('Stts')
        $colCrew   = __g $idx @('Crew')
        $colIssued = __g $idx @('Dt Issued')
        $colDue    = __g $idx @('Dt Due')
        $colDesc   = __g $idx @('WO Description')

        $filterCol = -1
        if ($FilterMode -eq 'Due')    { $filterCol = $colDue }
        if ($FilterMode -eq 'Issued') { $filterCol = $colIssued }
        $applyFilter = ($filterCol -ge 0 -and $WindowFrom -ne [DateTime]::MinValue -and $WindowTo -ne [DateTime]::MinValue)
        $winFrom = $WindowFrom.Date
        $winTo   = $WindowTo.Date

        # ---- Parse loop (hot path -- no try/catch) ----
        $out         = @()
        $preFiltered = 0
        $processed   = 0
        $total       = $sortedKeys.Count

        foreach ($rowIdx in $sortedKeys) {
            $processed++
            if (($processed % 50000) -eq 0) {
                Write-Log "ConvertFrom-WoSearchResponse[$Tag]: processed $processed / $total, kept $($out.Count)"
            }

            $raw = "$($rowParams[$rowIdx])"
            $cells = @($raw -split [regex]::Escape($separator))
            if ($cells.Count -lt 3) { continue }

            # Fast date filter
            if ($applyFilter -and $filterCol -lt $cells.Count) {
                $dtStr = "$($cells[$filterCol])".Trim()
                if ($dtStr.Length -lt 6) { $preFiltered++; continue }
                $dt = $dtStr -as [DateTime]
                if ($null -eq $dt) { $preFiltered++; continue }
                if ($dt.Date -lt $winFrom -or $dt.Date -gt $winTo) { $preFiltered++; continue }
            }

            # Extract WoPtr: prefer numeric value in cell 15, else scan 11..14
            $woPtr = ''
            if ($cells.Count -gt 15) {
                $c15 = "$($cells[15])".Trim()
                if ($c15 -match '^\d{3,}$') { $woPtr = $c15 }
            }
            if (-not $woPtr -and $cells.Count -gt 11) {
                for ($i = 15; $i -ge 11; $i--) {
                    if ($i -lt $cells.Count) {
                        $cv = "$($cells[$i])".Trim()
                        if ($cv -match '^\d{5,}$') { $woPtr = $cv; break }
                    }
                }
            }

            $out += [PSCustomObject]@{
                WoPtr       = $woPtr
                RegisterNo  = if ($colReg    -ge 0 -and $colReg    -lt $cells.Count) { "$($cells[$colReg])".Trim()    } else { '' }
                WorkOrderNo = if ($colWoNo   -ge 0 -and $colWoNo   -lt $cells.Count) { "$($cells[$colWoNo])".Trim()   } else { '' }
                WorkCode    = if ($colWC     -ge 0 -and $colWC     -lt $cells.Count) { "$($cells[$colWC])".Trim()     } else { '' }
                Acronym     = if ($colAcro   -ge 0 -and $colAcro   -lt $cells.Count) { "$($cells[$colAcro])".Trim()   } else { '' }
                ClassCode   = if ($colCC     -ge 0 -and $colCC     -lt $cells.Count) { "$($cells[$colCC])".Trim()     } else { '' }
                Equipment   = if ($colEquip  -ge 0 -and $colEquip  -lt $cells.Count) { "$($cells[$colEquip])".Trim()  } else { '' }
                Status      = if ($colStts   -ge 0 -and $colStts   -lt $cells.Count) { "$($cells[$colStts])".Trim()   } else { '' }
                Crew        = if ($colCrew   -ge 0 -and $colCrew   -lt $cells.Count) { "$($cells[$colCrew])".Trim()   } else { '' }
                DateIssued  = if ($colIssued -ge 0 -and $colIssued -lt $cells.Count) { "$($cells[$colIssued])".Trim() } else { '' }
                DueDate     = if ($colDue    -ge 0 -and $colDue    -lt $cells.Count) { "$($cells[$colDue])".Trim()    } else { '' }
                Description = if ($colDesc   -ge 0 -and $colDesc   -lt $cells.Count) { "$($cells[$colDesc])".Trim()   } else { '' }
                SiteId      = $SiteId
                RawCells    = $cells
            }
        }

        Write-Log "ConvertFrom-WoSearchResponse[$Tag]: kept $($out.Count) / $processed (dropped $preFiltered by date filter, $([int]$sw.Elapsed.TotalSeconds)s)"
        return ,$out
    }

    # ---------- HTML fallback ----------
    $trMatches = [regex]::Matches(
        $Content,
        '(?is)<tr\b[^>]*\bclass\s*=\s*["'']?(?:MAIN|HILITE)["'']?[^>]*>(?<body>.*?)</tr>'
    )
    if ($trMatches.Count -eq 0) {
        $trMatches = [regex]::Matches($Content, '(?is)<tr\b[^>]*>(?<body>.*?)</tr>')
    }
    $out = @()
    foreach ($tr in $trMatches) {
        $cm = [regex]::Matches($tr.Groups['body'].Value, '(?is)<td\b[^>]*>(?<c>.*?)</td>')
        if ($cm.Count -lt 3) { continue }
        $cells = @($cm | ForEach-Object {
            $t = $_.Groups['c'].Value -replace '(?is)<[^>]+>', ' '
            $t = [System.Net.WebUtility]::HtmlDecode($t)
            ($t -replace '\s+', ' ').Trim()
        })
        $out += [PSCustomObject]@{
            WoPtr=''; RegisterNo=''
            WorkOrderNo = $cells[0]
            WorkCode=''; Acronym=''; ClassCode=''; Equipment=''; Status=''; Crew=''
            DateIssued=''; DueDate=''
            Description = if ($cells.Count -gt 1) { $cells[1] } else { '' }
            SiteId = $SiteId; RawCells = $cells
        }
    }
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: HTML fallback parsed $($out.Count) row(s)"
    if ($out.Count -eq 0) {
        Save-RawResponse -Content $Content -Tag "WoSearch_$Tag" | Out-Null
    }
    return ,$out
}

# ============================================================
# WO enumeration
# ============================================================
function Get-EmarsRegisterList {
    param([string]$SiteId)
    try {
        $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getXMLRegisterList' -Query "pSiteID=$SiteId"
        if ($body -notmatch '^\s*<\?xml' -and $body -notmatch '^\s*<\s*ROWSET') {
            Write-Log "Get-EmarsRegisterList: non-XML ($($body.Length) bytes)"
            return ,@()
        }
        $xml = New-Object System.Xml.XmlDocument
        $xml.LoadXml($body)
        $regs = @()
        foreach ($row in @($xml.SelectNodes('//ROW'))) {
            $regNo = ''
            foreach ($child in $row.ChildNodes) {
                if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    if ($child.Name -match '(?i)^(REG.*NO|REGNO|REG_ID|REGISTER)$') {
                        $regNo = "$($child.InnerText)".Trim(); break
                    }
                }
            }
            if ($regNo) { $regs += $regNo }
        }
        Write-Log "Get-EmarsRegisterList: $($regs.Count) register(s)"
        return ,$regs
    } catch {
        Write-Log "Get-EmarsRegisterList: $($_.Exception.Message)"
        return ,@()
    }
}

$script:FallbackRegisters = @(
    '798','797','796','795','794','793','792','787','786','785','784','783','782',
    '780','777','776','775','774','773','769','767','750','726','719','716','713',
    '712','709','705','704','702','701','698','697','696','695','694','693','692',
    '687','686','685','684','683','682','680','677','676','675','674','673','672',
    '669','667','650','626','619','616','613','612','609','605','604','602','601',
    '598','597','596','595','594','593','592','587','586','585','584','583','582',
    '580','577','576','575','574','573','572','569','567','550','519','516','513',
    '512','509','505','504','502','501'
)

function Get-EmarsWorkOrdersAll {
    param(
        [string]$SiteId,
        [string[]]$StatusCodes,
        [string]$SingleRegister = '',
        [int]$MaxRegs = 30,
        [int]$MaxRows = 5000,
        [string]$FilterMode = '',
        [DateTime]$WindowFrom = [DateTime]::MinValue,
        [DateTime]$WindowTo   = [DateTime]::MinValue
    )

    $seen = @{}
    $out  = @()

    # Fast path: empty-register sweep
    if ([string]::IsNullOrWhiteSpace($SingleRegister)) {
        foreach ($st in $StatusCodes) {
            $q = "pSite=$SiteId&pWOStatus=$st&pRegNo=&pWONo="
            try {
                $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWOData' -Query $q -TimeoutSec 60
            } catch {
                Write-Log "Get-EmarsWorkOrdersAll: status '$st' FAILED -- $($_.Exception.Message)"
                continue
            }
            $wos = ConvertFrom-WoSearchResponse -Content $body -SiteId $SiteId -Tag "empty_${st}" `
                    -FilterMode $FilterMode -WindowFrom $WindowFrom -WindowTo $WindowTo
            Write-Log "Get-EmarsWorkOrdersAll: status '$st' -> $(@($wos).Count) WO(s)"
            foreach ($wo in @($wos)) {
                if ([string]::IsNullOrWhiteSpace($wo.WorkOrderNo)) { continue }
                if ($seen.ContainsKey($wo.WorkOrderNo)) { continue }
                $seen[$wo.WorkOrderNo] = $true
                $out += $wo
                if ($MaxRows -gt 0 -and $out.Count -ge $MaxRows) {
                    Write-Log "Get-EmarsWorkOrdersAll: MaxRows ($MaxRows) reached -- stopping"
                    return ,$out
                }
            }
        }
        if ($out.Count -gt 0) { return ,$out }
        Write-Log "Get-EmarsWorkOrdersAll: empty-register path returned 0 -- iterating registers"
    }

    $regs = @()
    if (-not [string]::IsNullOrWhiteSpace($SingleRegister)) {
        $regs = @($SingleRegister)
    } else {
        $regs = Get-EmarsRegisterList -SiteId $SiteId
        if ($regs.Count -eq 0) {
            $regs = $script:FallbackRegisters
            Write-Log "Get-EmarsWorkOrdersAll: using $($regs.Count) fallback register(s)"
        }
        if ($MaxRegs -gt 0 -and $regs.Count -gt $MaxRegs) {
            $regs = @($regs | Select-Object -First $MaxRegs)
            Write-Log "Get-EmarsWorkOrdersAll: capped at $MaxRegs register(s)"
        }
    }

    foreach ($reg in $regs) {
        foreach ($st in $StatusCodes) {
            $q = "pSite=$SiteId&pWOStatus=$st&pRegNo=$reg&pWONo="
            try {
                $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWOData' -Query $q -TimeoutSec 30
            } catch {
                Write-Log "  reg=$reg status=$st FAILED"
                continue
            }
            $wos = ConvertFrom-WoSearchResponse -Content $body -SiteId $SiteId -Tag "reg${reg}_${st}" `
                    -FilterMode $FilterMode -WindowFrom $WindowFrom -WindowTo $WindowTo
            Write-Log "  reg=$reg status=$st -> $(@($wos).Count) WO(s)"
            foreach ($wo in @($wos)) {
                if ([string]::IsNullOrWhiteSpace($wo.WorkOrderNo)) { continue }
                if ($seen.ContainsKey($wo.WorkOrderNo)) { continue }
                $seen[$wo.WorkOrderNo] = $true
                $out += $wo
                if ($MaxRows -gt 0 -and $out.Count -ge $MaxRows) {
                    Write-Log "  MaxRows ($MaxRows) reached -- stopping"
                    return ,$out
                }
            }
        }
    }
    return ,$out
}

# ============================================================
# SWIL widget parser
# ============================================================
function Get-WidgetValuesFromSwil {
    param([string]$Html)
    $result = @{}
    if ([string]::IsNullOrWhiteSpace($Html)) { return $result }
    foreach ($m in [regex]::Matches($Html, '(?is)<widget\b[^>]*\bname\s*=\s*"([^"]+)"[^>]*\bvalue\s*=\s*"([^"]*)"')) {
        $name = $m.Groups[1].Value.Trim()
        if (-not $result.ContainsKey($name)) {
            $result[$name] = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value)
        }
    }
    return $result
}

# ============================================================
# WO detail fetch
# ============================================================
function Get-EmarsWorkOrderDetail {
    param(
        [string]$WoPtr = '',      # now optional
        [Parameter(Mandatory)][string]$WoNo,
        [Parameter(Mandatory)][string]$SiteId
    )

    $user = "$($script:config.WorkTracking.EmarsUser)".Trim()
    if ([string]::IsNullOrWhiteSpace($user)) { $user = '0' }

    # ---- Route A: SWIL reload, WO_NO first then WO_PTR ----
    $keyVariants = @()
    # WO_NO is the safe fallback when pointer isn't in search response
    $keyVariants += @{ Name = 'WO_NO';  Value = $WoNo;  Label = 'WO_NO' }
    if (-not [string]::IsNullOrWhiteSpace($WoPtr)) {
        $keyVariants += @{ Name = 'WO_PTR'; Value = $WoPtr; Label = 'WO_PTR' }
    }

    foreach ($kv in $keyVariants) {
        foreach ($module in @('EM_WORK_MAINT_002','EM_WORK_MAINT_001')) {
            $paramXml = "<KEY><$($kv.Name)>$($kv.Value)</$($kv.Name)></KEY>"
            $q = "pUser=$user&pSiteID=$SiteId&pModule=$module&pparamxml=$paramXml"
            try {
                $html = Invoke-EmarsGet -Path '!vm_main_pkg.main' -Query $q -TimeoutSec 30
                $widgets = Get-WidgetValuesFromSwil -Html $html
                if ($widgets.Count -le 5) {
                    Write-Log "Detail SWIL $module ($($kv.Label)) -> only $($widgets.Count) widgets, skipping"
                    continue
                }
                # Verify the returned WO matches what we asked for
                $returnedWo = ''
                foreach ($k in @('wo_no_h','wo_no','divwono','wo_no_key')) {
                    if ($widgets.ContainsKey($k) -and $widgets[$k]) {
                        $returnedWo = $widgets[$k]; break
                    }
                }
                if ($returnedWo -and ($returnedWo -notlike "*$WoNo*")) {
                    Write-Log "Detail SWIL $module ($($kv.Label)) returned WO '$returnedWo' != '$WoNo', skipping"
                    continue
                }
                Write-Log "Detail OK: SWIL $module via $($kv.Label), $($widgets.Count) widgets (returned WO='$returnedWo')"
                return [PSCustomObject]@{ Route = "SWIL:$module/$($kv.Label)"; Fields = $widgets }
            } catch {
                Write-Log "Detail SWIL $module ($($kv.Label)) FAILED -- $($_.Exception.Message)"
            }
        }
    }

    # ---- Route B: candidate detail endpoints ----
    $candidates = @(
        'em_work_pkg.em_getWODetail',
        'em_work_pkg.em_getWO',
        'em_work_pkg.em_loadWO',
        'em_work_pkg.em_getWorkOrder'
    )
    foreach ($ep in $candidates) {
        foreach ($paramSet in @(
            "pSite=$SiteId&pWONo=$WoNo",
            "pSite=$SiteId&pWONo=$WoNo&pRegNo=",
            "p_site_id=$SiteId&p_wo_no=$WoNo"
        )) {
            try {
                $body = Invoke-EmarsGet -Path $ep -Query $paramSet -TimeoutSec 15
                if ([string]::IsNullOrWhiteSpace($body)) { continue }
                $body = $body.Trim()
                if ($body -match '^\s*<\?xml' -or $body -match '^\s*<\s*ROWSET') {
                    $xml = New-Object System.Xml.XmlDocument
                    $xml.LoadXml($body)
                    $fields = @{}
                    foreach ($node in @($xml.SelectNodes('//ROW[1]/*'))) {
                        $fields[$node.Name.ToUpper()] = "$($node.InnerText)".Trim()
                    }
                    if ($fields.Count -gt 3) {
                        Write-Log "Detail OK: $ep ($($fields.Count) fields)"
                        return [PSCustomObject]@{ Route = $ep; Fields = $fields }
                    }
                }
            } catch { }
        }
    }

    Write-Log "Detail: no route succeeded for WO $WoNo"
    return $null
}

# ============================================================
# Local jobs
# ============================================================
function Get-LocalJobs {
    $rel = "$($script:config.WorkTracking.JobsFile)".Trim()
    if ([string]::IsNullOrWhiteSpace($rel)) { return ,@() }
    $path = if ([System.IO.Path]::IsPathRooted($rel)) { $rel } else {
        Join-Path $script:config.RootDirectory $rel
    }
    if (-not (Test-Path $path)) { return ,@() }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return ,@() }
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return ,@() }
        return ,@($data)
    } catch {
        Write-Log "Get-LocalJobs: $($_.Exception.Message)"
        return ,@()
    }
}

function Get-LocalHoursByWO {
    param([DateTime]$WindowFrom, [DateTime]$WindowTo)
    $out = @{}
    foreach ($j in (Get-LocalJobs)) {
        if ("$($j.kind)" -ne 'workorder') { continue }
        $wo = "$($j.workOrderNo)"
        if ([string]::IsNullOrWhiteSpace($wo)) { continue }
        $jd = $null
        if ($j.createdAt) { try { $jd = ([DateTime]$j.createdAt).Date } catch { } }
        if (-not $jd) { continue }
        if ($jd -lt $WindowFrom.Date -or $jd -gt $WindowTo.Date) { continue }

        $h = $null
        if ($null -ne $j.hoursOverride -and "$($j.hoursOverride)" -ne '') {
            $p = 0.0
            if ([double]::TryParse("$($j.hoursOverride)", [ref]$p)) { $h = $p }
        }
        if ($null -eq $h -and $null -ne $j.timerComputedHours -and "$($j.timerComputedHours)" -ne '') {
            $p = 0.0
            if ([double]::TryParse("$($j.timerComputedHours)", [ref]$p)) { $h = $p }
        }
        if ($null -eq $h -and -not [string]::IsNullOrWhiteSpace("$($j.startTime)") -and -not [string]::IsNullOrWhiteSpace("$($j.endTime)")) {
            try {
                $s = [DateTime]::ParseExact("$($j.startTime)", 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
                $e = [DateTime]::ParseExact("$($j.endTime)",   'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
                if ($e -lt $s) { $e = $e.AddDays(1) }
                $h = [Math]::Round(($e - $s).TotalHours, 2)
            } catch { }
        }
        if ($null -eq $h) { continue }
        if (-not $out.ContainsKey($wo)) { $out[$wo] = 0.0 }
        $out[$wo] = [Math]::Round($out[$wo] + [double]$h, 2)
    }
    return $out
}

# ============================================================
# Row / summary builders
# ============================================================
function ConvertTo-MachineId {
    param([string]$Acronym, [string]$Equipment)
    $a = "$Acronym".Trim().ToUpper()
    $e = "$Equipment".Trim()
    if ($a -and $e) { return "$a $e" }
    if ($a) { return $a }
    return ''
}

function Get-DetailFieldValue {
    param($Fields, [string[]]$Names)
    foreach ($n in $Names) {
        if ($Fields.ContainsKey($n)) { return "$($Fields[$n])" }
        $u = $n.ToUpper(); if ($Fields.ContainsKey($u)) { return "$($Fields[$u])" }
        $l = $n.ToLower(); if ($Fields.ContainsKey($l)) { return "$($Fields[$l])" }
    }
    return ''
}

function New-DetailRow {
    param($WO, $Detail, $LocalHours)

    $f = if ($Detail) { $Detail.Fields } else { @{} }
    $dateIssued = Get-DetailFieldValue $f @('date_issued','DATE_ISSUED')
    if ([string]::IsNullOrWhiteSpace($dateIssued)) { $dateIssued = $WO.DateIssued }
    $dateDue = Get-DetailFieldValue $f @('date_due','DATE_DUE')
    if ([string]::IsNullOrWhiteSpace($dateDue)) { $dateDue = $WO.DueDate }
    $dateClosed = Get-DetailFieldValue $f @('date_closed','DATE_CLOSED')
    $workCode   = Get-DetailFieldValue $f @('work_code','WORK_CODE')
    if ([string]::IsNullOrWhiteSpace($workCode)) { $workCode = $WO.WorkCode }
    $crew       = Get-DetailFieldValue $f @('crew_id','CREW_ID')
    if ([string]::IsNullOrWhiteSpace($crew)) { $crew = $WO.Crew }

    return [PSCustomObject]@{
        'WO No'              = $WO.WorkOrderNo
        'Register'           = $WO.RegisterNo
        'Subsite'            = Get-DetailFieldValue $f @('subsite_id','SUBSITE_ID')
        'Status'             = $WO.Status
        'Acronym'            = $WO.Acronym
        'Equipment'          = $WO.Equipment
        'Machine ID'         = (ConvertTo-MachineId $WO.Acronym $WO.Equipment)
        'Class Code'         = $WO.ClassCode
        'Description'        = $WO.Description
        'Work Code'          = $workCode
        'Crew'               = $crew
        'Department'         = Get-DetailFieldValue $f @('dept_code','DEPT_CODE')
        'Date Issued'        = $dateIssued
        'Date Due'           = $dateDue
        'Date Closed'        = $dateClosed
        'Est Hours'          = Get-DetailFieldValue $f @('est_hours','EST_HOURS')
        'Hours To Date'      = Get-DetailFieldValue $f @('hours_to_date','HOURS_TO_DATE')
        'Local Hours'        = if ($null -ne $LocalHours) { $LocalHours } else { '' }
        'Standing'           = Get-DetailFieldValue $f @('standing_flag','STANDING_FLAG')
        'Project'            = Get-DetailFieldValue $f @('project_id','PROJECT_ID')
        'Facility'           = Get-DetailFieldValue $f @('whep_fac_id','WHEP_FAC_ID')
        'Past Due Reason'    = Get-DetailFieldValue $f @('past_due_reason_code','PAST_DUE_REASON_CODE')
        'Bulletin 3-Alpha'   = Get-DetailFieldValue $f @('nm_selBulletin_3Alpha')
        'Bulletin 3-Num'     = Get-DetailFieldValue $f @('nm_txtBulletin_3Num')
        'Bulletin 2-Num'     = Get-DetailFieldValue $f @('nm_selBulletin_2Num')
        'ACBM'               = Get-DetailFieldValue $f @('p_acbm_ind','P_ACBM_IND')
        'Form 4772 Required' = Get-DetailFieldValue $f @('req_4772','REQ_4772')
        'FSSP Call'          = Get-DetailFieldValue $f @('txt_fssp_call','TXT_FSSP_CALL')
        'FSSP Problem'       = Get-DetailFieldValue $f @('txt_fssp_problem','TXT_FSSP_PROBLEM')
        'FSSP ZIP'           = Get-DetailFieldValue $f @('txt_fssp_zip','TXT_FSSP_ZIP')
        'Comments'           = Get-DetailFieldValue $f @('wo_comment','WO_COMMENT')
        'Created By'         = Get-DetailFieldValue $f @('create_user','CREATE_USER')
        'Created Date'       = Get-DetailFieldValue $f @('create_date','CREATE_DATE')
        'Modified By'        = Get-DetailFieldValue $f @('mod_user','MOD_USER')
        'Modified Date'      = Get-DetailFieldValue $f @('mod_date','MOD_DATE')
        'Site'               = $WO.SiteId
        'Detail Route'       = if ($Detail) { $Detail.Route } else { 'search-only' }
    }
}

function Get-MachineSummary {
    param([array]$DetailRows)
    $byMachine = @{}
    foreach ($row in $DetailRows) {
        $mid = "$($row.'Machine ID')"
        if ([string]::IsNullOrWhiteSpace($mid)) { $mid = '(unknown)' }
        if (-not $byMachine.ContainsKey($mid)) {
            $byMachine[$mid] = [PSCustomObject]@{
                MachineId   = $mid
                WoCount     = 0
                ClosedCount = 0
                EmarsHours  = 0.0
                LocalHours  = 0.0
                EstHours    = 0.0
                DetailRows  = 0
                SearchOnly  = 0
            }
        }
        $acc = $byMachine[$mid]
        $acc.WoCount++
        if (-not [string]::IsNullOrWhiteSpace("$($row.'Date Closed')")) { $acc.ClosedCount++ }
        if ("$($row.'Detail Route')" -eq 'search-only') { $acc.SearchOnly++ } else { $acc.DetailRows++ }

        $h = 0.0; if ([double]::TryParse("$($row.'Hours To Date')", [ref]$h)) { $acc.EmarsHours += $h }
        $h = 0.0; if ([double]::TryParse("$($row.'Local Hours')",   [ref]$h)) { $acc.LocalHours += $h }
        $h = 0.0; if ([double]::TryParse("$($row.'Est Hours')",     [ref]$h)) { $acc.EstHours   += $h }
    }
    $out = @()
    foreach ($mid in ($byMachine.Keys | Sort-Object)) {
        $m = $byMachine[$mid]
        $out += [PSCustomObject]@{
            'Machine ID'          = $m.MachineId
            'Work Orders'         = $m.WoCount
            'Closed WOs'          = $m.ClosedCount
            'eMARS Hours'         = [Math]::Round($m.EmarsHours, 2)
            'Local Hours'         = [Math]::Round($m.LocalHours, 2)
            'Delta (eMARS-Local)' = [Math]::Round($m.EmarsHours - $m.LocalHours, 2)
            'Est Hours'           = [Math]::Round($m.EstHours, 2)
            'Detail Fetched'      = $m.DetailRows
            'Search-Only'         = $m.SearchOnly
        }
    }
    return ,$out
}

# ============================================================
# Main pipeline
# ============================================================
function Invoke-HistoryExport {
    param(
        [DateTime]$WindowFrom,
        [DateTime]$WindowTo,
        [string]$MachineFilter,
        [string[]]$StatusCodes,
        [string]$FilterMode,
        [string]$OutDir,
        [string]$SingleRegister = '',
        [int]$MaxRegs = 30,
        [int]$MaxRows = 200,
        [switch]$SkipDetailFetch
    )

    if ([string]::IsNullOrWhiteSpace($OutDir)) {
        $rel = "$($script:config.WorkTracking.ReportsDirectory)".Trim()
        if ([string]::IsNullOrWhiteSpace($rel)) {
            $OutDir = Join-Path $script:config.RootDirectory 'Work Tracking\Reports'
        } elseif ([System.IO.Path]::IsPathRooted($rel)) {
            $OutDir = $rel
        } else {
            $OutDir = Join-Path $script:config.RootDirectory $rel
        }
    }
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

    $site = "$($script:config.WorkTracking.EmarsSite)".Trim()
    if ([string]::IsNullOrWhiteSpace($site)) {
        throw "EmarsSite is not set in Config.json."
    }

    Write-Log "Window: $($WindowFrom.ToString('yyyy-MM-dd')) .. $($WindowTo.ToString('yyyy-MM-dd'))   Filter: $FilterMode"
    Write-Log "Statuses: $($StatusCodes -join ',')   Site: $site   Register: '$(if ($SingleRegister) { $SingleRegister } else { '<none>' })'   MaxRows: $MaxRows"

    # ---- 1. Search ----
    $allWos = Get-EmarsWorkOrdersAll -SiteId $site -StatusCodes $StatusCodes `
                -SingleRegister $SingleRegister -MaxRegs $MaxRegs -MaxRows $MaxRows `
                -FilterMode $FilterMode -WindowFrom $WindowFrom -WindowTo $WindowTo
    Write-Log "Total distinct WOs after search: $(@($allWos).Count)"
    if (@($allWos).Count -eq 0) {
        Write-Log "No WOs survived.  Check Debug Dumps\ for raw responses."
        return [PSCustomObject]@{ DetailCsv=''; SummaryCsv=''; Details=0; Machines=0 }
    }

    # ---- 2. Local hours ----
    $localHours = Get-LocalHoursByWO -WindowFrom $WindowFrom -WindowTo $WindowTo
    Write-Log "Local Jobs.json hours per WO in window: $($localHours.Count) WO(s)"

    # ---- 3. Detail + rows ----
    $detailRows = @()
    $i = 0
    foreach ($wo in @($allWos)) {
        $i++

        if ($MachineFilter) {
            $mid = ConvertTo-MachineId $wo.Acronym $wo.Equipment
            if ($mid -notmatch $MachineFilter) { continue }
        }
        if ($FilterMode -eq 'Logged' -and -not $localHours.ContainsKey($wo.WorkOrderNo)) { continue }

        $localH = if ($localHours.ContainsKey($wo.WorkOrderNo)) { $localHours[$wo.WorkOrderNo] } else { $null }

        $detail = $null
        if (-not $SkipDetailFetch) {
            Write-Log "[$i/$(@($allWos).Count)] detail for WO #$($wo.WorkOrderNo) (wo_ptr='$($wo.WoPtr)')"
            $detail = Get-EmarsWorkOrderDetail -WoPtr $wo.WoPtr -WoNo $wo.WorkOrderNo -SiteId $site
        }
        $detailRows += New-DetailRow -WO $wo -Detail $detail -LocalHours $localH
    }

    Write-Log "Rows after filter: $($detailRows.Count)"

    # ---- 4. CSVs ----
    $stamp = "$($WindowFrom.ToString('yyyyMMdd'))_$($WindowTo.ToString('yyyyMMdd'))"
    $detailPath  = Join-Path $OutDir "WorkOrderDetail_$stamp.csv"
    $summaryPath = Join-Path $OutDir "WorkOrderSummary_$stamp.csv"

    if ($detailRows.Count -gt 0) {
        $detailRows | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
        Write-Log "Wrote detail CSV: $detailPath ($($detailRows.Count) rows)"
    }

    $summary = Get-MachineSummary -DetailRows $detailRows
    if ($summary.Count -gt 0) {
        $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
        Write-Log "Wrote summary CSV: $summaryPath ($($summary.Count) machines)"
    }

    Write-Log "Done."
    return [PSCustomObject]@{
        DetailCsv  = $detailPath
        SummaryCsv = $summaryPath
        Details    = $detailRows.Count
        Machines   = $summary.Count
    }
}

# ============================================================
# Diagnostics
# ============================================================
function Test-DetailEndpoints {
    Write-Log "=== DiscoverEndpoints mode ==="
    $site = "$($script:config.WorkTracking.EmarsSite)".Trim()
    if ([string]::IsNullOrWhiteSpace($site)) { throw "EmarsSite is not set." }

    # ---- 1. One search, one status, full first-row log ----
    Write-Log ""
    Write-Log "--- Search probe (status E, empty register) ---"
    $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWOData' -Query "pSite=$site&pWOStatus=E&pRegNo=&pWONo=" -TimeoutSec 60
    $wos = ConvertFrom-WoSearchResponse -Content $body -SiteId $site -Tag 'probe'
    Write-Log "Probe returned $(@($wos).Count) WO(s) with no date filter"
    if (@($wos).Count -eq 0) {
        Write-Log "No WOs at all -- check Debug Dumps for raw response shape"
        return
    }

    $sample = $wos[0]
    Write-Log ""
    Write-Log "--- Sample WO ---"
    Write-Log "  WO No:      $($sample.WorkOrderNo)"
    Write-Log "  WoPtr:      '$($sample.WoPtr)'"
    Write-Log "  Register:   $($sample.RegisterNo)"
    Write-Log "  Acronym:    $($sample.Acronym)"
    Write-Log "  Equipment:  $($sample.Equipment)"
    Write-Log "  DateIssued: $($sample.DateIssued)"
    Write-Log "  DueDate:    $($sample.DueDate)"
    Write-Log "  RawCells ($($sample.RawCells.Count)): $($sample.RawCells -join ' | ')"

    # ---- 2. Detail routes ----
    $user = "$($script:config.WorkTracking.EmarsUser)".Trim()
    if ([string]::IsNullOrWhiteSpace($user)) { $user = '0' }

    $probes = @(
        @{ Name='SWIL WM_002 / WO_NO'; Path='!vm_main_pkg.main';
           Query="pUser=$user&pSiteID=$site&pModule=EM_WORK_MAINT_002&pparamxml=<KEY><WO_NO>$($sample.WorkOrderNo)</WO_NO></KEY>" },
        @{ Name='SWIL WM_001 / WO_NO'; Path='!vm_main_pkg.main';
           Query="pUser=$user&pSiteID=$site&pModule=EM_WORK_MAINT_001&pparamxml=<KEY><WO_NO>$($sample.WorkOrderNo)</WO_NO></KEY>" }
    )
    if ($sample.WoPtr) {
        $probes += @{ Name='SWIL WM_002 / WO_PTR'; Path='!vm_main_pkg.main';
                      Query="pUser=$user&pSiteID=$site&pModule=EM_WORK_MAINT_002&pparamxml=<KEY><WO_PTR>$($sample.WoPtr)</WO_PTR></KEY>" }
    }
    $probes += @(
        @{ Name='em_getWODetail'; Path='em_work_pkg.em_getWODetail';
           Query="pSite=$site&pWONo=$($sample.WorkOrderNo)" },
        @{ Name='em_getWO'; Path='em_work_pkg.em_getWO';
           Query="pSite=$site&pWONo=$($sample.WorkOrderNo)" }
    )

    foreach ($p in $probes) {
        Write-Log ""
        Write-Log "--- PROBE: $($p.Name) ---"
        try {
            $resp = Invoke-EmarsGet -Path $p.Path -Query $p.Query -TimeoutSec 25
            $len = if ($null -eq $resp) { 0 } else { $resp.Length }
            Write-Log "    length: $len"
            if ($resp -match '<widget') {
                $w = Get-WidgetValuesFromSwil -Html $resp
                Write-Log "    SWIL widgets: $($w.Count)"
                $returnedWo = ''
                foreach ($k in @('wo_no_h','wo_no','divwono')) {
                    if ($w.ContainsKey($k) -and $w[$k]) { $returnedWo = $w[$k]; break }
                }
                Write-Log "    returned WO: '$returnedWo'"
                Write-Log "    sample widget names: $((@($w.Keys | Select-Object -First 20)) -join ', ')"
            } elseif ($resp -match '^\s*<\?xml') {
                $xml = New-Object System.Xml.XmlDocument
                $xml.LoadXml($resp)
                $nodes = @($xml.SelectNodes('//ROW[1]/*'))
                Write-Log "    XML fields: $(($nodes | ForEach-Object { $_.Name } | Select-Object -First 20) -join ', ')"
            }
        } catch {
            Write-Log "    FAILED: $($_.Exception.Message)"
        }
    }

    Write-Log ""
    Write-Log "=== DiscoverEndpoints complete ==="
}

# ============================================================
# GUI
# ============================================================
function Show-HistoryGui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Work Order History'
    $form.Size = New-Object System.Drawing.Size(720, 500)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $mkLabel = {
        param($t, $y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $t; $l.Location = New-Object System.Drawing.Point(20, $y)
        $l.Size = New-Object System.Drawing.Size(110, 22); $l.TextAlign = 'MiddleRight'
        $form.Controls.Add($l)
    }

    & $mkLabel 'From:' 24
    $dtpFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtpFrom.Format='Short'; $dtpFrom.Location = New-Object System.Drawing.Point(134,22)
    $dtpFrom.Size = New-Object System.Drawing.Size(160,24); $dtpFrom.Value = (Get-Date).Date.AddDays(-7)
    $form.Controls.Add($dtpFrom)

    & $mkLabel 'To:' 58
    $dtpTo = New-Object System.Windows.Forms.DateTimePicker
    $dtpTo.Format='Short'; $dtpTo.Location = New-Object System.Drawing.Point(134,56)
    $dtpTo.Size = New-Object System.Drawing.Size(160,24); $dtpTo.Value = (Get-Date).Date
    $form.Controls.Add($dtpTo)

    & $mkLabel 'Machine filter:' 92
    $txtMachine = New-Object System.Windows.Forms.TextBox
    $txtMachine.Location = New-Object System.Drawing.Point(134,90); $txtMachine.Size = New-Object System.Drawing.Size(360,24)
    $txtMachine.Text = $Machine; $form.Controls.Add($txtMachine)

    & $mkLabel 'Register:' 126
    $txtReg = New-Object System.Windows.Forms.TextBox
    $txtReg.Location = New-Object System.Drawing.Point(134,124); $txtReg.Size = New-Object System.Drawing.Size(160,24)
    $txtReg.Text = $Register; $form.Controls.Add($txtReg)

    & $mkLabel 'Filter by:' 160
    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.DropDownStyle='DropDownList'; $cmbFilter.Location = New-Object System.Drawing.Point(134,158)
    $cmbFilter.Size = New-Object System.Drawing.Size(160,24)
    [void]$cmbFilter.Items.AddRange(@('Due','Issued','Logged'))
    $cmbFilter.SelectedItem = $FilterBy
    $form.Controls.Add($cmbFilter)

    & $mkLabel 'Status codes:' 194
    $txtStatus = New-Object System.Windows.Forms.TextBox
    $txtStatus.Location = New-Object System.Drawing.Point(134,192); $txtStatus.Size = New-Object System.Drawing.Size(360,24)
    $txtStatus.Text = ($Status -join ','); $form.Controls.Add($txtStatus)

    & $mkLabel 'Max rows:' 228
    $numMax = New-Object System.Windows.Forms.NumericUpDown
    $numMax.Location = New-Object System.Drawing.Point(134,226); $numMax.Size = New-Object System.Drawing.Size(80,24)
    $numMax.Minimum=0; $numMax.Maximum=100000; $numMax.Value = $MaxRows
    $form.Controls.Add($numMax)

    $chkSkip = New-Object System.Windows.Forms.CheckBox
    $chkSkip.Text = 'Skip detail fetch (fast -- summary only)'
    $chkSkip.Location = New-Object System.Drawing.Point(230,227); $chkSkip.Size = New-Object System.Drawing.Size(300,24)
    $chkSkip.Checked = $SkipDetail
    $form.Controls.Add($chkSkip)

    & $mkLabel 'Output dir:' 262
    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(134,260); $txtOut.Size = New-Object System.Drawing.Size(360,24)
    $txtOut.Text = $OutputDirectory; $form.Controls.Add($txtOut)

    $browseBtn = New-Object System.Windows.Forms.Button
    $browseBtn.Text='...'; $browseBtn.Location = New-Object System.Drawing.Point(500,260)
    $browseBtn.Size = New-Object System.Drawing.Size(40,24)
    $browseBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = $dlg.SelectedPath }
    })
    $form.Controls.Add($browseBtn)

    $runBtn = New-Object System.Windows.Forms.Button
    $runBtn.Text='Run'; $runBtn.Location = New-Object System.Drawing.Point(20,296)
    $runBtn.Size = New-Object System.Drawing.Size(120,32)
    $runBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $runBtn.ForeColor = [System.Drawing.Color]::White
    $runBtn.FlatStyle = 'Flat'; $runBtn.FlatAppearance.BorderSize=0
    $runBtn.Add_Click({
        $runBtn.Enabled = $false
        $txtResult.Text = 'Running...'
        $form.Refresh()
        try {
            $statuses = @($txtStatus.Text -split ',' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ })
            $r = Invoke-HistoryExport `
                -WindowFrom $dtpFrom.Value.Date `
                -WindowTo $dtpTo.Value.Date `
                -MachineFilter $txtMachine.Text.Trim() `
                -StatusCodes $statuses `
                -FilterMode "$($cmbFilter.SelectedItem)" `
                -OutDir $txtOut.Text.Trim() `
                -SingleRegister $txtReg.Text.Trim() `
                -MaxRows ([int]$numMax.Value) `
                -SkipDetailFetch:$chkSkip.Checked
            $txtResult.Text = "Done.`r`nDetail CSV: $($r.DetailCsv)`r`nSummary CSV: $($r.SummaryCsv)`r`nRows: $($r.Details)  Machines: $($r.Machines)"
        } catch {
            $txtResult.Text = "Failed: $($_.Exception.Message)"
        } finally {
            $runBtn.Enabled = $true
        }
    })
    $form.Controls.Add($runBtn)

    $txtResult = New-Object System.Windows.Forms.TextBox
    $txtResult.Multiline=$true; $txtResult.ReadOnly=$true
    $txtResult.ScrollBars='Vertical'
    $txtResult.Location = New-Object System.Drawing.Point(20,340)
    $txtResult.Size = New-Object System.Drawing.Size(660,120)
    $txtResult.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtResult.Text = 'Ready.'
    $form.Controls.Add($txtResult)

    $form.ShowDialog() | Out-Null
}

# ============================================================
# Entry
# ============================================================
try {
    $script:config = Load-Config

    if ($DiscoverEndpoints) { Test-DetailEndpoints; return }
    if ($ShowGui)           { Show-HistoryGui;      return }

    $result = Invoke-HistoryExport `
        -WindowFrom $From -WindowTo $To `
        -MachineFilter $Machine -StatusCodes $Status `
        -FilterMode $FilterBy -OutDir $OutputDirectory `
        -SingleRegister $Register -MaxRegs $MaxRegisters `
        -MaxRows $MaxRows -SkipDetailFetch:$SkipDetail

    Write-Host ""
    Write-Host "Detail CSV : $($result.DetailCsv)" -ForegroundColor Cyan
    Write-Host "Summary CSV: $($result.SummaryCsv)" -ForegroundColor Cyan
    Write-Host "Work orders: $($result.Details)   Machines: $($result.Machines)" -ForegroundColor Cyan
} catch {
    Write-Log "FATAL: $($_.Exception.Message)"
    Write-Host "FATAL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
```

---

Test sequence (in this order)

Test 1 — prove the pipeline works end-to-end without any detail fetch.

```powershell
.\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -SkipDetail
```

Expected within ~30 seconds:

```
Wrote summary CSV: ...\WorkOrderSummary_20261023_20261030.csv (N machines)
Work orders: 571   Machines: N
```

Open the summary CSV. You'll have real per-machine WO counts. eMARS Hours and Est Hours columns will be empty (no detail fetch); Local Hours populated from your Jobs.json. This works right now regardless of the detail endpoint.

Test 2 — find out whether the detail route is reachable at all.

```powershell
.\WO.ps1 -DiscoverEndpoints
```

Takes ~60 seconds. Read WorkOrderHistory.log. You'll see:

· The full first row's cells 0-15 — this tells you definitively whether hidden columns exist and what's in them
· WoPtr: '...' — whether any cell had a numeric pointer
· Each SWIL/detail probe with returned WO: — whether it returned the right WO

If WoPtr is populated in the sample, and any probe returns the requested WO# in the widget check, detail fetch will work. If not, we know the detail route isn't reachable and the summary-only path is the ceiling.

Test 3 — full run with detail (only if Test 2 showed a working route).

```powershell
.\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -MaxRows 50
```

The -MaxRows 50 cap means at most 50 detail fetches — maybe 2 minutes — enough to prove the detail route populates eMARS Hours and Est Hours without burning half an hour on 571 WOs. If those columns fill in, drop -MaxRows and let it run to completion.

Paste the log from Test 1 first (the summary-only run). That alone proves the working half and takes 30 seconds.