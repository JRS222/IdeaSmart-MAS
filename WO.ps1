<#
.SYNOPSIS
    Standalone eMARS Work Order history extractor.

.DESCRIPTION
    Pulls corrective-maintenance work orders from eMARS for a date window,
    aggregates them per machine, and emits two CSVs:

      WorkOrderSummary_<from>_<to>.csv   one row per machine
      WorkOrderDetail_<from>_<to>.csv    one row per work order

    Search  : em_work_pkg.em_getWOData    (applet <param> grid HTML)
    Detail  : em_work_pkg.em_getWO        (XML ROWSET keyed on p_rowid)
    Registers: em_work_pkg.em_getXMLRegisterListAll

    The detail route was reverse-engineered from the EM_WORK_MAINT_001
    form's own JavaScript: the search grid's click handler reads column 15
    (an Oracle ROWID) and the form loads the record with
    em_getWO?p_rowid=<ROWID>&pSiteID=<site>.  Earlier attempts to load
    detail through the !vm_main_pkg.main "SWIL" shell do not work and are
    intentionally absent from this script.

    Diagnostics are written to WorkOrderHistory.log next to this script;
    raw responses that fail to parse are saved to Debug Dumps\.

.PARAMETER From / To
    Date window (inclusive).  Defaults to the last 7 days.

.PARAMETER Machine
    Optional substring/regex filter on the machine ID (e.g. 'DBCS 49').

.PARAMETER Status
    WO status code(s) to search.  Default E.

.PARAMETER Register
    Optional single register number.  Blank = all registers.

.PARAMETER WONo
    Optional exact work-order number (added to the search query).

.PARAMETER SkipDetail
    Skip the per-WO detail fetch (fast).  Summary will still carry
    machine-level work-order counts; eMARS hours/estimates stay empty.

.PARAMETER DiscoverEndpoints
    Diagnostic mode: runs the search probe, dumps the first row's cells,
    then tests em_getWO with the real pointer.  Writes no CSVs.

.EXAMPLE
    .\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -SkipDetail

.EXAMPLE
    .\WO.ps1 -From 2026-10-23 -To 2026-10-30 -Status E -Machine 'DBCS 49' -MaxRows 0

.EXAMPLE
    .\WO.ps1 -DiscoverEndpoints
#>

[CmdletBinding()]
param(
    [DateTime]$From   = (Get-Date).Date.AddDays(-7),
    [DateTime]$To     = (Get-Date).Date,
    [string]$Machine  = '',
    [string[]]$Status = @('E'),
    [ValidateSet('Due','Issued','Logged')]
    [string]$FilterBy = 'Due',
    [string]$OutputDirectory = '',
    [string]$Register = '',
    [string]$WONo     = '',
    [string]$Site     = '',
    [string]$Config   = '',
    [int]$MaxRegisters = 30,
    [int]$MaxRows = 0,
    [switch]$SkipDetail,
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
    if (-not [string]::IsNullOrWhiteSpace($Config)) {
        if (Test-Path $Config) { return (Resolve-Path $Config).Path }
        Write-Log "Config file not found: $Config"
        return $null
    }
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
    if (-not $path) { throw "Config.json not found. Point -Config at the app's Config.json." }
    Write-Log "Loading config from $path"
    return (Get-Content -Path $path -Raw | ConvertFrom-Json)
}

function Get-CfgSetting {
    param([string]$Name, [string]$Default = '')
    $cfg = $script:config
    if ($null -eq $cfg) { return $Default }
    $wt = $cfg.WorkTracking
    if ($null -eq $wt) { return $Default }
    $p = $wt.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    $v = "$($p.Value)".Trim()
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}

# ============================================================
# HTTP
# ============================================================
function Get-EmarsBase {
    $origin = Get-CfgSetting 'EmarsOrigin' 'http://emarssu3.eng.usps.gov'
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
# WO search response parser (applet <param> grid)
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
        return @()
    }
    $Content = $Content.Trim()
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: $($Content.Length) bytes"

    # ---------- XML (some endpoints answer with a ROWSET) ----------
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
                WoPtr=$fields['WO_PTR']; WorkOrderNo=$fields['WONO']
                RegisterNo=$fields['REGNO']; WorkCode=$fields['WORK_CODE']
                Acronym=$fields['ACRONYM']; ClassCode=$fields['CLASSCODE']
                Equipment=$fields['EQUIP_NO']; Status=$fields['WO_STATUS']
                Crew=$fields['CREW_ID']; DateIssued=$fields['DATE_ISSUED']
                DueDate=$fields['DATE_DUE']; Description=$fields['WO_DESC']
                SiteId=$SiteId; RawCells=@()
            }
        }
        return $out
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # ---------- APPLET <param> format ----------
    $paramMatches = [regex]::Matches(
        $Content,
        '(?is)<param\b[^>]*?\bname\s*=\s*["''](?<name>[^"'']+)["''][^>]*?\bvalue\s*=\s*["''](?<val>[^"'']*)["'']'
    )
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: $($paramMatches.Count) <param> in $([int]$sw.Elapsed.TotalMilliseconds)ms"

    if ($paramMatches.Count -eq 0) {
        Save-RawResponse -Content $Content -Tag "WoSearch_$Tag" | Out-Null
        return @()
    }

    $rowParams = @{}
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

    $appletNumCols = 0
    if ($keyParams.ContainsKey('numcols')) {
        [void][int]::TryParse("$($keyParams['numcols'])", [ref]$appletNumCols)
    }
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: headers ($($headers.Count)) = $($headers -join ' | ')   numcols=$appletNumCols"

    if ($rowParams.Count -eq 0) { return @() }

    $headerCount = $headers.Count
    $sortedKeys  = @($rowParams.Keys | Sort-Object { [int]$_ })

    # ---- Delimiter detection ----
    $separator = $null
    foreach ($sk in (@($sortedKeys | Select-Object -First 10))) {
        $rawTest = "$($rowParams[$sk])"
        foreach ($sep in @(',', '|', ';', '~', "`t")) {
            $test = @($rawTest -split [regex]::Escape($sep))
            if ($test.Count -ge $headerCount) { $separator = $sep; break }
        }
        if ($separator) { break }
    }
    if (-not $separator) { $separator = ',' }
    Write-Log "ConvertFrom-WoSearchResponse[$Tag]: delimiter = '$separator'"

    # ---- Log first row cells 0-N ----
    $firstKey = $sortedKeys[0]
    $firstRaw = "$($rowParams[$firstKey])"
    $firstCells = @($firstRaw -split [regex]::Escape($separator))
    Write-Log "  --- FIRST ROW ($($firstCells.Count) cells) ---"
    for ($pi = 0; $pi -lt $firstCells.Count; $pi++) {
        $label = if ($pi -lt $headerCount) { $headers[$pi] } else { "(hidden_$pi)" }
        $val = "$($firstCells[$pi])"
        if ($val.Length -gt 60) { $val = $val.Substring(0,60) + '...' }
        $val = $val -replace "`r|`n", ' '
        Write-Log "    [$pi] $label = '$val'"
    }

    # ---- Column index map ----
    $idx = @{}
    for ($i = 0; $i -lt $headerCount; $i++) { $idx[$headers[$i]] = $i }
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

    # ---- Parse loop ----
    $out = @()
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

        # WoPtr lives in cell 14 (0-indexed): an 18-char Oracle ROWID
        # (alphanumeric with +/=, base64-ish).  Fall back to scanning the
        # hidden tail cells (11..N) for the same shape.
        $woPtr = ''
        if ($cells.Count -ge 15) {
            $c14 = "$($cells[14])".Trim()
            if ($c14 -match '^[A-Za-z0-9+/=]{10,}$') { $woPtr = $c14 }
        }
        if (-not $woPtr -and $cells.Count -gt 11) {
            for ($i = $cells.Count - 1; $i -ge 11; $i--) {
                $cv = "$($cells[$i])".Trim()
                if ($cv -match '^[A-Za-z0-9+/=]{15,}$') { $woPtr = $cv; break }
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
    return $out
}

# ============================================================
# Register list + enumeration
# ============================================================
function Get-EmarsRegisterList {
    param([string]$SiteId)
    try {
        # em_getXMLRegisterListAll is what the form itself uses in go1().
        $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getXMLRegisterListAll' -Query "pSiteID=$SiteId"
        if ($body -notmatch '^\s*<\?xml' -and $body -notmatch '^\s*<\s*ROWSET') {
            Write-Log "Get-EmarsRegisterList: non-XML ($($body.Length) bytes)"
            return @()
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
        return $regs
    } catch {
        Write-Log "Get-EmarsRegisterList: $($_.Exception.Message)"
        return @()
    }
}

function Get-EmarsWorkOrdersAll {
    param(
        [string]$SiteId,
        [string[]]$StatusCodes,
        [string]$SingleRegister = '',
        [string]$SingleWONo = '',
        [int]$MaxRegs = 30,
        [int]$MaxRows = 0,
        [string]$FilterMode = '',
        [DateTime]$WindowFrom = [DateTime]::MinValue,
        [DateTime]$WindowTo   = [DateTime]::MinValue
    )

    $seen = @{}
    $out  = @()

    $doSearch = {
        param([string]$reg)
        foreach ($st in $StatusCodes) {
            $q = "pSite=$SiteId&pWOStatus=$st&pRegNo=$reg&pWONo=$SingleWONo"
            try {
                $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWOData' -Query $q -TimeoutSec 60
            } catch {
                Write-Log "  reg='$reg' status='$st' FAILED -- $($_.Exception.Message)"
                continue
            }
            $tag = if ($reg) { "reg${reg}_${st}" } else { "empty_${st}" }
            $wos = ConvertFrom-WoSearchResponse -Content $body -SiteId $SiteId -Tag $tag `
                    -FilterMode $FilterMode -WindowFrom $WindowFrom -WindowTo $WindowTo
            Write-Log "  reg='$reg' status='$st' -> $(@($wos).Count) WO(s)"
            foreach ($wo in @($wos)) {
                if ([string]::IsNullOrWhiteSpace($wo.WorkOrderNo)) { continue }
                if ($seen.ContainsKey($wo.WorkOrderNo)) { continue }
                $seen[$wo.WorkOrderNo] = $true
                $script:__wos_out += $wo
                if ($MaxRows -gt 0 -and $script:__wos_out.Count -ge $MaxRows) {
                    Write-Log "  MaxRows ($MaxRows) reached -- stopping"
                    return $true
                }
            }
        }
        return $false
    }

    $script:__wos_out = $out

    if ([string]::IsNullOrWhiteSpace($SingleRegister) -or -not [string]::IsNullOrWhiteSpace($SingleWONo)) {
        if (& $doSearch $SingleRegister) { return $script:__wos_out }
        if ($script:__wos_out.Count -gt 0) { return $script:__wos_out }
        Write-Log "Get-EmarsWorkOrdersAll: empty-register path returned 0 -- trying registers"
    }

    $regs = @()
    if (-not [string]::IsNullOrWhiteSpace($SingleRegister)) {
        $regs = @($SingleRegister)
    } else {
        $regs = Get-EmarsRegisterList -SiteId $SiteId
        if ($regs.Count -eq 0) {
            Write-Log "Get-EmarsWorkOrdersAll: no register list; skipping register sweep (target a register with -Register)"
            return $script:__wos_out
        }
        if ($MaxRegs -gt 0 -and $regs.Count -gt $MaxRegs) {
            $regs = @($regs | Select-Object -First $MaxRegs)
        }
    }

    foreach ($reg in $regs) {
        if (& $doSearch $reg) { return $script:__wos_out }
    }
    return $script:__wos_out
}

# ============================================================
# Detail: em_getWO (XML ROWSET keyed on the grid's ROWID pointer)
# ============================================================
function Get-EmarsWorkOrderDetail {
    param(
        [string]$WoPtr = '',
        [Parameter(Mandatory)][string]$WoNo,
        [Parameter(Mandatory)][string]$SiteId
    )

    if ([string]::IsNullOrWhiteSpace($WoPtr)) {
        Write-Log "Detail: WO $WoNo has no WoPtr; cannot fetch"
        return $null
    }

    # The form's own go1(rowid) does:
    #   XMLDOM.load("em_work_pkg.em_getWO?p_rowid="+rowid+"&pSiteID="+currentSite)
    # and every element child of the ROW is a form field.
    # ROWIDs can contain '+' and '/', so they must be URL-encoded.
    $rowid = [uri]::EscapeDataString("$WoPtr".Trim())
    $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWO' -Query "p_rowid=$rowid&pSiteID=$SiteId" -TimeoutSec 30
    $body = "$body".Trim()

    if ($body -notmatch '^\s*<\?xml' -and $body -notmatch '^\s*<\s*ROWSET') {
        Write-Log "Detail: em_getWO for WO $WoNo returned non-XML ($($body.Length) bytes)"
        Save-RawResponse -Content $body -Tag "DetailWO_$WoNo" | Out-Null
        return $null
    }

    $xml = New-Object System.Xml.XmlDocument
    try { $xml.LoadXml($body) } catch {
        Save-RawResponse -Content $body -Tag "DetailWO_$WoNo" | Out-Null
        Write-Log "Detail: em_getWO for WO $WoNo returned unparsable XML -- $($_.Exception.Message)"
        return $null
    }

    $rows = @($xml.SelectNodes('//ROW'))
    if ($rows.Count -eq 0) {
        Write-Log "Detail: em_getWO for WO $WoNo returned an empty ROWSET"
        return $null
    }

    $fields = @{}
    foreach ($child in $rows[0].ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $fields[$child.Name.ToLower()] = "$($child.InnerText)".Trim()
        }
    }

    Write-Log "Detail OK: em_getWO WO=$WoNo ptr=$WoPtr -> $($fields.Count) field(s)"
    return [PSCustomObject]@{ Route = 'em_getWO'; Fields = $fields }
}

function Get-EmarsWoPtrByNumber {
    # The form's own WO-number -> ROWID bridge:
    #   em_work_pkg.em_wono9?pSiteID=<site>&p_wo=<wono>&p_regno=<reg> -> <WONO9>rowid</WONO9>
    param(
        [Parameter(Mandatory)][string]$WoNo,
        [Parameter(Mandatory)][string]$SiteId,
        [string]$RegNo = ''
    )
    $body = Invoke-EmarsGet -Path 'em_work_pkg.em_wono9' -Query "pSiteID=$SiteId&p_wo=$WoNo&p_regno=$RegNo" -TimeoutSec 30
    $body = "$body".Trim()
    if ($body -notmatch '^\s*<\?xml') {
        Write-Log "Get-EmarsWoPtrByNumber: non-XML ($($body.Length) bytes) for WO $WoNo"
        Save-RawResponse -Content $body -Tag "Wono9_$WoNo" | Out-Null
        return ''
    }
    $xml = New-Object System.Xml.XmlDocument
    try { $xml.LoadXml($body) } catch { return '' }
    $node = $xml.SelectSingleNode('//WONO9')
    if ($null -eq $node) { return '' }
    return "$($node.InnerText)".Trim()
}

# ============================================================
# Local jobs
# ============================================================
function Get-LocalJobs {
    $rel = Get-CfgSetting 'JobsFile'
    if ([string]::IsNullOrWhiteSpace($rel)) { return @() }
    $path = if ([System.IO.Path]::IsPathRooted($rel)) { $rel } else {
        Join-Path "$($script:config.RootDirectory)" $rel
    }
    if (-not (Test-Path $path)) { return @() }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return @() }
        return @($data)
    } catch {
        Write-Log "Get-LocalJobs: $($_.Exception.Message)"
        return @()
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
    $workCode = Get-DetailFieldValue $f @('work_code','WORK_CODE')
    if ([string]::IsNullOrWhiteSpace($workCode)) { $workCode = $WO.WorkCode }
    $crew = Get-DetailFieldValue $f @('crew_id','CREW_ID')
    if ([string]::IsNullOrWhiteSpace($crew)) { $crew = $WO.Crew }

    return [PSCustomObject]@{
        'WO No'              = $WO.WorkOrderNo
        'Register'           = Get-DetailFieldValue $f @('reg_no','REG_NO')
        'Subsite'            = Get-DetailFieldValue $f @('subsite_id','SUBSITE_ID')
        'Status'             = Get-DetailFieldValue $f @('wo_status','WO_STATUS')
        'Acronym'            = Get-DetailFieldValue $f @('acronym','ACRONYM')
        'Equipment'          = Get-DetailFieldValue $f @('equip_no','EQUIP_NO')
        'Machine ID'         = (ConvertTo-MachineId $WO.Acronym $WO.Equipment)
        'Class Code'         = Get-DetailFieldValue $f @('class_code','CLASS_CODE')
        'Description'        = Get-DetailFieldValue $f @('wo_desc','WO_DESC')
        'Work Code'          = $workCode
        'Crew'               = $crew
        'Department'         = Get-DetailFieldValue $f @('dept_code','DEPT_CODE')
        'Date Issued'        = $dateIssued
        'Date Due'           = $dateDue
        'Date Closed'        = Get-DetailFieldValue $f @('date_closed','DATE_CLOSED')
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
                MachineId=$mid; WoCount=0; ClosedCount=0
                EmarsHours=0.0; LocalHours=0.0; EstHours=0.0
                DetailRows=0; SearchOnly=0
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
    return $out
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
        [string]$SingleWONo = '',
        [int]$MaxRegs = 30,
        [int]$MaxRows = 0,
        [switch]$SkipDetailFetch
    )

    if ([string]::IsNullOrWhiteSpace($OutDir)) {
        $rel = Get-CfgSetting 'ReportsDirectory'
        if ([string]::IsNullOrWhiteSpace($rel)) {
            $OutDir = Join-Path "$($script:config.RootDirectory)" 'Work Tracking\Reports'
        } elseif ([System.IO.Path]::IsPathRooted($rel)) {
            $OutDir = $rel
        } else {
            $OutDir = Join-Path "$($script:config.RootDirectory)" $rel
        }
    }
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

    $site = if (-not [string]::IsNullOrWhiteSpace($Site)) { $Site } else { Get-CfgSetting 'EmarsSite' '345' }
    Write-Log "Site: $site"

    Write-Log "Window: $($WindowFrom.ToString('yyyy-MM-dd')) .. $($WindowTo.ToString('yyyy-MM-dd'))   Filter: $FilterMode"
    Write-Log "Statuses: $($StatusCodes -join ',')   Register: '$(if ($SingleRegister) { $SingleRegister } else { '<none>' })'   WONo: '$(if ($SingleWONo) { $SingleWONo } else { '<none>' })'   MaxRows: $MaxRows   SkipDetail: $SkipDetailFetch"

    $allWos = Get-EmarsWorkOrdersAll -SiteId $site -StatusCodes $StatusCodes `
                -SingleRegister $SingleRegister -SingleWONo $SingleWONo `
                -MaxRegs $MaxRegs -MaxRows $MaxRows `
                -FilterMode $FilterMode -WindowFrom $WindowFrom -WindowTo $WindowTo
    Write-Log "Total distinct WOs after search: $(@($allWos).Count)"
    if (@($allWos).Count -eq 0) {
        Write-Log "No WOs survived.  Check Debug Dumps\."
        return [PSCustomObject]@{ DetailCsv=''; SummaryCsv=''; Details=0; Machines=0 }
    }

    $localHours = Get-LocalHoursByWO -WindowFrom $WindowFrom -WindowTo $WindowTo
    Write-Log "Local Jobs.json hours per WO in window: $($localHours.Count) WO(s)"

    # ---- Detail + rows ----
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
            Write-Log "[$i/$(@($allWos).Count)] detail for WO #$($wo.WorkOrderNo) (ptr=$($wo.WoPtr))"
            $detail = Get-EmarsWorkOrderDetail -WoPtr $wo.WoPtr -WoNo $wo.WorkOrderNo -SiteId $site

            # 'ALL'-equipment WOs carry their estimate behind em_esthours_prc
            # (<ESTHOUR>text</ESTHOUR>) instead of a plain est_hours field.
            if ($detail -and "$($detail.Fields['equip_no'])" -ieq 'ALL') {
                try {
                    $estBody = Invoke-EmarsGet -Path 'em_work_pkg.em_esthours_prc' -Query "pSiteID=$site&pWoPtr=$([uri]::EscapeDataString("$($wo.WoPtr)".Trim()))" -TimeoutSec 20
                    if ("$estBody".Trim() -match '^\s*<\?xml') {
                        $ex = New-Object System.Xml.XmlDocument
                        $ex.LoadXml("$estBody".Trim())
                        $n = $ex.SelectSingleNode('//ESTHOUR')
                        if ($null -ne $n) { $detail.Fields['est_hours'] = "$($n.InnerText)".Trim() }
                    }
                } catch {
                    Write-Log "  est-hours lookup failed for WO $($wo.WorkOrderNo): $($_.Exception.Message)"
                }
            }
        }
        $detailRows += New-DetailRow -WO $wo -Detail $detail -LocalHours $localH
    }

    Write-Log "Rows after filter: $($detailRows.Count)"

    # ---- CSVs ----
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
    Write-Log "=== DiscoverEndpoints ==="
    $site = if (-not [string]::IsNullOrWhiteSpace($Site)) { $Site } else { Get-CfgSetting 'EmarsSite' '345' }
    Write-Log "Site: $site"

    # Only the last 7 days of Due so the parse loop terminates in seconds
    $winFrom = (Get-Date).Date.AddDays(-7)
    $winTo   = (Get-Date).Date

    Write-Log ""
    Write-Log "--- Search probe (status E, due $($winFrom.ToString('yyyy-MM-dd'))..$($winTo.ToString('yyyy-MM-dd'))) ---"
    $body = Invoke-EmarsGet -Path 'em_work_pkg.em_getWOData' -Query "pSite=$site&pWOStatus=E&pRegNo=&pWONo=" -TimeoutSec 60
    $wos = ConvertFrom-WoSearchResponse -Content $body -SiteId $site -Tag 'probe' `
            -FilterMode 'Due' -WindowFrom $winFrom -WindowTo $winTo
    Write-Log "Probe returned $(@($wos).Count) WO(s)"
    if (@($wos).Count -eq 0) { return }

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

    # ---- em_getWO with the pointer (the form's own go1 route) ----
    Write-Log ""
    Write-Log "--- em_getWO with p_rowid=$($sample.WoPtr) ---"
    try {
        $resp = Invoke-EmarsGet -Path 'em_work_pkg.em_getWO' -Query "p_rowid=$($sample.WoPtr)&pSiteID=$site" -TimeoutSec 30
        $trimmed = "$resp".Trim()
        Write-Log "  response: $($trimmed.Length) bytes, starts: '$($trimmed.Substring(0, [Math]::Min(80, $trimmed.Length)))'"
        if ($trimmed -match '^\s*<\?xml' -or $trimmed -match '^\s*<\s*ROWSET') {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($trimmed)
            $rows = @($xml.SelectNodes('//ROW'))
            Write-Log "  XML rows: $($rows.Count)"
            if ($rows.Count -gt 0) {
                $names = @()
                foreach ($child in $rows[0].ChildNodes) {
                    if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                        $names += $child.Name.ToLower()
                    }
                }
                Write-Log "  field names: $($names -join ', ')"
            }
        } else {
            $dump = Save-RawResponse -Content $resp -Tag "DetailProbe_$($sample.WorkOrderNo)"
            Write-Log "  non-XML response saved to: $dump"
        }
    } catch {
        Write-Log "  FAILED: $($_.Exception.Message)"
    }

    Write-Log ""
    Write-Log "=== DiscoverEndpoints complete ==="
}

# ============================================================
# Entry
# ============================================================
try {
    $script:config = Load-Config

    if ($DiscoverEndpoints) { Test-DetailEndpoints; return }

    $result = Invoke-HistoryExport `
        -WindowFrom $From -WindowTo $To `
        -MachineFilter $Machine -StatusCodes $Status `
        -FilterMode $FilterBy -OutDir $OutputDirectory `
        -SingleRegister $Register -SingleWONo $WONo `
        -MaxRegs $MaxRegisters `
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
