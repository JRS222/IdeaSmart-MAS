################################################################################
#                                                                              #
#                     Parts Management System UI                               #
#                                                                              #
################################################################################

################################################################################
#                          Required .NET Assemblies                            #
################################################################################

# Load required assemblies for the Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Modern animated progress bar
Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class ModernProgressBar : Control
{
    private int _value = 0;
    private int _maximum = 100;
    private int _minimum = 0;
    private bool _marquee = false;
    private int _marqueeOffset = 0;
    private Timer _marqueeTimer;
    private float _glowPhase = 0f;
    private Timer _glowTimer;

    public Color BarColorStart = Color.FromArgb(52, 152, 219);   // blue
    public Color BarColorEnd   = Color.FromArgb(155, 89, 182);   // purple
    public Color TrackColor    = Color.FromArgb(226, 232, 240);
    public bool  ShowPercentage = true;

    public int Minimum { get { return _minimum; } set { _minimum = value; Invalidate(); } }
    public int Maximum { get { return _maximum; } set { _maximum = value; Invalidate(); } }
    public int Value {
        get { return _value; }
        set { _value = Math.Max(_minimum, Math.Min(_maximum, value)); Invalidate(); }
    }

    // Drop-in compatibility with System.Windows.Forms.ProgressBar
    public ProgressBarStyle Style {
        get { return _marquee ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks; }
        set {
            if (value == ProgressBarStyle.Marquee) { Marquee = true;  ShowPercentage = false; }
            else                                   { Marquee = false; ShowPercentage = true;  }
        }
    }

    public bool Marquee {
        get { return _marquee; }
        set {
            _marquee = value;
            if (_marquee) _marqueeTimer.Start();
            else { _marqueeTimer.Stop(); Invalidate(); }
        }
    }

    public ModernProgressBar()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint
               | ControlStyles.OptimizedDoubleBuffer
               | ControlStyles.UserPaint
               | ControlStyles.ResizeRedraw
               | ControlStyles.SupportsTransparentBackColor, true);
        Height = 26;
        BackColor = Color.Transparent;

        _marqueeTimer = new Timer { Interval = 25 };
        _marqueeTimer.Tick += (s, e) => {
            _marqueeOffset = (_marqueeOffset + 6) % (Width + 140);
            Invalidate();
        };

        _glowTimer = new Timer { Interval = 35 };
        _glowTimer.Tick += (s, e) => {
            _glowPhase += 0.09f;
            if (_glowPhase > (float)(Math.PI * 2)) _glowPhase -= (float)(Math.PI * 2);
            if (!_marquee) Invalidate();
        };
        _glowTimer.Start();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        int radius = Height / 2;
        var trackRect = new Rectangle(0, 0, Width - 1, Height - 1);

        using (var path  = RoundedRect(trackRect, radius))
        using (var brush = new SolidBrush(TrackColor))
            g.FillPath(brush, path);

        if (!_marquee)
        {
            double pct = (_maximum > _minimum)
                ? (double)(_value - _minimum) / (_maximum - _minimum) : 0;
            int fillWidth = (int)((Width - 2) * pct);

            if (fillWidth > 0)
            {
                var fillRect   = new Rectangle(1, 1, fillWidth, Height - 2);
                int fillRadius = Math.Min(radius, fillWidth / 2);

                using (var path = RoundedRect(fillRect, fillRadius))
                using (var brush = new LinearGradientBrush(
                    new Rectangle(0, 0, Math.Max(1, fillWidth), Height),
                    BarColorStart, BarColorEnd, LinearGradientMode.Horizontal))
                    g.FillPath(brush, path);

                // Subtle top shine
                var shineRect = new Rectangle(1, 1, fillWidth, Math.Max(1, Height / 2 - 1));
                using (var path = RoundedRect(fillRect, fillRadius))
                using (var shine = new LinearGradientBrush(
                    shineRect,
                    Color.FromArgb(95, 255, 255, 255),
                    Color.FromArgb(0, 255, 255, 255),
                    LinearGradientMode.Vertical))
                {
                    var state = g.Save();
                    g.SetClip(path);
                    g.FillRectangle(shine, shineRect);
                    g.Restore(state);
                }

                // Pulsing leading-edge glow
                float alpha = 0.30f + 0.30f * (float)Math.Sin(_glowPhase);
                int glowWidth = Math.Min(32, fillWidth);
                int glowX     = fillWidth - glowWidth + 1;
                if (glowX >= 0)
                {
                    var glowRect = new Rectangle(glowX, 1, glowWidth, Height - 2);
                    using (var path = RoundedRect(glowRect, Math.Min(radius, glowWidth / 2)))
                    using (var glowBrush = new LinearGradientBrush(
                        glowRect,
                        Color.FromArgb((int)(alpha * 255), 255, 255, 255),
                        Color.FromArgb(0, 255, 255, 255),
                        LinearGradientMode.Horizontal))
                        g.FillPath(glowBrush, path);
                }
            }
        }
        else
        {
            // Marquee: sliding gradient block
            int blockWidth = Math.Max(90, Width / 3);
            int x = _marqueeOffset - blockWidth;
            var clipRect  = new Rectangle(x, 1, blockWidth, Height - 2);
            var intersect = Rectangle.Intersect(clipRect, new Rectangle(1, 1, Width - 2, Height - 2));
            if (intersect.Width > 0)
            {
                using (var path = RoundedRect(new Rectangle(1, 1, Width - 2, Height - 2), radius))
                using (var brush = new LinearGradientBrush(
                    clipRect, BarColorStart, BarColorEnd, LinearGradientMode.Horizontal))
                {
                    var state = g.Save();
                    g.SetClip(path);
                    g.FillRectangle(brush, clipRect);
                    g.Restore(state);
                }
            }
        }

        // Thin outline
        using (var path = RoundedRect(trackRect, radius))
        using (var pen  = new Pen(Color.FromArgb(28, 0, 0, 0), 1))
            g.DrawPath(pen, path);

        // Percentage overlay
        if (ShowPercentage && !_marquee)
        {
            double pct = (_maximum > _minimum)
                ? (double)(_value - _minimum) / (_maximum - _minimum) : 0;
            string text = ((int)Math.Round(pct * 100)) + "%";
            using (var fmt = new StringFormat {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center })
            using (var font = new Font("Segoe UI", 9, FontStyle.Bold))
            {
                using (var shadow = new SolidBrush(Color.FromArgb(110, 0, 0, 0)))
                    g.DrawString(text, font, shadow, new RectangleF(1, 1, Width, Height), fmt);
                using (var fg = new SolidBrush(Color.White))
                    g.DrawString(text, font, fg, new RectangleF(0, 0, Width, Height), fmt);
            }
        }
    }

    private GraphicsPath RoundedRect(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        if (radius <= 0) { path.AddRectangle(bounds); return path; }
        int d = radius * 2;
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}
"@ -ErrorAction Stop

################################################################################
#                           Global Variables                                   #
################################################################################

# Make sure this dictionary always exists before anything tries to read it
if ($null -eq $script:pmSuppressEvents)      { $script:pmSuppressEvents = $false }
if ($null -eq $script:pmTotalSelectedHours)  { $script:pmTotalSelectedHours = 0.0 }

################################################################################
#                            Core Utilities                                    #
################################################################################

$script:configPath = Join-Path $PSScriptRoot 'PartsMgmt\Config.json'

#UI Log
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp - $message"
    # Optionally, you can also write to a log file:
    "$timestamp - $message" | Out-File -Append -FilePath "UI.log"
}

function New-DefaultConfigObject {
    # Returns a fresh config object for the given root.
    param([string]$Root, [string]$ScriptsDir)

    if ([string]::IsNullOrWhiteSpace($Root))       { throw "Root is required." }
    if ([string]::IsNullOrWhiteSpace($ScriptsDir)) { $ScriptsDir = $PSScriptRoot }

    $dd  = Join-Path $Root 'Dropdown CSVs'
    $pr  = Join-Path $Root 'Parts Room'
    $pb  = Join-Path $Root 'Parts Books'
    $wt  = Join-Path $Root 'Work Tracking'

    return [ordered]@{
        RootDirectory         = $Root
        DropdownCsvsDirectory = $dd
        PartsRoomDirectory    = $pr
        PartsBooksDirectory   = $pb

        Books                 = @{}
        SameDayPartsRooms     = @()

        PrerequisiteFiles     = [ordered]@{
            Machines  = Join-Path $dd 'Machines.csv'
        }

        SupervisorEmail       = 'default@example.com'

        WorkTracking          = [ordered]@{
            JobsFile                   = 'Work Tracking/Jobs.json'
            HistorianFile              = 'Work Tracking/Historian.jsonl'
            WeeklyWorksheetsDirectory  = 'Work Tracking/Weekly Worksheets'
            PmChecklistsDirectory      = 'Work Tracking/PM Checklists'
            ReportsDirectory           = 'Work Tracking/Reports'
            eDacUrl                    = ''
            ReactiveToWorkOrderMinutes = 15
            TechnicianName             = ''
            WorkBudget                 = [ordered]@{
                DayLengthHours        = 8.5
                LunchMinutes          = 30
                WashupMinutes         = 15
                StartupMinutes        = 15
                PaidBreaksMinutes     = @(15, 15)
                EndOfDayWashupMinutes = 15
                PaperworkMinutes      = 15
                WorkTargetHours       = 6.5
            }
        }
    }
}

#Initialize the config
function Initialize-Config {
    if (Test-Path $configPath) {
        $config = Get-Content -Path $script:configPath | ConvertFrom-Json
        if (-not ($config.PSObject.Properties.Name -contains 'SameDayPartsRooms')) {
            $config | Add-Member -NotePropertyName 'SameDayPartsRooms' -NotePropertyValue @()
            Write-Log "Backfilled missing SameDayPartsRooms key."
        }
        if (-not ($config.PSObject.Properties.Name -contains 'Books')) {
            $config | Add-Member -NotePropertyName 'Books' -NotePropertyValue @{}
            Write-Log "Backfilled missing Books key."
        }
        if (-not ($config.PSObject.Properties.Name -contains 'WorkTracking')) {
            $defaults = New-DefaultConfigObject -Root $config.RootDirectory -ScriptsDir $PSScriptRoot
            $config | Add-Member -NotePropertyName 'WorkTracking' -NotePropertyValue $defaults.WorkTracking
            Write-Log "Backfilled missing WorkTracking key."
        }
        if (-not ($config.PSObject.Properties.Name -contains 'SupervisorEmail')) {
            $config | Add-Member -NotePropertyName 'SupervisorEmail' -NotePropertyValue 'default@example.com'
            Write-Log "Backfilled missing SupervisorEmail key."
        }
        Write-Log "Config loaded successfully"
    } else {
        Write-Log "Config file not found. Using default configuration."
        $config = New-DefaultConfigObject -Root $PSScriptRoot -ScriptsDir $PSScriptRoot
    }

    return $config
}

$config = Initialize-Config

# Helper function to create buttons with optional color styles
function New-Button {
    param(
        [string]$text,
        [scriptblock]$action,
        [object]$Tag = $null,
        [ValidateSet('Primary','Success','Danger','Secondary','Ghost')]
        [string]$Style = 'Primary',
        [int]$Width = 450,
        [int]$Height = 36,
        [System.Drawing.ContentAlignment]$TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Width = $Width
    $button.Height = $Height
    $button.Margin = New-Object System.Windows.Forms.Padding(6)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.TextAlign = $TextAlign
    $button.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $button.UseVisualStyleBackColor = $false

    switch ($Style) {
        'Primary'   { $base = [System.Drawing.Color]::FromArgb(52,152,219);  $button.ForeColor = [System.Drawing.Color]::White }
        'Success'   { $base = [System.Drawing.Color]::FromArgb(39,174,96);   $button.ForeColor = [System.Drawing.Color]::White }
        'Danger'    { $base = [System.Drawing.Color]::FromArgb(231,76,60);   $button.ForeColor = [System.Drawing.Color]::White }
        'Secondary' { $base = [System.Drawing.Color]::FromArgb(189,195,199); $button.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80) }
        'Ghost'     { $base = [System.Drawing.Color]::FromArgb(236,240,241); $button.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80) }
    }
    $button.BackColor = $base

    # Native flat-button hover/pressed states — no event-handler closures needed
    $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Min(255, $base.R + 25),
        [Math]::Min(255, $base.G + 25),
        [Math]::Min(255, $base.B + 25))
    $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $base.R - 25),
        [Math]::Max(0, $base.G - 25),
        [Math]::Max(0, $base.B - 25))

    if ($Tag -ne $null) { $button.Tag = $Tag }
    $button.Add_Click($action)
    return $button
}

# Helper to apply a consistent ListView style
function Set-ListViewStyle {
    param([System.Windows.Forms.ListView]$ListView)
    $ListView.View = [System.Windows.Forms.View]::Details
    $ListView.FullRowSelect = $true
    $ListView.GridLines = $true
    $ListView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $ListView.BackColor = [System.Drawing.Color]::White
    $ListView.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $ListView.HideSelection = $false
}

################################################################################
#                    Data Normalization and Search                             #
################################################################################

# Function to normalize NSNs by removing non-digit characters except '*'
function Normalize-NSN {
    param([string]$nsn)
    if ($nsn -ne $null) {
        return ($nsn -replace '[^\d*]', '')
    } else {
        return ''
    }
}

# Function to normalize OEM numbers by removing common delimiters and converting to uppercase
function Normalize-OEM {
    param([string]$oem)
    if ($oem -ne $null) {
        return ($oem -replace '[\.\-\/\s]', '').ToUpper()
    } else {
        return ''
    }
}

# Function to check if a string matches a pattern with wildcards
function Test-WildcardMatch {
    param(
        [string]$InputString,
        [string]$Pattern
    )
    
    Write-Log "Testing match: Input='$InputString' Pattern='$Pattern'"
    
    if ([string]::IsNullOrWhiteSpace($InputString) -or [string]::IsNullOrWhiteSpace($Pattern)) {
        return $false
    }
    
    # Convert the wildcard pattern to a regex pattern
    # First escape any regex special characters
    $regexPattern = [regex]::Escape($Pattern)
    # Then replace * with .*
    $regexPattern = $regexPattern.Replace('\*', '.*')
    # Add start and end anchors
    $regexPattern = "^$regexPattern$"
    
    Write-Log "Regex pattern: $regexPattern"
    
    $result = $InputString -match $regexPattern
    Write-Log "Match result: $result"
    return $result
}

# Simple helper function to search parts
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description
    )

    $results = @()

    # --- NSN pattern (digits + * wildcard, same rules as main search) ---
    $nsnPattern = $null
    if (-not [string]::IsNullOrWhiteSpace($NSN)) {
        $nsnCleaned = $NSN -replace '[^0-9*]', ''
        if (-not [string]::IsNullOrEmpty($nsnCleaned) -and $nsnCleaned -ne '*') {
            $nsnPattern = ([regex]::Escape($nsnCleaned)) -replace '\\\*', '.*'
        }
    }

    # --- Description tokens (all must be present) ---
    $descTokens = @()
    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        $descTokens = @(
            ($Description -replace '[^A-Za-z0-9]', ' ').ToLower() -split '\s+' |
            Where-Object { $_ -ne '' }
        )
    }

    # --- Row match helper (renamed to avoid the $-hyphen parse trap) ---
    $rowTester = {
        param($row)
        $nsnOk = $true
        if ($nsnPattern) {
            $rowNSN = ([string]$row.'Part (NSN)') -replace '[^0-9]', ''
            $nsnOk = $rowNSN -match $nsnPattern
        }
        $descOk = $true
        if ($descTokens.Count -gt 0) {
            $rowDesc = (([string]$row.Description -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
            foreach ($t in $descTokens) {
                if (-not $rowDesc.Contains($t)) { $descOk = $false; break }
            }
        }
        $nsnOk -and $descOk
    }

    # 1. Local Parts Room
    $localFiles = Get-ChildItem -Path (Join-Path $config.PartsRoomDirectory "*.csv") -File
    foreach ($f in $localFiles) {
        foreach ($row in (Import-Csv -Path $f.FullName)) {
            if ([int]$row.QTY -le 0) { continue }
            if (-not (& $rowTester $row)) { continue }
            $results += [PSCustomObject]@{
                PartNumber = $row.'Part (NSN)'
                Description = $row.Description
                Quantity = [int]$row.QTY
                Location = $row.Location
                OEMNumber = $row.'OEM 1'
                Source = "Local Parts Room"
            }
        }
    }

    # 2. Same Day Parts Room
    $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    if (Test-Path $sameDayDir) {
        foreach ($f in (Get-ChildItem -Path (Join-Path $sameDayDir "*.csv") -File)) {
            $siteName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            foreach ($row in (Import-Csv -Path $f.FullName)) {
                if ([int]$row.QTY -le 0) { continue }
                if (-not (& $rowTester $row)) { continue }
                $results += [PSCustomObject]@{
                    PartNumber = $row.'Part (NSN)'
                    Description = $row.Description
                    Quantity = [int]$row.QTY
                    Location = $row.Location
                    OEMNumber = $row.'OEM 1'
                    Source = "$siteName (Same Day)"
                }
            }
        }
    }

    return $results
}

# Data retrieval logic
function Search-CrossReferenceData {
    param ($NSN, $OEM, $Description)

    $crossRefResults = @()

    foreach ($book in $config.Books.PSObject.Properties) {
        $bookName = $book.Name
        $volumesCsvPath = $book.Value.VolumesToUrlCsvPath
        if ($volumesCsvPath) {
            $bookDir = Split-Path -Path $volumesCsvPath -Parent
        } else {
            $bookDir = Join-Path $config.PartsBooksDirectory $bookName
        }
        $combinedSectionsDir = Join-Path $bookDir "CombinedSections"

        if (-not (Test-Path $combinedSectionsDir)) {
            Write-Log "CombinedSections directory not found for $bookName"
            continue
        }

        $sectionCsvFiles = Get-ChildItem -Path $combinedSectionsDir -Filter "*.csv" -File
        Write-Log "Found $($sectionCsvFiles.Count) CSV files in $bookName"

        foreach ($csvFile in $sectionCsvFiles) {
            $csvFilePath = $csvFile.FullName
            $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)

            try {
                $sectionData = Import-Csv -Path $csvFilePath
                Write-Log "Processed $($sectionData.Count) rows from $($csvFile.Name)"
            } catch {
                Write-Log "Failed to read CSV file $csvFilePath. Error: $_"
                continue
            }

            $filteredSectionData = $sectionData | Where-Object {
                ($NSN -eq '' -or $_.'STOCK NO.' -like "*$NSN*") -and
                ($OEM -eq '' -or $_.'PART NO.' -like "*$OEM*") -and
                ($Description -eq '' -or $_.'PART DESCRIPTION' -like "*$Description*")
            }

            foreach ($item in $filteredSectionData) {
                $resultItem = [PSCustomObject]@{
                    Handbook = $bookName
                    SectionName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
                    NO = if ($item.PSObject.Properties['NO']) { $item.NO } else { "" }
                    PartDescription = if ($item.PSObject.Properties['PART DESCRIPTION']) { $item.'PART DESCRIPTION' } else { "" }
                    REF = if ($item.PSObject.Properties['REF.']) { $item.'REF.' } else { "" }
                    StockNo = if ($item.PSObject.Properties['STOCK NO.']) { $item.'STOCK NO.' } else { "" }
                    PartNo = if ($item.PSObject.Properties['PART NO.']) { $item.'PART NO.' } else { "" }
                    Location = if ($item.PSObject.Properties['Location']) { $item.Location } else { "" }
                    Source = $sourceFileName
                }
                $crossRefResults += $resultItem
            }
        }
    }

    return $crossRefResults
}

################################################################################
#                           File Management                                    #
################################################################################

# Function to create and format Excel file from CSV
function Create-ExcelFromCsv {
    param(
        [string]$siteName,
        [string]$csvDirectory,
        [string]$excelDirectory,
        [string]$tableName = "My_Parts_Room"
    )

    $excel = $null
    $workbook = $null
    $progressForm = $null

    try {

        # Add timing diagnostics
        $startTime = Get-Date

        # Show progress form if not already visible
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Creating Parts Room Excel"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = 'CenterScreen'

        $progressBar = New-Object ModernProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(360, 24)
        $progressBar.Location = New-Object System.Drawing.Point(10, 12)
        $progressForm.Controls.Add($progressBar)

        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Size = New-Object System.Drawing.Size(360, 40)
        $progressLabel.Location = New-Object System.Drawing.Point(10, 44)
        $progressForm.Controls.Add($progressLabel)

        $timeLabel = New-Object System.Windows.Forms.Label
        $timeLabel.Size = New-Object System.Drawing.Size(360, 20)
        $timeLabel.Location = New-Object System.Drawing.Point(10, 92)
        $progressForm.Controls.Add($timeLabel)

        $progressForm.Show()
        $progressForm.Refresh()

        $csvFilePath = Join-Path $csvDirectory "$siteName.csv"
        $excelFilePath = Join-Path $excelDirectory "$siteName.xlsx"

        $progressLabel.Text = "Verifying CSV file..."
        $progressBar.Value = 5
        $progressForm.Refresh()

        if (-not (Test-Path $csvFilePath)) {
            throw "CSV file not found at $csvFilePath"
        }

        $progressLabel.Text = "Creating Excel application..."
        $progressBar.Value = 10
        $progressForm.Refresh()

        # Create Excel
        $excelStart = Get-Date
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excelDuration = ((Get-Date) - $excelStart).TotalSeconds
        Write-Log "Excel application created in $excelDuration seconds"
        $timeLabel.Text = "Excel app created in $excelDuration seconds"
        $progressForm.Refresh()

        # ============================================================
        # Load the CSV directly.  This is the load-bearing change.
        #
        # Prior versions built a .NET 2D array and assigned it to
        # Range.Value2 in one shot.  That works in a standalone script
        # but throws "Unable to cast object of type 'System.Object[,]'
        # to type 'System.String'" inside UI-Script.ps1, because
        # PowerShell 5.1's COM adapter marshals the array as a scalar
        # in some contexts.  We cannot reliably control that from the
        # caller's side, so we sidestep it entirely: Excel reads the
        # CSV file with its own native parser, no array ever crosses
        # the PowerShell/COM boundary.
        # ============================================================
        $progressLabel.Text = "Loading CSV into Excel..."
        $progressBar.Value = 30
        $progressForm.Refresh()

        $loadStart = Get-Date
        # Open(FilePath, UpdateLinks, ReadOnly).  We pass 0 for UpdateLinks
        # (no external links to refresh) and $false for ReadOnly (we intend
        # to SaveAs the loaded workbook to a new path).
        $workbook = $excel.Workbooks.Open($csvFilePath, 0, $false)
        $loadDuration = ((Get-Date) - $loadStart).TotalSeconds
        Write-Log "CSV loaded into Excel in $loadDuration seconds"
        $timeLabel.Text = "CSV loaded in $loadDuration seconds"
        $progressForm.Refresh()

        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Parts Data"

        # Row 1 holds the header row (Export-Csv wrote it).  Data starts at
        # row 2.  We only need $headers so the "Importing data..." column
        # cleanup can find columns by name.
        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count - 1     # minus the header
        $colCount = [int]$usedRange.Columns.Count
        Write-Log "Loaded sheet size: $rowCount data rows x $colCount columns"

        $headers = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $headers += "$($worksheet.Cells.Item(1, $col).Value2)"
        }

        # ============================================================
        # Formatting — identical to before, applied to the loaded sheet.
        # ============================================================

        $progressLabel.Text = "Formatting table..."
        $progressBar.Value = 70
        $progressForm.Refresh()

        $formatStart = Get-Date
        if ($worksheet.ListObjects.Count -gt 0) {
            $worksheet.ListObjects.Item(1).Unlist()
        }
        $listObject = $worksheet.ListObjects.Add(
            [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
            $usedRange, $null,
            [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
        $listObject.Name = $tableName
        $listObject.TableStyle = "TableStyleMedium2"
        Write-Log "Table formatting completed in $(((Get-Date) - $formatStart).TotalSeconds) seconds"

        $progressLabel.Text = "Applying cell formatting..."
        $progressBar.Value = 80
        $progressForm.Refresh()

        $cellFormatStart = Get-Date
        $usedRange.Cells.VerticalAlignment   = -4108   # xlCenter
        $usedRange.Cells.HorizontalAlignment = -4108   # xlCenter
        $usedRange.Cells.WrapText            = $false
        $usedRange.Cells.Font.Name           = "Courier New"
        $usedRange.Cells.Font.Size           = 12
        Write-Log "Cell formatting completed in $(((Get-Date) - $cellFormatStart).TotalSeconds) seconds"

        $progressLabel.Text = "Auto-fitting columns..."
        $progressBar.Value = 90
        $progressForm.Refresh()

        $autoFitStart = Get-Date
        $usedRange.Columns.AutoFit() | Out-Null
        Write-Log "Column auto-fit completed in $(((Get-Date) - $autoFitStart).TotalSeconds) seconds"

        # Left-align the Description column if it exists
        try {
            $descriptionColumn = $listObject.ListColumns | Where-Object { $_.Name -eq "Description" }
            if ($descriptionColumn) {
                $descriptionColumn.Range.Offset(1, 0).HorizontalAlignment = -4131   # xlLeft
            }
        } catch {
            Write-Log "Error while formatting Description column: $_"
        }

        # Remove any columns named "Importing data..."
        try {
            for ($col = $colCount; $col -ge 1; $col--) {
                $columnHeader = "$($worksheet.Cells.Item(1, $col).Value2)"
                if ($columnHeader -eq "Importing data...") {
                    $worksheet.Columns.Item($col).Delete()
                    Write-Log "Removed 'Importing data...' column at position $col"
                }
            }
        } catch {
            Write-Log "Error while removing columns: $_"
        }

        $progressLabel.Text = "Saving Excel file..."
        $progressBar.Value = 95
        $progressForm.Refresh()

        # Save and close.  SaveAs to xlsx with an explicit format so the
        # file lands as .xlsx regardless of any Excel default-format setting.
        $saveStart = Get-Date
        $workbook.SaveAs($excelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
        $workbook.Close($false)
        $workbook = $null
        $excel.Quit()
        $excel = $null
        $saveDuration = ((Get-Date) - $saveStart).TotalSeconds
        Write-Log "Excel file saved in $saveDuration seconds"

        $endTime = Get-Date
        $totalDuration = ($endTime - $startTime).TotalSeconds
        Write-Log "Excel file created successfully at $excelFilePath in total time: $totalDuration seconds"

        $progressBar.Value = 100
        $progressLabel.Text = "Excel file created successfully!"
        $timeLabel.Text = "Total time: $totalDuration seconds"
        $progressForm.Refresh()
        Start-Sleep -Seconds 2
        $progressForm.Close()
    }
    catch {
        Write-Log "Error during Excel file creation: $($_.Exception.Message)"
        if ($progressForm -and $progressForm.Visible) {
            $progressLabel.Text = "Error: $($_.Exception.Message)"
            $progressForm.Refresh()
            Start-Sleep -Seconds 3
            $progressForm.Close()
        }
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ============================================================
# Parts Book figure → CSV
# ============================================================

function Convert-HtmlFigureToCsv {
    param(
        [string]$HtmlPath,
        [string]$CsvPath = $null
    )

    if (-not (Test-Path $HtmlPath)) { throw "HTML file not found: $HtmlPath" }
    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        $CsvPath = [System.IO.Path]::ChangeExtension($HtmlPath, '.csv')
    }

    $htmlContent = Get-Content -Path $HtmlPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($htmlContent)) { throw "Empty HTML: $HtmlPath" }

    # Locate the parts table. It is <TABLE BORDER="1" COLS="5" BORDERCOLOR="#808080">.
    # Attribute order isn't guaranteed; try plausible variations.
    $patterns = @(
        '(?is)<table[^>]*border\s*=\s*"?#?1"?[^>]*cols\s*=\s*"?#?5"?[^>]*bordercolor\s*=\s*"?#?808080"?[^>]*>(.*?)</table>',
        '(?is)<table[^>]*bordercolor\s*=\s*"?#?808080"?[^>]*cols\s*=\s*"?#?5"?[^>]*>(.*?)</table>',
        '(?is)<table[^>]*cols\s*=\s*"?#?5"?[^>]*>(.*?)</table>'
    )

    $tableHtml = $null
    foreach ($p in $patterns) {
        $m = [regex]::Match($htmlContent, $p)
        if ($m.Success) { $tableHtml = $m.Groups[1].Value; break }
    }
    if ($null -eq $tableHtml) { throw "Data table not found in $HtmlPath" }

    $rowMatches = [regex]::Matches($tableHtml, '(?is)<tr[^>]*>(.*?)</tr>')
    if ($rowMatches.Count -le 2) { throw "Data table has no data rows in $HtmlPath" }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('"NO.","PART DESCRIPTION","REF.","STOCK NO.","PART NO.","CAGE"')

    $written = 0
    # Skip rows 0 and 1 (two header rows), same as the old script.
    for ($r = 2; $r -lt $rowMatches.Count; $r++) {
        $rowHtml     = $rowMatches[$r].Groups[1].Value
        $cellMatches = [regex]::Matches($rowHtml, '(?is)<td[^>]*>(.*?)</td>')
        if ($cellMatches.Count -eq 0) { continue }

        $cleanCells = @()
        foreach ($cm in $cellMatches) {
            $text = $cm.Groups[1].Value
            $text = [regex]::Replace($text, '(?is)<[^>]+>', '')
            $text = $text -replace '&nbsp;', ' '
            $text = $text -replace '&amp;',  '&'
            $text = $text -replace '&lt;',   '<'
            $text = $text -replace '&gt;',   '>'
            $text = $text -replace '&quot;', '"'
            $text = $text -replace '&#39;',  "'"
            $text = ($text -replace '\s+', ' ').Trim()
            $cleanCells += '"' + ($text -replace '"', '""') + '"'
        }

        if (($cleanCells -join '') -ne '""""""""""') {
            $lines.Add($cleanCells -join ',')
            $written++
        }
    }

    $csvDir = Split-Path -Path $CsvPath -Parent
    if ($csvDir -and -not (Test-Path $csvDir)) {
        New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
    }

    # UTF-8 without BOM — PS 5.1's Out-File -Encoding UTF8 writes a BOM
    [System.IO.File]::WriteAllLines(
        $CsvPath, $lines, (New-Object System.Text.UTF8Encoding($false)))

    Write-Log "Convert-HtmlFigureToCsv: $written row(s) -> $CsvPath"
    return $written
}

function Convert-FigureFolderToCsv {
    param(
        [string]$HtmlDir,
        [switch]$Overwrite
    )

    if (-not (Test-Path $HtmlDir)) { throw "Folder not found: $HtmlDir" }

    $converted = 0; $skipped = 0; $failed = 0; $errors = @()

    $files = @(Get-ChildItem -Path $HtmlDir -Filter "*.html" -File |
        Sort-Object -Property @{
            Expression = { if ($_.BaseName -match 'Figure (\d+)-(\d+)') { [int]$matches[1] * 10000 + [int]$matches[2] } else { 0 } }
        })

    foreach ($f in $files) {
        $csv = [System.IO.Path]::ChangeExtension($f.FullName, '.csv')
        if ((Test-Path $csv) -and -not $Overwrite) { $skipped++; continue }
        try {
            Convert-HtmlFigureToCsv -HtmlPath $f.FullName -CsvPath $csv | Out-Null
            $converted++
        } catch {
            $failed++
            $errors += "$($f.Name): $($_.Exception.Message)"
        }
    }

    return [PSCustomObject]@{
        Converted = $converted
        Skipped   = $skipped
        Failed    = $failed
        Errors    = $errors
    }
}

function Show-FigureConvertDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Convert Figure HTML to CSV'
    $form.Size = New-Object System.Drawing.Size(600, 340)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblIntro = New-Object System.Windows.Forms.Label
    $lblIntro.Text = "Converts figure HTML files from a parts book into CSVs. The CSVs are written alongside the HTMLs."
    $lblIntro.Location = New-Object System.Drawing.Point(14, 14)
    $lblIntro.Size = New-Object System.Drawing.Size(560, 40)
    $lblIntro.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $form.Controls.Add($lblIntro)

    $lblDir = New-Object System.Windows.Forms.Label
    $lblDir.Text = "Figure folder:"
    $lblDir.Location = New-Object System.Drawing.Point(14, 70)
    $lblDir.Size = New-Object System.Drawing.Size(100, 24)
    $lblDir.TextAlign = 'MiddleRight'
    $lblDir.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblDir)

    $txtDir = New-Object System.Windows.Forms.TextBox
    $txtDir.Location = New-Object System.Drawing.Point(120, 70)
    $txtDir.Size = New-Object System.Drawing.Size(360, 24)
    $form.Controls.Add($txtDir)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point(486, 70)
    $btnBrowse.Size = New-Object System.Drawing.Size(90, 24)
    $btnBrowse.FlatStyle = 'Flat'
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnBrowse.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBrowse.Cursor = 'Hand'
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Pick the folder containing figure HTML files"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtDir.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowse)

    $chkOverwrite = New-Object System.Windows.Forms.CheckBox
    $chkOverwrite.Text = "Overwrite existing CSVs"
    $chkOverwrite.Location = New-Object System.Drawing.Point(120, 102)
    $chkOverwrite.Size = New-Object System.Drawing.Size(200, 24)
    $chkOverwrite.Checked = $false
    $form.Controls.Add($chkOverwrite)

    $txtResult = New-Object System.Windows.Forms.TextBox
    $txtResult.Multiline = $true
    $txtResult.ReadOnly = $true
    $txtResult.ScrollBars = 'Vertical'
    $txtResult.Location = New-Object System.Drawing.Point(14, 136)
    $txtResult.Size = New-Object System.Drawing.Size(560, 130)
    $txtResult.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtResult.BackColor = [System.Drawing.Color]::White
    $txtResult.Text = "(no run yet)"
    $form.Controls.Add($txtResult)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Convert Folder"
    $btnRun.Location = New-Object System.Drawing.Point(14, 276)
    $btnRun.Size = New-Object System.Drawing.Size(140, 30)
    $btnRun.FlatStyle = 'Flat'
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRun.Cursor = 'Hand'
    $btnRun.Add_Click({
        $dir = $txtDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path $dir)) {
            $txtResult.Text = "Pick a valid folder first."
            return
        }
        try {
            $r = Convert-FigureFolderToCsv -HtmlDir $dir -Overwrite:$chkOverwrite.Checked
            $msg = "Converted: $($r.Converted)`r`nSkipped:   $($r.Skipped)`r`nFailed:    $($r.Failed)"
            if ($r.Errors.Count -gt 0) {
                $msg += "`r`n`r`nErrors:`r`n" + (($r.Errors | Select-Object -First 5) -join "`r`n")
                if ($r.Errors.Count -gt 5) { $msg += "`r`n... ($($r.Errors.Count - 5) more, see UI.log)" }
            }
            $txtResult.Text = $msg
        } catch {
            $txtResult.Text = "Failed: $($_.Exception.Message)"
        }
    })
    $form.Controls.Add($btnRun)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(474, 276)
    $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnClose.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClose.Cursor = 'Hand'
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    $form.CancelButton = $btnClose
    $form.ShowDialog() | Out-Null
}

# ============================================================
# Parts Book Creator — helpers
# ============================================================

function Get-BookKey {
    param([string]$FullName)
    return ($FullName -replace '[^\w\s-]', '' -replace '\s+', ' ').Trim()
}

function Sanitize-BookFolderName {
    param([string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars()
    foreach ($c in $invalid) { $Name = $Name -replace [regex]::Escape($c), '-' }
    return $Name
}

function New-PartsBookProgressForm {
    param([string]$Title = 'Processing Parts Books')

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(520, 200)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $bar = New-Object ModernProgressBar
    $bar.Location = New-Object System.Drawing.Point(12, 12)
    $bar.Size = New-Object System.Drawing.Size(480, 26)
    $form.Controls.Add($bar)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(12, 46)
    $label.Size = New-Object System.Drawing.Size(480, 44)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($label)

    $bookLabel = New-Object System.Windows.Forms.Label
    $bookLabel.Location = New-Object System.Drawing.Point(12, 94)
    $bookLabel.Size = New-Object System.Drawing.Size(480, 22)
    $bookLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $bookLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $form.Controls.Add($bookLabel)

    return [PSCustomObject]@{ Form = $form; Bar = $bar; Label = $label; BookLabel = $bookLabel }
}

function Update-PartsBookProgress {
    param($Handle, [string]$Message, $Percent, [string]$Book = '')

    if (-not $Handle) { return }
    $pct = 0
    try {
        if ($Percent -is [array]) { if ($Percent.Count -gt 0) { $pct = [int][double]$Percent[0] } }
        elseif ($null -ne $Percent -and "$Percent" -ne '') { $pct = [int][double]$Percent }
    } catch { $pct = 0 }
    if ($pct -lt 0) { $pct = 0 }; if ($pct -gt 100) { $pct = 100 }

    $Handle.Bar.Value = $pct
    $Handle.Label.Text = $Message
    if ($Book) { $Handle.BookLabel.Text = "Current book: $Book" }
    $Handle.Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-PartsBookTreeHtml {
    param([string]$BookName)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Paste Parts Book Tree HTML — $BookName"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "In the browser window that opened: View Page Source (Ctrl+U), Ctrl+A, Ctrl+C, paste below."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 28
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(680, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Accept"
    $okBtn.Location = New-Object System.Drawing.Point(780, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.Anchor = 'Top,Right'
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPaste.Text)) { return }
        $form.Tag = $txtPaste.Text
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    $txtPaste = New-Object System.Windows.Forms.TextBox
    $txtPaste.Multiline = $true
    $txtPaste.MaxLength = 0
    $txtPaste.Dock = 'Fill'
    $txtPaste.ScrollBars = 'Both'
    $txtPaste.WordWrap = $false
    $txtPaste.Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtPaste.AcceptsReturn = $true
    $txtPaste.Padding = New-Object System.Windows.Forms.Padding(8)
    $form.Controls.Add($txtPaste)
    $txtPaste.BringToFront()

    $form.CancelButton = $cancelBtn
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function ConvertFrom-PartsBookTree {
    param(
        [string]$HtmlContent,
        $Row,
        [string]$DirectoryPath
    )

    if ([string]::IsNullOrWhiteSpace($HtmlContent)) { throw "Empty HTML." }
    if (-not (Test-Path $DirectoryPath)) { New-Item -ItemType Directory -Path $DirectoryPath -Force | Out-Null }

    $htmlDoc = $null
    try {
        $htmlDoc = New-Object -ComObject "HTMLFile"
        try   { $htmlDoc.IHTMLDocument2_write($HtmlContent) }
        catch {
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($HtmlContent)
            $htmlDoc.write($bytes)
        }

        $msBookNo = "$($Row.'MS Book No')"
        $volume   = "$($Row.Volume)"

        if ([string]::IsNullOrWhiteSpace($msBookNo) -or [string]::IsNullOrWhiteSpace($volume)) {
            $titleEl = $htmlDoc.getElementById("book_title")
            if ($titleEl -and $titleEl.innerText -match "MS(\d+)\s+VOLUME\s+([A-Z])") {
                $msBookNo = $matches[1]; $volume = $matches[2]
            }
        }
        if ([string]::IsNullOrWhiteSpace($msBookNo) -or [string]::IsNullOrWhiteSpace($volume)) {
            throw "Could not determine MS Book No / Volume for this book."
        }

        $volumesToUrlPath  = Join-Path $DirectoryPath "Volumes-to-URL.csv"
        $sectionNamesPath  = Join-Path $DirectoryPath "SectionNames.txt"

        "Figure No.,Name,Section No.,MS Book No,Volume,URL" |
            Out-File -FilePath $volumesToUrlPath -Encoding UTF8

        $container = $htmlDoc.getElementById("Ryan_fault")
        if (-not $container) {
            foreach ($d in $htmlDoc.getElementsByTagName("div")) {
                try {
                    $s = $d.getAttribute("style")
                    if ($s -and $s -like "*height:100%*overflow-y:scroll*") { $container = $d; break }
                } catch { }
            }
        }
        if (-not $container) { throw "Could not find phbk content container." }

        $treeUl = $null
        foreach ($u in $container.getElementsByTagName("ul")) {
            try {
                if ($u.id -eq "phbk_tree" -or $u.className -eq "treeview") { $treeUl = $u; break }
            } catch { }
        }
        if (-not $treeUl) { throw "Could not find phbk_tree." }

        # Collect sections, then sort numerically ascending
        $sectionData = @()
        foreach ($secLi in @($treeUl.getElementsByTagName("li") | Where-Object { $_.getAttribute("sno") })) {
            $sectionNo = $secLi.getAttribute("sno")
            $secSpan   = $secLi.getElementsByTagName("span") | Where-Object { $_.className -ne "go_fig" } | Select-Object -First 1
            if (-not $secSpan) { continue }

            $secText  = $secSpan.innerText -replace '^Section \d+\s+', ''
            $secName  = "Section $sectionNo $secText"
            $figList  = $secLi.getElementsByTagName("ul") | Select-Object -First 1

            $figures = @()
            if ($figList) {
                foreach ($figLi in @($figList.getElementsByTagName("li") | Where-Object { $_.getAttribute("figno") })) {
                    $figno = $figLi.getAttribute("figno")
                    if (-not $figno) { continue }
                    $figSpan = $figLi.getElementsByTagName("span") | Where-Object { $_.className -eq "go_fig" } | Select-Object -First 1
                    if (-not $figSpan) { continue }
                    $figures += [PSCustomObject]@{
                        Figno = $figno
                        Name  = ($figSpan.innerText -replace "^\d+-\d+\s+", "")
                    }
                }
            }

            $sectionData += [PSCustomObject]@{
                Number  = [int]$sectionNo
                Name    = $secName
                Figures = $figures
            }
        }

        $sectionData = @($sectionData | Sort-Object -Property Number)

        $sectionNames = @()
        $figCount     = 0

        foreach ($sec in $sectionData) {
            $sectionNames += $sec.Name

            # Sort figures numerically within the section too
            $orderedFigs = @($sec.Figures | Sort-Object -Property @{
                Expression = { if ($_.Figno -match '-(\d+)$') { [int]$matches[1] } else { 0 } }
            })

            foreach ($fig in $orderedFigs) {
                $url = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$($sec.Number)&figno=$($fig.Figno)&viewerflag=d&layout=L11"
                "$($fig.Figno),`"$($fig.Name)`",$($sec.Number),$msBookNo,$volume,$url" |
                    Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Append
                $figCount++
            }
        }

        $sectionNames | Out-File -FilePath $sectionNamesPath -Encoding UTF8

        return [PSCustomObject]@{
            VolumesToUrlPath = $volumesToUrlPath
            SectionNamesPath = $sectionNamesPath
            FigureCount      = $figCount
            MsBookNo         = $msBookNo
            Volume           = $volume
        }
    } finally {
        if ($htmlDoc) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function ConvertFrom-PartsBookTreeFallback {
    param([string]$HtmlContent, $Row, [string]$DirectoryPath)

    $msBookNo = "$($Row.'MS Book No')"
    $volume   = "$($Row.Volume)"

    $volumesToUrlPath = Join-Path $DirectoryPath "Volumes-to-URL.csv"
    $sectionNamesPath = Join-Path $DirectoryPath "SectionNames.txt"

    "Figure No.,Name,Section No.,MS Book No,Volume,URL" |
        Out-File -FilePath $volumesToUrlPath -Encoding UTF8

    $sectionNames = @()
    $figCount = 0

    $secRx = '<li[^>]*sno="(\d+)"[^>]*><div[^>]*></div><span[^>]*>(.*?)</span>'
    $secMatches = [regex]::Matches($HtmlContent, $secRx)

    $parsed = @()
    foreach ($sm in $secMatches) {
        $parsed += [PSCustomObject]@{
            Number = [int]$sm.Groups[1].Value
            Text   = $sm.Groups[2].Value
        }
    }
    $parsed = @($parsed | Sort-Object -Property Number)

    foreach ($sec in $parsed) {
        $sectionNo = $sec.Number
        $secText   = $sec.Text -replace '^Section \d+\s+', ''
        $sectionNames += "Section $sectionNo $secText"

        $figRx = '<li[^>]*figno="' + $sectionNo + '-(\d+)"[^>]*><span\s+class="go_fig">(.*?)</span>'
        $figMatches = [regex]::Matches($HtmlContent, $figRx)
        $figMatches = @($figMatches | Sort-Object -Property @{ Expression = { [int]$_.Groups[1].Value } })

        foreach ($fm in $figMatches) {
            $figno = "$sectionNo-$($fm.Groups[1].Value)"
            $cleanFigName = ($fm.Groups[2].Value -replace "^\d+-\d+\s+", "")
            $url = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figno&viewerflag=d&layout=L11"
            "$figno,`"$cleanFigName`",$sectionNo,$msBookNo,$volume,$url" |
                Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Append
            $figCount++
        }
    }

    $sectionNames | Out-File -FilePath $sectionNamesPath -Encoding UTF8

    return [PSCustomObject]@{
        VolumesToUrlPath = $volumesToUrlPath
        SectionNamesPath = $sectionNamesPath
        FigureCount      = $figCount
        MsBookNo         = $msBookNo
        Volume           = $volume
    }
}

function Export-PartsBookFigures {
    param(
        $VolumesToUrlData,
        [string]$DirectoryPath,
        [string]$BookName,
        $ProgressHandle
    )

    $htmlCsvDir = Join-Path $DirectoryPath "HTML and CSV Files"
    if (-not (Test-Path $htmlCsvDir)) { New-Item -ItemType Directory -Path $htmlCsvDir -Force | Out-Null }

    # Sort numerically: section first, then figure within section
    $rows = @(ConvertTo-FlatArray -Source $VolumesToUrlData) |
        Sort-Object -Property @{
            Expression = {
                $fn = "$($_.'Figure No.')"
                if ($fn -match '^(\d+)-(\d+)$') { [int]$matches[1] * 10000 + [int]$matches[2] }
                else { 0 }
            }
        }

    $total = $rows.Count
    if ($total -eq 0) { return 0 }

    $done = 0
    foreach ($row in $rows) {
        $done++
        Update-PartsBookProgress -Handle $ProgressHandle `
            -Message "Downloading figure $done of $total" `
            -Percent ([double]$done / $total * 100) `
            -Book $BookName

        $figno = "$($row.'Figure No.')" -replace '[^\w\d-]', '_'
        $url   = "$($row.URL)"
        if ([string]::IsNullOrWhiteSpace($url)) { continue }

        $htmlPath = Join-Path $htmlCsvDir "Figure $figno.html"
        try {
            $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
            if ($resp.StatusCode -eq 200) {
                Set-Content -Path $htmlPath -Value $resp.Content -Encoding UTF8
                Convert-HtmlFigureToCsv -HtmlPath $htmlPath | Out-Null
            }
        } catch {
            Write-Log "Export-PartsBookFigures: failed $figno : $($_.Exception.Message)"
        }
    }
    return $done
}

function Merge-PartsBookSections {
    param(
        [string]$SourceDir,
        [string]$SiteCsvPath,
        [string]$PartsBookName
    )

    $combinedDir = Join-Path $SourceDir "CombinedSections"
    if (-not (Test-Path $combinedDir)) { New-Item -ItemType Directory -Path $combinedDir -Force | Out-Null }

    $csvFiles = @(Get-ChildItem -Path (Join-Path $SourceDir "HTML and CSV Files") -Filter "Figure *.csv" -ErrorAction SilentlyContinue)

    # Numeric ascending group sort (Section 2, 3, ..., 9, 10, ..., 20)
    $groups = @(
        $csvFiles |
            Group-Object { $_.BaseName -replace 'Figure (\d+)-\d+', '$1' } |
            Sort-Object -Property @{ Expression = { [int]$_.Name } }
    )

    $siteData = $null
    if ($SiteCsvPath -and (Test-Path $SiteCsvPath)) {
        $siteData = Import-Csv -Path $SiteCsvPath
        if ($siteData.Count -gt 0) {
            if (-not $siteData[0].PSObject.Properties['Changed Part (NSN)']) {
                $siteData | ForEach-Object { $_ | Add-Member -NotePropertyName 'Changed Part (NSN)' -NotePropertyValue '' -Force }
            }
            $bookCol = $PartsBookName -replace '[^\w\s-]', '' -replace '\s+', ' '
            if (-not $siteData[0].PSObject.Properties[$bookCol]) {
                $siteData | ForEach-Object { $_ | Add-Member -NotePropertyName $bookCol -NotePropertyValue '' -Force }
            }
        }
    }

    foreach ($group in $groups) {
        $figNum = $group.Name
        $sectionFile = Join-Path $combinedDir "Section $figNum.csv"
        $allRows = @()

        # Numeric ascending within a section (2-1, 2-2, ..., 2-10)
        $orderedFiles = @($group.Group | Sort-Object -Property @{
            Expression = { if ($_.BaseName -match 'Figure \d+-(\d+)') { [int]$matches[1] } else { 0 } }
        })

        foreach ($file in $orderedFiles) {
            try { $csvContent = @(Import-Csv -Path $file.FullName) } catch { continue }
            if ($csvContent.Count -eq 0) { continue }

            if (-not $csvContent[0].PSObject.Properties['Location']) {
                $csvContent | ForEach-Object { $_ | Add-Member -NotePropertyName 'Location' -NotePropertyValue '' -Force }
            }
            if (-not $csvContent[0].PSObject.Properties['QTY']) {
                $csvContent | ForEach-Object { $_ | Add-Member -NotePropertyName 'QTY' -NotePropertyValue '' -Force }
            }

            foreach ($row in $csvContent) {
                if ($row.'STOCK NO.' -eq "" -and $row.'PART NO.' -eq "" -and $row.'CAGE' -eq "" -and $row.'Location' -eq "") { continue }
                if (-not $row.'REF.') { $row.'REF.' = $file.Name }

                if ($siteData) {
                    $stockNo = $row.'STOCK NO.'
                    $partNo  = $row.'PART NO.'

                    $match = $siteData | Where-Object { $_.'Part (NSN)' -eq $stockNo } | Select-Object -First 1
                    if ($match) {
                        $row.'Location' = $match.'Location'
                        $row.'QTY'      = $match.'QTY'
                        $idx = $siteData.IndexOf($match)
                        $bookCol = $PartsBookName -replace '[^\w\s-]', '' -replace '\s+', ' '
                        $figRef = "Figure " + ($row.'REF.' -replace '^Figure\s+', '' -replace '\.csv$', '')
                        $existing = "$($siteData[$idx].$bookCol)"
                        if ([string]::IsNullOrEmpty($existing)) { $siteData[$idx].$bookCol = $figRef }
                        elseif ($existing -notmatch [regex]::Escape($figRef)) { $siteData[$idx].$bookCol = "$existing | $figRef" }
                    } else {
                        $oemMatch = $siteData | Where-Object { $_.'OEM 1' -eq $partNo -or $_.'OEM 2' -eq $partNo -or $_.'OEM 3' -eq $partNo } | Select-Object -First 1
                        if ($oemMatch) {
                            $idx = $siteData.IndexOf($oemMatch)
                            $prev = $oemMatch.'Part (NSN)'
                            if ($row.'STOCK NO.' -eq "NSL" -or [string]::IsNullOrEmpty($row.'STOCK NO.')) {
                                $siteData[$idx].'Changed Part (NSN)' = 'No standard NSN'
                            } else {
                                $siteData[$idx].'Changed Part (NSN)' = $prev
                                $siteData[$idx].'Part (NSN)' = $row.'STOCK NO.'
                            }
                            $row.'Location' = $siteData[$idx].'Location'
                            $row.'QTY'      = $siteData[$idx].'QTY'
                            $bookCol = $PartsBookName -replace '[^\w\s-]', '' -replace '\s+', ' '
                            $figRef = "Figure " + ($row.'REF.' -replace '^Figure\s+', '' -replace '\.csv$', '')
                            $existing = "$($siteData[$idx].$bookCol)"
                            if ([string]::IsNullOrEmpty($existing)) { $siteData[$idx].$bookCol = $figRef }
                            elseif ($existing -notmatch [regex]::Escape($figRef)) { $siteData[$idx].$bookCol = "$existing | $figRef" }
                        } else {
                            $row.'Location' = 'Not Stocked Locally'
                        }
                    }
                } else {
                    $row.'Location' = 'Site data not available'
                }

                $allRows += $row
            }
        }

        $allRows | Export-Csv -Path $sectionFile -NoTypeInformation
    }

    if ($siteData -and (Test-Path $SiteCsvPath)) {
        $siteData | Export-Csv -Path $SiteCsvPath -NoTypeInformation
        Write-Log "Merge-PartsBookSections: updated $SiteCsvPath"
    }

    return $combinedDir
}

function New-PartsBookExcel {
    param(
        [string]$SourceDir,
        [string]$CombinedCsvDir,
        [string]$BookName,
        $ProgressHandle
    )

    $excelPath = Join-Path $SourceDir "$((Split-Path $SourceDir -Leaf)).xlsx"
    $excel = $null; $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        while ($workbook.Sheets.Count -gt 1) { $workbook.Sheets.Item(1).Delete() }

        $sectionCsvFiles = @(Get-ChildItem -Path $CombinedCsvDir -Filter "Section *.csv" -ErrorAction SilentlyContinue)
        if ($sectionCsvFiles.Count -eq 0) { throw "No section CSVs in $CombinedCsvDir" }

        # Numeric ascending (Section 2, 3, ..., 20)
        $sectionCsvFiles = @($sectionCsvFiles | Sort-Object -Property @{
            Expression = { if ($_.BaseName -match '^Section (\d+)$') { [int]$matches[1] } else { 0 } }
        })

        $total = [int]$sectionCsvFiles.Count
        $i = 0

        foreach ($file in $sectionCsvFiles) {
            $i++
            Update-PartsBookProgress -Handle $ProgressHandle `
                -Message "Building worksheet $i of $total : $($file.BaseName)" `
                -Percent ([double]$i / $total * 100) `
                -Book $BookName

            $csvContent = @(Import-Csv -Path $file.FullName)
            $ws = $workbook.Sheets.Add()
            $ws.Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            $headers = @('NO.','STOCK NO.','PART DESCRIPTION','PART NO.','REF.','QTY','LOCATION','CAGE')
            for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1, $c + 1) = $headers[$c] }

            $rowCount = $csvContent.Count
            if ($rowCount -gt 0) {
                $arr = New-Object 'object[,]' $rowCount, $headers.Count
                for ($r = 0; $r -lt $rowCount; $r++) {
                    for ($c = 0; $c -lt $headers.Count; $c++) {
                        $arr[$r, $c] = $csvContent[$r].$($headers[$c])
                    }
                }
                $rng = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($rowCount + 1, $headers.Count))
                $rng.Value2 = $arr
            }

            $used = $ws.UsedRange
            if ($ws.ListObjects.Count -gt 0) { $ws.ListObjects.Item(1).Unlist() }
            $lo = $ws.ListObjects.Add(
                [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
                $used, $null,
                [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
            $lo.Name = "$($ws.Name)Table"
            $lo.TableStyle = "TableStyleMedium2"

            foreach ($col in $lo.ListColumns) {
                $col.Range.EntireColumn.AutoFit() | Out-Null
                $col.Range.VerticalAlignment = -4108
                if ($col.Name -eq 'PART DESCRIPTION') {
                    $col.Range.Cells(1,1).HorizontalAlignment = -4108
                    $col.Range.Offset(1,0).HorizontalAlignment = -4131
                } else {
                    $col.Range.HorizontalAlignment = -4108
                }
            }

            $refIdx = -1
            for ($c = 0; $c -lt $headers.Count; $c++) { if ($headers[$c] -eq 'REF.') { $refIdx = $c + 1; break } }
            if ($refIdx -gt 0 -and $rowCount -gt 0) {
                for ($r = 2; $r -le $rowCount + 1; $r++) {
                    $v = $ws.Cells.Item($r, $refIdx).Value2
                    if ($v -and "$v" -match 'Figure \d+-\d+') {
                        $figName = "$v"
                        $htmlPath = Join-Path $SourceDir "HTML and CSV Files\$figName.html"
                        if (Test-Path $htmlPath) {
                            $cell = $ws.Cells.Item($r, $refIdx)
                            $ws.Hyperlinks.Add($cell, $htmlPath, "", "", $figName) | Out-Null
                        }
                    }
                }
            }
        }

        Update-PartsBookProgress -Handle $ProgressHandle -Message "Saving workbook..." -Percent 95 -Book $BookName
        $workbook.SaveAs($excelPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
        return $excelPath
    } finally {
        if ($workbook) { try { $workbook.Close($false) } catch { } }
        if ($excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Rename-PartsBookWorksheets {
    param([string]$ExcelFilePath, [hashtable]$SectionNameMap)

    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($ExcelFilePath)

        foreach ($sheet in @($wb.Sheets)) {
            $cur = $sheet.Name
            if ($cur -eq 'Sheet1') { $sheet.Delete(); continue }

            $target = $null
            if ($SectionNameMap.ContainsKey($cur)) { $target = $SectionNameMap[$cur] }
            elseif ($cur -match '^Section (\d+)') {
                $prefix = "Section $($matches[1])"
                foreach ($k in $SectionNameMap.Keys) {
                    if ($k -match "^$prefix") { $target = $SectionNameMap[$k]; break }
                }
            }

            if ($target) {
                $safe = $target.Substring(0, [Math]::Min(31, $target.Length)) -replace '[:\\/?*\[\]]', ''
                if (-not [string]::IsNullOrWhiteSpace($safe)) { $sheet.Name = $safe }
            }
        }

        $wb.Save()
    } finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-ConfigBooks {
    $books = @{}
    if (-not $script:config -or -not ($script:config.PSObject.Properties.Name -contains 'Books')) { return $books }
    $b = $script:config.Books
    if ($null -eq $b) { return $books }
    foreach ($prop in @($b.PSObject.Properties)) { $books[$prop.Name] = $prop.Value }
    return $books
}

function Save-ConfigBooks {
    param([hashtable]$NewBooks)

    if (-not ($script:config.PSObject.Properties.Name -contains 'Books') -or $null -eq $script:config.Books) {
        $script:config | Add-Member -NotePropertyName 'Books' -NotePropertyValue @{} -Force
    }
    foreach ($k in $NewBooks.Keys) {
        if ($script:config.Books.PSObject.Properties[$k]) {
            $script:config.Books.$k = $NewBooks[$k]
        } else {
            $script:config.Books | Add-Member -NotePropertyName $k -NotePropertyValue $NewBooks[$k] -Force
        }
    }
    $script:config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
}

function Show-PartsBookCreatorDialog {

    $parsedCsvPath = Join-Path $script:config.DropdownCsvsDirectory "Parsed-Parts-Volumes.csv"
    if (-not (Test-Path $parsedCsvPath) -or (Get-Item $parsedCsvPath).Length -eq 0) {
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "Parsed-Parts-Volumes.csv is missing or empty.`r`n`r`nGenerate it now?",
            "Parts Book Creator",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        if (-not (Show-PartsVolumesWizard)) { return }
    }

    $csvData = Import-Csv $parsedCsvPath
    if ($csvData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Parsed-Parts-Volumes.csv contains no rows.", "Parts Book Creator", "OK", "Warning")
        return
    }

    $existingBooks = Get-ConfigBooks

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Create Parts Books"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Check the books you want to build, then click Process Selected."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 26
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Size = New-Object System.Drawing.Size(100, 32)
    $closeBtn.Location = New-Object System.Drawing.Point(680, 10)
    $closeBtn.Anchor = 'Top,Right'
    $closeBtn.FlatStyle = 'Flat'
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Cursor = 'Hand'
    $closeBtn.Add_Click({ $form.Close() })
    $btnBar.Controls.Add($closeBtn)

    $processBtn = New-Object System.Windows.Forms.Button
    $processBtn.Text = "Process Selected"
    $processBtn.Size = New-Object System.Drawing.Size(160, 32)
    $processBtn.Location = New-Object System.Drawing.Point(510, 10)
    $processBtn.Anchor = 'Top,Right'
    $processBtn.FlatStyle = 'Flat'
    $processBtn.FlatAppearance.BorderSize = 0
    $processBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $processBtn.ForeColor = [System.Drawing.Color]::White
    $processBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $processBtn.Cursor = 'Hand'
    $btnBar.Controls.Add($processBtn)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    $lv.CheckBoxes = $true
    $lv.Columns.Add("Full Name", 420)  | Out-Null
    $lv.Columns.Add("MS Book No", 100) | Out-Null
    $lv.Columns.Add("Volume", 80)      | Out-Null
    $lv.Columns.Add("Status", 130)     | Out-Null
    Set-ListViewStyle -ListView $lv
    $form.Controls.Add($lv)
    $lv.BringToFront()

    foreach ($row in $csvData) {
        $key = Get-BookKey $row.'Full Name'
        $it = New-Object System.Windows.Forms.ListViewItem("$($row.'Full Name')")
        $it.SubItems.Add("$($row.'MS Book No')") | Out-Null
        $it.SubItems.Add("$($row.Volume)")       | Out-Null
        if ($existingBooks.ContainsKey($key)) {
            $it.SubItems.Add("already installed") | Out-Null
            $it.ForeColor = [System.Drawing.Color]::FromArgb(140,150,165)
        } else {
            $it.SubItems.Add("") | Out-Null
        }
        $it.Tag = $row
        $lv.Items.Add($it) | Out-Null
    }

    $processBtn.Add_Click({
        $selected = @($lv.CheckedItems | Where-Object {
            (Get-BookKey $_.Tag.'Full Name') -notin $existingBooks.Keys
        })
        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select at least one book that isn't already installed.",
                "Parts Book Creator", "OK", "Warning")
            return
        }

        $form.Hide()

        $progress = New-PartsBookProgressForm -Title "Creating Parts Books"
        $progress.Form.Show()
        $progress.Form.Refresh()

        $newBooks = @{}
        $partsRoomCsv = Get-ChildItem -Path $script:config.PartsRoomDirectory -Filter "*.csv" -File -ErrorAction SilentlyContinue |
                        Select-Object -First 1 -ExpandProperty FullName

        try {
            $idx = 0
            foreach ($item in $selected) {
                $idx++
                $row = $item.Tag
                $bookName = "$($row.'Full Name')"
                $safeName = Sanitize-BookFolderName $bookName
                $bookDir = Join-Path $script:config.PartsBooksDirectory $safeName
                if (-not (Test-Path $bookDir)) { New-Item -ItemType Directory -Path $bookDir -Force | Out-Null }

                Update-PartsBookProgress -Handle $progress `
                    -Message "Book $idx of $($selected.Count): $bookName" `
                    -Percent ([double]($idx-1) / $selected.Count * 100) `
                    -Book $bookName

                $url = "https://www1.mtsc.usps.gov/apps/phbk/index.php?msbookno=$($row.'MS Book No')&volno=$($row.Volume)"
                Start-Process $url

                $html = Get-PartsBookTreeHtml -BookName $bookName
                if ([string]::IsNullOrWhiteSpace($html)) {
                    Write-Log "Parts Book Creator: skipped $bookName (no HTML pasted)"
                    continue
                }

                try {
                    $tree = ConvertFrom-PartsBookTree -HtmlContent $html -Row $row -DirectoryPath $bookDir
                } catch {
                    Write-Log "Parts Book Creator: DOM parse failed for $bookName, trying fallback: $($_.Exception.Message)"
                    $tree = ConvertFrom-PartsBookTreeFallback -HtmlContent $html -Row $row -DirectoryPath $bookDir
                }

                $newBooks[(Get-BookKey $bookName)] = @{
                    VolumesToUrlCsvPath = $tree.VolumesToUrlPath
                    SectionNamesCsvPath = $tree.SectionNamesPath
                }

                $figData = Import-Csv -Path $tree.VolumesToUrlPath
                Export-PartsBookFigures -VolumesToUrlData $figData -DirectoryPath $bookDir -BookName $bookName -ProgressHandle $progress | Out-Null

                $combined = Merge-PartsBookSections -SourceDir $bookDir -SiteCsvPath $partsRoomCsv -PartsBookName $safeName

                $excelPath = New-PartsBookExcel -SourceDir $bookDir -CombinedCsvDir $combined -BookName $bookName -ProgressHandle $progress

                $sectionMap = @{}
                foreach ($line in (Get-Content -Path $tree.SectionNamesPath)) {
                    if ($line -match '^(Section \d+) (.+)$') { $sectionMap[$matches[1]] = $line }
                }
                Rename-PartsBookWorksheets -ExcelFilePath $excelPath -SectionNameMap $sectionMap
            }

            Save-ConfigBooks -NewBooks $newBooks

            if ($partsRoomCsv) {
                $siteName = [System.IO.Path]::GetFileNameWithoutExtension($partsRoomCsv)
                Create-ExcelFromCsv -siteName $siteName `
                    -csvDirectory $script:config.PartsRoomDirectory `
                    -excelDirectory $script:config.PartsRoomDirectory
            }

            $progress.Form.Close()
            [System.Windows.Forms.MessageBox]::Show(
                "Created $($newBooks.Count) book(s).", "Parts Book Creator", "OK", "Information")
        } catch {
            $progress.Form.Close()
            Write-Log "Parts Book Creator: fatal: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed:`r`n$($_.Exception.Message)", "Parts Book Creator", "OK", "Error")
        } finally {
            $form.Close()
        }
    })

    $form.CancelButton = $closeBtn
    $form.ShowDialog() | Out-Null
}

# Helper function to parse HTML content into CSV data
function Update-SinglePartsRoomFromUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteID,
        [Parameter(Mandatory)][string]$SiteName,
        [Parameter(Mandatory)][string]$TargetCsvPath
    )

    Write-Log "=== Update-SinglePartsRoomFromUrl: $SiteName (ID $SiteID) ==="

    $targetDir = Split-Path -Path $TargetCsvPath -Parent
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
    $htmlFilePath = Join-Path $targetDir "$SiteName.html"
    $url = "http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$SiteID&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

    # 1. Download fresh HTML
    try {
        Write-Log "Downloading $SiteName from $url"
        $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
        Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8
        Write-Log "Saved HTML to $htmlFilePath"
    } catch {
        Write-Log "Download failed for $SiteName : $($_.Exception.Message)"
        return $null
    }

    # 2. Parse (reuses the existing Parse-HTMLToCSV function)
    $parsedData = @(Parse-HTMLToCSV -htmlFilePath $htmlFilePath -siteName $SiteName)
    if ($parsedData.Count -eq 0) {
        Write-Log "No data parsed for $SiteName"
        return $null
    }

    # 3. Merge with existing CSV (preserving any book-reference columns)
    $result = [PSCustomObject]@{
        SiteName    = $SiteName
        TotalParsed = $parsedData.Count
        Updated     = 0
        New         = 0
        Removed     = 0
    }

    $baseColumns = @('Part (NSN)','Description','QTY','13 Period Usage','Location','OEM 1','OEM 2','OEM 3')

    if (Test-Path $TargetCsvPath) {
        $existingData = @(Import-Csv -Path $TargetCsvPath)

        $extraColumns = @()
        if ($existingData.Count -gt 0) {
            $extraColumns = @($existingData[0].PSObject.Properties.Name |
                              Where-Object { $baseColumns -notcontains $_ })
        }

        foreach ($row in $parsedData) {
            foreach ($col in $extraColumns) {
                if ($row.PSObject.Properties.Name -notcontains $col) {
                    $row | Add-Member -NotePropertyName $col -NotePropertyValue '' -Force
                }
            }
        }

        $existingDict = @{}
        foreach ($item in $existingData) {
            $nsn = $item.'Part (NSN)'
            if (-not [string]::IsNullOrEmpty($nsn)) { $existingDict[$nsn] = $item }
        }

        $incoming = @{}
        foreach ($item in $parsedData) {
            if (-not [string]::IsNullOrEmpty($item.'Part (NSN)')) { $incoming[$item.'Part (NSN)'] = $true }
        }

        $merged = New-Object System.Collections.ArrayList
        foreach ($existing in $existingData) { [void]$merged.Add($existing) }

        foreach ($newItem in $parsedData) {
            $nsn = $newItem.'Part (NSN)'
            if ([string]::IsNullOrEmpty($nsn)) { continue }

            if ($existingDict.ContainsKey($nsn)) {
                $existing = $existingDict[$nsn]
                $changed = $false

                if ("$($existing.QTY)" -ne "$($newItem.QTY)") {
                    $existing.QTY = $newItem.QTY; $changed = $true
                }
                if ("$($existing.Location)" -ne "$($newItem.Location)") {
                    $existing.Location = $newItem.Location; $changed = $true
                }
                if ("$($existing.'13 Period Usage')" -ne "$($newItem.'13 Period Usage')") {
                    $existing.'13 Period Usage' = $newItem.'13 Period Usage'
                }
                if ($changed) { $result.Updated++ }

                foreach ($oem in 'OEM 1','OEM 2','OEM 3') {
                    if ([string]::IsNullOrEmpty($existing.$oem) -and
                        -not [string]::IsNullOrEmpty($newItem.$oem)) {
                        $existing.$oem = $newItem.$oem
                    }
                }
            } else {
                [void]$merged.Add($newItem)
                $existingDict[$nsn] = $newItem
                $result.New++
            }
        }

        foreach ($existing in $existingData) {
            $nsn = $existing.'Part (NSN)'
            if (-not [string]::IsNullOrEmpty($nsn) -and -not $incoming.ContainsKey($nsn)) {
                if ("$($existing.QTY)" -ne '0') {
                    $existing.QTY = '0'
                    $existing.Location = 'Not in current inventory'
                    $result.Removed++
                }
            }
        }

        $merged | Export-Csv -Path $TargetCsvPath -NoTypeInformation
    } else {
        $parsedData | Export-Csv -Path $TargetCsvPath -NoTypeInformation
        $result.New = $parsedData.Count
    }

    Write-Log "Updated $SiteName : $($result.Updated) updated, $($result.New) new, $($result.Removed) removed"
    return $result
}

function Parse-HTMLToCSV {
    param(
        [string]$htmlFilePath,
        [string]$siteName
    )

    Write-Log "Processing the downloaded HTML file for $siteName..."

    $logPath = Join-Path $config.PartsRoomDirectory "error_log.txt"

    try {
        Write-Log "Reading HTML content from file $htmlFilePath"
        $htmlContent = Get-Content -Path $htmlFilePath -Raw -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($htmlContent)) {
            throw "HTML content is empty or null"
        }

        Write-Log "HTML content read successfully. Parsing content..."
        $parsedData = @()
        $rows = @($htmlContent -split '<TR CLASS="MAIN"')

        Write-Log "Number of rows found: $($rows.Count)"

        if ($rows.Count -le 1) {
            throw "No data rows found in HTML content"
        }

        for ($i = 1; $i -lt $rows.Count; $i++) {
            $row = $rows[$i]
            Write-Log "Processing row ${i}"

            if ($null -eq $row) {
                Write-Log "Row ${i} is null, skipping"
                continue
            }

            $cells = @($row -split '<TD')
            Write-Log "Number of cells in row ${i}: $($cells.Count)"

            if ($cells.Count -ge 7) {
                try {
                    $partNSN = if ($cells[1]) { ($cells[1] -replace '>|</TD>').Trim() } else { "" }
                    $description = if ($cells[2]) { ($cells[2] -replace '>|</TD>').Trim() } else { "" }
                    $qty = if ($cells[3]) { ($cells[3] -replace '>|</TD>|style="text-align:right;"').Trim() } else { "0" }
                    $usage = if ($cells[4]) { ($cells[4] -replace '>|</TD>|style="text-align:right;"').Trim() } else { "0" }

                    $oemData = if ($cells[5]) { $cells[5] -replace '<DIV>|</DIV>|<SPAN.*?>|</SPAN>' } else { "" }
                    $oems = @($oemData -split 'OEM:' | Select-Object -Skip 1)
                    $oem1 = if ($oems.Count -gt 0 -and $oems[0]) { ($oems[0] -split ' ', 2)[1].Trim() -replace '</TD>' } else { "" }
                    $oem2 = if ($oems.Count -gt 1 -and $oems[1]) { ($oems[1] -split ' ', 2)[1].Trim() -replace '</TD>' } else { "" }
                    $oem3 = if ($oems.Count -gt 2 -and $oems[2]) { ($oems[2] -split ' ', 2)[1].Trim() -replace '</TD>' } else { "" }

                    $location = if ($cells[6]) { ($cells[6] -replace '>|</TD>|</TR>').Trim() } else { "" }

                    $parsedData += [PSCustomObject]@{
                        "Part (NSN)" = $partNSN
                        "Description" = $description
                        "QTY" = [int]($qty -replace '[^\d]')
                        "13 Period Usage" = [int]($usage -replace '[^\d]')
                        "Location" = $location
                        "OEM 1" = $oem1
                        "OEM 2" = $oem2
                        "OEM 3" = $oem3
                    }

                    Write-Log "Added row ${i}: Part(NSN)=$partNSN, Description=$description, QTY=$qty, Location=$location"

                }
                catch {
                    Write-Log "Error processing row ${i}: $($_.Exception.Message)"
                }
            }
            else {
                Write-Log "Row ${i} does not have enough cells, skipping"
            }
        }

        Write-Log "Number of parsed data entries: $($parsedData.Count)"

        if ($parsedData.Count -eq 0) {
            throw "No data parsed from HTML content"
        }

        return $parsedData
    }
    catch {
        Write-Log "Error: $($_.Exception.Message)"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("An error occurred while parsing the HTML for $siteName. Please check the log for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

function Update-AllFiles {
    Write-Log "=== Starting Update-AllFiles ==="

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Update Files"
    $progressForm.Size = New-Object System.Drawing.Size(520, 220)
    $progressForm.StartPosition = 'CenterScreen'
    $progressForm.FormBorderStyle = 'FixedDialog'
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 10)
    $progressLabel.Size = New-Object System.Drawing.Size(490, 40)
    $progressLabel.Text = "Starting..."
    $progressForm.Controls.Add($progressLabel)

	$progressBar = New-Object ModernProgressBar
	$progressBar.Location = New-Object System.Drawing.Point(10, 55)
	$progressBar.Size = New-Object System.Drawing.Size(490, 26)
	$progressForm.Controls.Add($progressBar)

    $detailLabel = New-Object System.Windows.Forms.Label
    $detailLabel.Location = New-Object System.Drawing.Point(10, 85)
    $detailLabel.Size = New-Object System.Drawing.Size(490, 90)
    $detailLabel.Text = ""
    $progressForm.Controls.Add($detailLabel)

    $progressForm.Show()
    $progressForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    $summary = New-Object System.Collections.ArrayList

    try {
        # ===== Phase 1: Same Day Parts Rooms =====
        $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
        if (-not (Test-Path $sameDayDir)) {
            New-Item -ItemType Directory -Path $sameDayDir -Force | Out-Null
        }

        $sameDaySites = @()
        if ($config.SameDayPartsRooms) { $sameDaySites = @($config.SameDayPartsRooms) }

        $progressBar.Maximum = $sameDaySites.Count + 2
        $progressBar.Value = 0

        $i = 0
        foreach ($site in $sameDaySites) {
            $i++
            $progressBar.Value = $i
            $progressLabel.Text = "Same Day Parts Room $i of $($sameDaySites.Count): $($site.FullName)"
            $progressForm.Refresh()
            [System.Windows.Forms.Application]::DoEvents()

            $csvPath = Join-Path $sameDayDir "$($site.FullName).csv"
            $r = Update-SinglePartsRoomFromUrl -SiteID $site.SiteID -SiteName $site.FullName -TargetCsvPath $csvPath
            if ($r) {
                [void]$summary.Add("Same Day - $($site.FullName): $($r.Updated) updated, $($r.New) new, $($r.Removed) removed")
            } else {
                [void]$summary.Add("Same Day - $($site.FullName): FAILED (see log)")
            }
        }

        # ===== Phase 2: Local Parts Room =====
        $progressBar.Value = $sameDaySites.Count + 1
        $progressLabel.Text = "Local Parts Room..."
        $progressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        $localCsvFiles = @(Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.csv" -File -ErrorAction SilentlyContinue)
        $localCsvPath = $null
        $localSiteName = $null
        $localSiteID = $null

        if ($localCsvFiles.Count -eq 1) {
            $localCsvPath = $localCsvFiles[0].FullName
            $localSiteName = [System.IO.Path]::GetFileNameWithoutExtension($localCsvFiles[0].Name)

            $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
            if (Test-Path $sitesPath) {
                $sites = @(Import-Csv -Path $sitesPath)
                if ($sites.Count -gt 0) {
                    $SiteIDColumn = if ($sites[0].PSObject.Properties.Name -contains 'Site ID') { 'Site ID' } else { $sites[0].PSObject.Properties.Name[0] }
                    $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains 'Full Name') { 'Full Name' } else { $sites[0].PSObject.Properties.Name[1] }
                    $match = $sites | Where-Object { $_.$fullNameColumn -eq $localSiteName } | Select-Object -First 1
                    if ($match) { $localSiteID = $match.$SiteIDColumn }
                }
            }

            if ($localSiteID) {
                $r = Update-SinglePartsRoomFromUrl -SiteID $localSiteID -SiteName $localSiteName -TargetCsvPath $localCsvPath
                if ($r) {
                    [void]$summary.Add("Local Parts Room - ${localSiteName}: $($r.Updated) updated, $($r.New) new, $($r.Removed) removed")
                } else {
                    [void]$summary.Add("Local Parts Room - ${localSiteName}: FAILED (see log)")
                }
            } else {
                [void]$summary.Add("Local Parts Room - ${localSiteName}: Site ID not found in Sites.csv, skipped")
            }
        } else {
            [void]$summary.Add("Local Parts Room: expected 1 CSV, found $($localCsvFiles.Count) - skipped")
        }

        # ===== Phase 3: Parts Books =====
        $progressBar.Value = $sameDaySites.Count + 2
        $progressLabel.Text = "Updating Parts Books..."
        $progressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        if ($localCsvPath -and (Test-Path $localCsvPath)) {
            $progressForm.Hide()
            try {
                Update-PartsBooks -sourceCSVPath $localCsvPath | Out-Null
                [void]$summary.Add("Parts Books: updated from $localSiteName data")
            } catch {
                Write-Log "Update-PartsBooks failed: $($_.Exception.Message)"
                [void]$summary.Add("Parts Books: FAILED - $($_.Exception.Message)")
            } finally {
                $progressForm.Show()
            }
        } else {
            [void]$summary.Add("Parts Books: skipped (no local parts room CSV)")
        }

        $progressLabel.Text = "Complete."
        $detailLabel.Text = ($summary -join "`r`n")
        $progressForm.Refresh()
        Start-Sleep -Seconds 2

        $fullSummary = "Update Files complete:`r`n`r`n" + ($summary -join "`r`n")
        [System.Windows.Forms.MessageBox]::Show(
            $fullSummary,
            "Update Files Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Write-Log "Update-AllFiles error: $($_.Exception.Message)"
        Write-Log "Stack: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show(
            "Error: $($_.Exception.Message)",
            "Update Files Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $progressForm.Close()
    }

    Write-Log "=== Update-AllFiles complete ==="
}

################################################################################
#                          Parts Management                                    #
################################################################################

# Function to take a part out
function Take-PartOut {
    Write-Log "Taking a part out..."
    [System.Windows.Forms.MessageBox]::Show("Take Part Out process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Simple helper function to get quantity
function Get-PartQuantity {
    param(
        [string]$PartNumber,
        [int]$MaxQty
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Quantity for $PartNumber"
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = "CenterParent"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Quantity (Max: $MaxQty):"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = "1"
    $textBox.Location = New-Object System.Drawing.Point(160, 20)
    $textBox.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(60, 70)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(150, 70)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $form.Controls.Add($cancelButton)
    
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $qty = 0
        if ([int]::TryParse($textBox.Text, [ref]$qty)) {
            if ($qty -gt 0 -and $qty -le $MaxQty) {
                return $qty
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid quantity between 1 and $MaxQty", "Invalid Quantity", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return 0
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid number", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return 0
        }
    }
    
    return 0
}

################################################################################
#                     External Systems Integration                             #
################################################################################

# Function to request a part order
function Request-PartOrder {
    Write-Log "Requesting a part order..."
    [System.Windows.Forms.MessageBox]::Show("Part Order Request process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to request a work order
function Request-WorkOrder {
    Write-Log "Requesting a work order..."
    [System.Windows.Forms.MessageBox]::Show("Work Order Request process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to make an MTSC ticket
function Make-MTSCTicket {
    Write-Log "Navigating to MTSC login page..."
    [System.Windows.Forms.MessageBox]::Show("Navigating to MTSC login page.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    # Open the MTSC ticket webpage in the default browser
    Start-Process "https://tickets.mtsc.usps.gov/login.php"
}

# Function to search the knowledge base
function Search-KnowledgeBase {
    Write-Log "Searching knowledge base..."
    $searchForm = New-Object System.Windows.Forms.Form
    $searchForm.Text = "Search Knowledge Base"
    $searchForm.Size = New-Object System.Drawing.Size(300, 150)
    $searchForm.StartPosition = 'CenterScreen'

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 20)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $searchForm.Controls.Add($textBox)

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(100, 60)
    $searchButton.Size = New-Object System.Drawing.Size(75, 23)
    $searchButton.Text = "Search"
    $searchButton.Add_Click({
        $searchTerms = $textBox.Text
        if (-not [string]::IsNullOrWhiteSpace($searchTerms)) {
            $baseUrl = "https://mtscprod.servicenowservices.com/kb?id=kb_search&query="
            $formattedTerms = $searchTerms -replace ' ', '%20'
            $spaceCount = ($searchTerms -split ' ').Count - 1
            $fullUrl = "${baseUrl}${formattedTerms}&spa=${spaceCount}"
            Start-Process $fullUrl
            Write-Log "Searched Knowledge Base for: $searchTerms"
        }
        $searchForm.Close()
    })
    $searchForm.Controls.Add($searchButton)

    $searchForm.ShowDialog()
}

################################################################################
#                       UI Setup and Components                                #
################################################################################

function ConvertFrom-eDacWorksheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$TargetDate
    )

    if (-not (Test-Path $HtmlPath)) { throw "eDAC HTML not found: $HtmlPath" }

    $htmlDoc = $null
    try {
        $htmlDoc = New-Object -ComObject "HTMLFile"
        $content = Get-Content -Path $HtmlPath -Raw
        try   { $htmlDoc.IHTMLDocument2_write($content) }
        catch {
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($content)
            $htmlDoc.write($bytes)
        }

        $rows = $htmlDoc.getElementsByTagName("tr")
        $edacList = @()

        $inTargetSection = $false
        $lastWasDataRow  = $false

        for ($i = 0; $i -lt $rows.length; $i++) {
            $row  = $rows.item($i)
            $text = ($row.innerText -replace '\s+', ' ').Trim()
            if ([string]::IsNullOrWhiteSpace($text)) { $lastWasDataRow = $false; continue }

            if ($text -match 'Scheduled Date\s*:') {
                $inTargetSection = $text.Contains($TargetDate)
                $lastWasDataRow  = $false
                continue
            }
            if (-not $inTargetSection) { continue }

            if ($text -match 'PM Description\s*:\s*(.+)') {
                if ($lastWasDataRow -and $edacList.Count -gt 0) {
                    $edacList[$edacList.Count - 1].PmDescription = $matches[1].Trim()
                }
                $lastWasDataRow = $false
                continue
            }

            $cells = $row.getElementsByTagName("td")
            if ($null -eq $cells -or $cells.length -lt 10) { $lastWasDataRow = $false; continue }

            $rowType = $cells.item(0).innerText.Trim()
            if ($rowType -notin @('W','A','B')) { $lastWasDataRow = $false; continue }

            $acronymCl = $cells.item(3).innerText.Trim()
            $acronym   = $acronymCl
            $classCode = ''
            if ($acronymCl -match '^([^-]+)-(.+)$') {
                $acronym   = $matches[1].Trim()
                $classCode = $matches[2].Trim()
            }

            $estText = $cells.item(10).innerText.Trim()
            $estHrs  = 0.0
            if ($estText -match '[\d.]+') { $estHrs = [double]$matches[0] }

            $desc    = $cells.item(9).innerText.Trim()
            $checkNo = ''
            if ($desc -match 'Checklist No\.:\s*(\d+)') { $checkNo = $matches[1] }

            $edacList += [PSCustomObject]@{
                RowType       = $rowType
                WorkOrderNo   = $cells.item(1).innerText.Trim()
                WorkCode      = $cells.item(2).innerText.Trim()
                Acronym       = $acronym
                ClassCode     = $classCode
                Equipment     = $cells.item(4).innerText.Trim()
                Route         = $cells.item(5).innerText.Trim()
                Frequency     = $cells.item(6).innerText.Trim()
                Priority      = $cells.item(7).innerText.Trim()
                DueDate       = $cells.item(8).innerText.Trim()
                Description   = $desc
                ChecklistNo   = $checkNo
                EstimatedHrs  = $estHrs
                PmDescription = ''
            }
            $lastWasDataRow = $true
        }

        return $edacList
    }
    finally {
        if ($null -ne $htmlDoc) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-SafeCellText {
    param($Cells, [int]$Index)
    if ($null -eq $Cells)         { return '' }
    if ($Index -ge $Cells.length) { return '' }
    $cell = $Cells.item($Index)
    if ($null -eq $cell)          { return '' }
    $t = $cell.innerText
    if ($null -eq $t)             { return '' }
    return $t.Trim()
}

function Get-SafeCellHtml {
    param($Cells, [int]$Index)
    if ($null -eq $Cells)         { return '' }
    if ($Index -ge $Cells.length) { return '' }
    $cell = $Cells.item($Index)
    if ($null -eq $cell)          { return '' }
    $h = $cell.innerHTML
    if ($null -eq $h)             { return '' }
    return $h
}

function ConvertFrom-PmChecklist {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HtmlPath)

    if (-not (Test-Path $HtmlPath)) { throw "PM Checklist HTML not found: $HtmlPath" }

    $htmlDoc = $null
    try {
        $htmlDoc = New-Object -ComObject "HTMLFile"
        $content = Get-Content -Path $HtmlPath -Raw
        try   { $htmlDoc.IHTMLDocument2_write($content) }
        catch {
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($content)
            $htmlDoc.write($bytes)
        }

        $allRows   = $htmlDoc.getElementsByTagName("tr")
        $allTables = $htmlDoc.getElementsByTagName("table")

        # ---------- Identification ----------
        $acronym       = ''
        $workCode      = ''
        $classCode     = ''
        $idNumber      = ''
        $type          = ''
        $equipmentNo   = ''
        $siteName      = ''
        $checklistDate = ''
        $checklistNo   = ''
        $skills        = ''
        $generatedOn   = ''

        for ($i = 0; $i -lt $allRows.length; $i++) {
            $tr   = $allRows.item($i)
            $text = ($tr.innerText -replace '\s+', ' ').Trim()

            if ($text -match 'Equipment Acronym') {
                $next  = $allRows.item($i + 1)
                $cells = $next.getElementsByTagName("th")
                if ($cells.length -eq 0) { $cells = $next.getElementsByTagName("td") }
                if ($cells.length -ge 5) {
                    $workCode  = $cells.item(0).innerText.Trim()
                    $acronym   = $cells.item(1).innerText.Trim()
                    $classCode = $cells.item(2).innerText.Trim()
                    $idNumber  = $cells.item(3).innerText.Trim()
                    $type      = $cells.item(4).innerText.Trim()
                }
            }

            if ($text -match 'Site Name') {
                $next  = $allRows.item($i + 1)
                $cells = $next.getElementsByTagName("th")
                if ($cells.length -eq 0) { $cells = $next.getElementsByTagName("td") }
                if ($cells.length -ge 3) {
                    $equipmentNo   = $cells.item(0).innerText.Trim()
                    $siteName      = $cells.item(1).innerText.Trim()
                    $checklistDate = $cells.item(2).innerText.Trim()
                }
            }

            if ($text -match 'Checklist No\.\s*:\s*(\d+)') { $checklistNo = $matches[1] }
            if ($text -match 'Skill\(s\)\s*:\s*([^,]+)')  { $skills = $matches[1].Trim() }
            if ($text -match 'Time\s*:\s*([^)]+?)\s*\)')  { $generatedOn = $matches[1].Trim() }
        }

        # ---------- Totals ----------
        $estOpen     = 0.0
        $estDueToday = 0.0
        for ($i = 0; $i -lt $allRows.length; $i++) {
            $t = $allRows.item($i).innerText
            if ($t -match 'for All Open Tasks:\s*([\d.]+)')  { $estOpen     = [double]$matches[1] }
            if ($t -match 'for Tasks Due Today:\s*([\d.]+)') { $estDueToday = [double]$matches[1] }
        }

        # ---------- Locate task tables ----------
        $pmTaskTables = @()
        for ($t = 0; $t -lt $allTables.length; $t++) {
            $table = $allTables.item($t)
            $trows = $table.getElementsByTagName("tr")
            if ($trows.length -lt 2) { continue }
            $headerText = ($trows.item(0).innerText -replace '\s+', ' ')
            if ($headerText -match 'Item No' -and $headerText -match 'Task Statement and Instruction') {
                $pmTaskTables += $table
            }
        }

        

        # ---------- Parse task rows ----------
        $pmTasks = @()
        $tableIdx = 0
        foreach ($table in $pmTaskTables) {
            if ($tableIdx -eq 0) { $powerState = 'Power Off' } else { $powerState = 'Power On' }
            $tableIdx++

            $trows = $table.getElementsByTagName("tr")

            for ($r = 1; $r -lt $trows.length; $r++) {
                $cells = $trows.item($r).getElementsByTagName("td")
                if ($null -eq $cells -or $cells.length -lt 7) { continue }

                $itemNo = Get-SafeCellText -Cells $cells -Index 0
                if ([string]::IsNullOrWhiteSpace($itemNo)) { continue }

                $componentText = Get-SafeCellText -Cells $cells -Index 1
                $estText       = Get-SafeCellText -Cells $cells -Index 2
                $skillText     = Get-SafeCellText -Cells $cells -Index 3
                $bypassText    = Get-SafeCellText -Cells $cells -Index 4
                $dueDateText   = Get-SafeCellText -Cells $cells -Index 5

                $estMin = 0
                if ($estText -match '[\d.]+') { $estMin = [int][double]$matches[0] }

                $rawHtml = Get-SafeCellHtml -Cells $cells -Index 6
                $cleanText = $rawHtml -replace '(?i)<br\s*/?>', "`n"
                $cleanText = $cleanText -replace '(?i)</(p|li|ol|ul|div|h\d)>', "`n"
                $cleanText = $cleanText -replace '(?s)<[^>]+>', ''
                $cleanText = [System.Net.WebUtility]::HtmlDecode($cleanText)
                $cleanText = $cleanText -replace '[ \t]+', ' '
                $cleanText = ($cleanText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join "`n"

                $taskObject = [PSCustomObject]@{
                    PowerState      = $powerState
                    ItemNo          = $itemNo
                    Component       = $componentText
                    EstTimeMin      = $estMin
                    MinSkillLevel   = $skillText
                    BypassCode      = $bypassText
                    DueDate         = $dueDateText
                    InstructionText = $cleanText
                    InstructionHtml = $rawHtml
                }

                $pmTasks += $taskObject
            }
        }

        return [PSCustomObject]@{
            Meta = [PSCustomObject]@{
                ChecklistNo      = $checklistNo
                Skills           = $skills
                GeneratedOn      = $generatedOn
                WorkCode         = $workCode
                Acronym          = $acronym
                ClassCode        = $classCode
                IdNumber         = $idNumber
                Type             = $type
                EquipmentNo      = $equipmentNo
                SiteName         = $siteName
                ChecklistDate    = $checklistDate
                EstOpenHours     = $estOpen
                EstDueTodayHours = $estDueToday
            }
            Tasks = $pmTasks
        }
    }
    finally {
        if ($null -ne $htmlDoc) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

# ============================================================
# Work Tracking — data layer
# ============================================================

function Format-MachineId {
    # Canonicalizes "dbcs-49", "DBCS_0049", "DBCS 49" -> "DBCS 49".
    # Bare digits "49" -> "MNO 49".  Returns $null on unparseable input.
    param([string]$Raw)

    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    $t = $Raw.Trim()

    if ($t -match '^([A-Za-z]+)[\s\-_]*0*(\d+)$') {
        $acr = $matches[1].ToUpper()
        $mno = [int]$matches[2]
        return [PSCustomObject]@{
            Acronym   = $acr
            Mno       = $mno.ToString()
            Canonical = "$acr $mno"
            IsNamed   = $true
        }
    }

    if ($t -match '^0*(\d+)$') {
        $mno = [int]$matches[1]
        return [PSCustomObject]@{
            Acronym   = ''
            Mno       = $mno.ToString()
            Canonical = "MNO $mno"
            IsNamed   = $false
        }
    }

    return $null
}

function Register-MachineIfNew {
    # Appends <Acronym,Mno> to Machines.csv if that exact pair isn't already
    # present.  Same MNO with a *different* acronym is allowed — e.g.
    # ASD 47 and DBCS 47 can both exist.
    param(
        [string]$Acronym,
        [string]$Mno
    )

    if ([string]::IsNullOrWhiteSpace($Acronym) -or [string]::IsNullOrWhiteSpace($Mno)) { return }

    $mnoInt = 0
    if (-not [int]::TryParse($Mno, [ref]$mnoInt)) { return }

    $path = Join-Path $script:config.DropdownCsvsDirectory "Machines.csv"
    $existing = @()
    if (Test-Path $path) {
        try { $existing = @(Import-Csv -Path $path) } catch { $existing = @() }
    } else {
        "Machine Acronym,Machine Number" | Out-File -FilePath $path -Encoding utf8
    }

    $acrUpper = $Acronym.Trim().ToUpper()

    foreach ($row in $existing) {
        $rowMno = 0
        if (-not [int]::TryParse($row.'Machine Number', [ref]$rowMno)) { continue }
        if ($rowMno -ne $mnoInt) { continue }

        $rowAcr = "$($row.'Machine Acronym')".Trim().ToUpper()
        if ($rowAcr -eq $acrUpper) {
            return   # exact (Acronym, MNO) pair already registered
        }
        # else: same MNO, different acronym — fall through and add
    }

    "$acrUpper,$mnoInt" | Out-File -FilePath $path -Append -Encoding utf8
    Write-Log "Registered new machine: $acrUpper $mnoInt"
}

function Get-JobsFilePath {
    $rel = $script:config.WorkTracking.JobsFile
    if ([string]::IsNullOrWhiteSpace($rel)) { throw "WorkTracking.JobsFile is not set in config." }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    return Join-Path $script:config.RootDirectory $rel
}

function Get-Jobs {
    $path = Get-JobsFilePath
    if (-not (Test-Path $path)) { return ,@() }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return ,@() }

        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return ,@() }

        # Flatten any nested arrays from earlier bad saves.
        $flat = New-Object System.Collections.ArrayList
        $stack = New-Object System.Collections.Stack
        $stack.Push($data)
        while ($stack.Count -gt 0) {
            $item = $stack.Pop()
            if ($null -eq $item) { continue }
            if ($item -is [System.Array]) {
                for ($i = $item.Count - 1; $i -ge 0; $i--) { $stack.Push($item[$i]) }
            } else {
                [void]$flat.Add($item)
            }
        }

        Write-Log "Get-Jobs: read $($flat.Count) job(s) from $path"
        return ,$flat.ToArray()
    } catch {
        Write-Log "Error reading jobs: $($_.Exception.Message)"
        return ,@()
    }
}

function Save-Jobs {
    param([array]$Jobs)

    $path = Get-JobsFilePath
    $dir  = Split-Path -Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $tmp = "$path.tmp"
    try {
        # Flatten input — guarantees we never write a nested array
        $flat = New-Object System.Collections.ArrayList
        if ($null -ne $Jobs) {
            $stack = New-Object System.Collections.Stack
            $stack.Push($Jobs)
            while ($stack.Count -gt 0) {
                $item = $stack.Pop()
                if ($null -eq $item) { continue }
                if ($item -is [System.Array]) {
                    for ($i = $item.Count - 1; $i -ge 0; $i--) { $stack.Push($item[$i]) }
                } else {
                    [void]$flat.Add($item)
                }
            }
        }

        $arr = $flat.ToArray()
        Write-Log "Save-Jobs: writing $($arr.Count) job(s) to $path"

        $json = if ($arr.Count -eq 0) { "[]" } else {
            $j = ConvertTo-Json -InputObject $arr -Depth 12
            if (-not $j.TrimStart().StartsWith('[')) { $j = "[$j]" }
            $j
        }

        $json | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $path -Force
        return $true
    } catch {
        Write-Log "Error saving jobs: $($_.Exception.Message)"
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        return $false
    }
}

# ============================================================
# Installation management
# ============================================================

$script:installSubdirs = @(
    'Scripts'
    'Dropdown CSVs'
    'Parts Room'
    'Parts Books'
    'Work Tracking'
    'Work Tracking\Weekly Worksheets'
    'Work Tracking\PM Checklists'
    'Work Tracking\Reports'
)

$script:installCsvFiles = @(
    'Sites.csv'
    'Parsed-Parts-Volumes.csv'
    'Machines.csv'
)

$script:installCsvHeaders = @{
    'Sites.csv'                  = 'Site ID,Full Name'
    'Parsed-Parts-Volumes.csv'   = 'Full Name,MS Book No,Volume'
    'Machines.csv'               = 'Machine Acronym,Machine Number'
}

function New-InstallationTree {
    # Creates the full folder tree under $Root. Returns the count created.
    param([string]$Root)

    if ([string]::IsNullOrWhiteSpace($Root)) { throw "Root is required." }
    if (-not (Test-Path $Root)) { New-Item -ItemType Directory -Path $Root -Force | Out-Null }

    $created = 0
    foreach ($sub in $script:installSubdirs) {
        $full = Join-Path $Root $sub
        if (-not (Test-Path $full)) {
            New-Item -ItemType Directory -Path $full -Force | Out-Null
            $created++
        }
    }

    # Same Day Parts Room lives under Parts Room
    $sameDay = Join-Path $Root 'Parts Room\Same Day Parts Room'
    if (-not (Test-Path $sameDay)) {
        New-Item -ItemType Directory -Path $sameDay -Force | Out-Null
        $created++
    }

    return $created
}

function Copy-DropdownTemplates {
    # Copies CSVs from $SourceDir into <Root>\Dropdown CSVs\, only if missing.
    # Creates header-only blanks for any not found in $SourceDir.
    # Returns @{ Copied = N; Seeded = N }
    param([string]$Root, [string]$SourceDir)

    $destDir = Join-Path $Root 'Dropdown CSVs'
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }

    $copied = 0
    $seeded = 0

    foreach ($name in $script:installCsvFiles) {
        $dest = Join-Path $destDir $name
        if (Test-Path $dest) { continue }   # don't overwrite

        $src = $null
        if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
            $candidate = Join-Path $SourceDir $name
            if (Test-Path $candidate) { $src = $candidate }
        }

        if ($src) {
            Copy-Item -Path $src -Destination $dest -Force
            $copied++
        } else {
            $hdr = $script:installCsvHeaders[$name]
            if ($hdr) { $hdr | Out-File -FilePath $dest -Encoding UTF8 }
            $seeded++
        }
    }

    return [PSCustomObject]@{ Copied = $copied; Seeded = $seeded }
}

function Write-DefaultConfigFile {
    param([string]$Root, [string]$ScriptsDir)

    if (-not (Test-Path $Root)) { New-Item -ItemType Directory -Path $Root -Force | Out-Null }

    $cfg  = New-DefaultConfigObject -Root $Root -ScriptsDir $ScriptsDir
    $path = Join-Path $Root 'Config.json'
    $cfg | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Show-InstallationWizard {
    # Guided setup. Returns $true if config was written, $false otherwise.
    param([string]$InitialRoot = '')

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Installation Setup'
    $form.Size = New-Object System.Drawing.Size(640, 480)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Installation Setup'
    $lblTitle.Location = New-Object System.Drawing.Point(20, 14)
    $lblTitle.Size = New-Object System.Drawing.Size(600, 26)
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblTitle)

    $lblIntro = New-Object System.Windows.Forms.Label
    $lblIntro.Text = "This will create (or verify) the folder tree and seed the dropdown CSVs.`r`nExisting data is never overwritten."
    $lblIntro.Location = New-Object System.Drawing.Point(20, 44)
    $lblIntro.Size = New-Object System.Drawing.Size(600, 40)
    $lblIntro.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $form.Controls.Add($lblIntro)

    # --- Root row ---
    $lblRoot = New-Object System.Windows.Forms.Label
    $lblRoot.Text = "Data root:"
    $lblRoot.Location = New-Object System.Drawing.Point(20, 100)
    $lblRoot.Size = New-Object System.Drawing.Size(80, 24)
    $lblRoot.TextAlign = 'MiddleRight'
    $lblRoot.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblRoot)

    $txtRoot = New-Object System.Windows.Forms.TextBox
    $txtRoot.Location = New-Object System.Drawing.Point(108, 100)
    $txtRoot.Size = New-Object System.Drawing.Size(400, 24)
    $txtRoot.Text = $InitialRoot
    $form.Controls.Add($txtRoot)

    $btnBrowseRoot = New-Object System.Windows.Forms.Button
    $btnBrowseRoot.Text = "Browse..."
    $btnBrowseRoot.Location = New-Object System.Drawing.Point(516, 100)
    $btnBrowseRoot.Size = New-Object System.Drawing.Size(90, 24)
    $btnBrowseRoot.FlatStyle = 'Flat'
    $btnBrowseRoot.FlatAppearance.BorderSize = 0
    $btnBrowseRoot.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnBrowseRoot.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnBrowseRoot.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBrowseRoot.Cursor = 'Hand'
    $btnBrowseRoot.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select or create the data root folder"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtRoot.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseRoot)

    # --- Source CSVs row (optional) ---
    $lblSrc = New-Object System.Windows.Forms.Label
    $lblSrc.Text = "CSV source (opt):"
    $lblSrc.Location = New-Object System.Drawing.Point(20, 136)
    $lblSrc.Size = New-Object System.Drawing.Size(80, 24)
    $lblSrc.TextAlign = 'MiddleRight'
    $lblSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblSrc)

    $txtSrc = New-Object System.Windows.Forms.TextBox
    $txtSrc.Location = New-Object System.Drawing.Point(108, 136)
    $txtSrc.Size = New-Object System.Drawing.Size(400, 24)
    $txtSrc.Text = ''
    $form.Controls.Add($txtSrc)

    $btnBrowseSrc = New-Object System.Windows.Forms.Button
    $btnBrowseSrc.Text = "Browse..."
    $btnBrowseSrc.Location = New-Object System.Drawing.Point(516, 136)
    $btnBrowseSrc.Size = New-Object System.Drawing.Size(90, 24)
    $btnBrowseSrc.FlatStyle = 'Flat'
    $btnBrowseSrc.FlatAppearance.BorderSize = 0
    $btnBrowseSrc.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnBrowseSrc.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnBrowseSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnBrowseSrc.Cursor = 'Hand'
    $btnBrowseSrc.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Folder containing Sites.csv, Causes.csv, etc. (optional)"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtSrc.Text = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseSrc)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "If left blank, blank CSVs with headers are created. You can fill them in later."
    $lblHint.Location = New-Object System.Drawing.Point(108, 162)
    $lblHint.Size = New-Object System.Drawing.Size(500, 20)
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(140,150,165)
    $form.Controls.Add($lblHint)

    # --- Summary ---
    $lblSum = New-Object System.Windows.Forms.Label
    $lblSum.Location = New-Object System.Drawing.Point(20, 200)
    $lblSum.Size = New-Object System.Drawing.Size(600, 170)
    $lblSum.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lblSum.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblSum.Text = ''
    $form.Controls.Add($lblSum)

    $updateSummary = {
        $root = $txtRoot.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($root)) {
            $lblSum.Text = "Select a data root to continue."
            return
        }
        $exists = Test-Path $root
        $files  = if ($exists) { @(Get-ChildItem -Path $root -File -Recurse -ErrorAction SilentlyContinue).Count } else { 0 }
        $dirs   = if ($exists) { @(Get-ChildItem -Path $root -Directory -Recurse -ErrorAction SilentlyContinue).Count } else { 0 }

        $src = $txtSrc.Text.Trim()
        $srcSummary = if ($src) { "CSV source: $src" } else { "CSV source: (none - blanks will be created)" }

        $lines = @()
        $lines += "Root:         $root"
        $lines += "Scripts dir:  $PSScriptRoot"
        $lines += "Status:       $(if ($exists) { "folder exists ($dirs subdirs, $files files)" } else { "will be created" })"
        $lines += $srcSummary
        $lines += ""
        $lines += "Will create:"
        foreach ($s in $script:installSubdirs) { $lines += "  $s" }
        $lines += "  Parts Room\Same Day Parts Room"
        $lblSum.Text = ($lines -join "`r`n")
    }

    $txtRoot.Add_TextChanged($updateSummary)
    $txtSrc.Add_TextChanged($updateSummary)
    & $updateSummary

    # --- Buttons ---
    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Size = New-Object System.Drawing.Size(100, 32)
    $cancelBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 220), 10)
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $false; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Create / Verify"
    $okBtn.Size = New-Object System.Drawing.Size(120, 32)
    $okBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 116), 10)
    $okBtn.Anchor = 'Top,Right'
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Add_Click({
        $root = $txtRoot.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($root)) {
            [System.Windows.Forms.MessageBox]::Show("Pick a data root first.", "Installation Setup", "OK", "Warning")
            return
        }

        try {
            $dirsCreated = New-InstallationTree -Root $root
            $csvResult   = Copy-DropdownTemplates -Root $root -SourceDir $txtSrc.Text.Trim()
            $configPath  = Write-DefaultConfigFile -Root $root -ScriptsDir $PSScriptRoot

            $msg = @()
            $msg += "Installation ready."
            $msg += ""
            $msg += "Root:            $root"
            $msg += "Folders created: $dirsCreated"
            $msg += "CSVs copied:     $($csvResult.Copied)"
            $msg += "CSVs seeded:     $($csvResult.Seeded)"
            $msg += "Config written:  $configPath"

            [System.Windows.Forms.MessageBox]::Show(($msg -join "`r`n"), "Installation Setup", "OK", "Information")

            Write-Log "Installation wizard: root=$root, dirs=$dirsCreated, csvCopied=$($csvResult.Copied), csvSeeded=$($csvResult.Seeded)"

            $form.Tag = $true
            $form.Close()
        } catch {
            Write-Log "Installation wizard error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Setup failed:`r`n$($_.Exception.Message)", "Installation Setup", "OK", "Error")
        }
    })
    $btnBar.Controls.Add($okBtn)

    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn
    $form.ShowDialog() | Out-Null
    return [bool]$form.Tag
}

# ============================================================
# Handbook Parts-Volumes creator
# ============================================================

function New-PartsVolumesCsv {
    param(
        [string]$Html,
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($Html))        { throw "No HTML supplied." }
    if ([string]::IsNullOrWhiteSpace($TargetPath))  { throw "No target path supplied." }

    $regex = '<option\s+value="(?<msbookno>[^"]+)"\s+volno="(?<volno>[^"]+)">(?<fullname>.*?)</option>'
    $matches = [regex]::Matches($Html, $regex)

    if ($matches.Count -eq 0) {
        throw "No <option> elements matched. Confirm you pasted the entire <select> block."
    }

    $dir = Split-Path -Path $TargetPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("MS Book No,Volume,Full Name")

    foreach ($m in $matches) {
        $msbookno = $m.Groups['msbookno'].Value
        $volno    = $m.Groups['volno'].Value
        $fullname = $m.Groups['fullname'].Value.Replace('"', '""')
        $lines.Add("$msbookno,$volno,""$fullname""")
    }

    $lines | Out-File -LiteralPath $TargetPath -Encoding UTF8
    Write-Log "New-PartsVolumesCsv: wrote $($matches.Count) row(s) to $TargetPath"
    return $matches.Count
}

function Show-PartsVolumesWizard {
    $targetPath = Join-Path $script:config.DropdownCsvsDirectory 'Parsed-Parts-Volumes.csv'

    if (Test-Path $targetPath) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Parsed-Parts-Volumes.csv already exists:`r`n$targetPath`r`n`r`nReplace it?",
            "Parts Volumes List",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    try {
        Start-Process "https://www1.mtsc.usps.gov/apps/mtsc/index.php#Doc&partssearch&0&NA"
    } catch {
        Write-Log "Parts-Volumes wizard: could not open browser: $($_.Exception.Message)"
    }

    [System.Windows.Forms.MessageBox]::Show(
        "In the browser window that just opened:`r`n`r`n" +
        "  1. Navigate to the Parts Search page.`r`n" +
        "  2. Right-click the parts-book dropdown and choose Inspect.`r`n" +
        "  3. Copy the entire <select> element's HTML.`r`n" +
        "  4. Paste it into the next window.",
        "Parts Volumes List",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Paste Parts Handbook HTML"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Paste the entire <select> block from the MTSC Parts Search page."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 28
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(680, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Parse & Save"
    $okBtn.Location = New-Object System.Drawing.Point(780, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.Anchor = 'Top,Right'
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPaste.Text)) { return }
        $form.Tag = $txtPaste.Text
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    $txtPaste = New-Object System.Windows.Forms.TextBox
    $txtPaste.Multiline = $true
    $txtPaste.MaxLength = 0
    $txtPaste.Dock = 'Fill'
    $txtPaste.ScrollBars = 'Both'
    $txtPaste.WordWrap = $false
    $txtPaste.Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtPaste.AcceptsReturn = $true
    $txtPaste.Padding = New-Object System.Windows.Forms.Padding(8)
    $form.Controls.Add($txtPaste)
    $txtPaste.BringToFront()

    $form.CancelButton = $cancelBtn
    $form.ShowDialog() | Out-Null

    $html = $form.Tag
    if ([string]::IsNullOrWhiteSpace($html)) { return $false }

    try {
        $count = New-PartsVolumesCsv -Html $html -TargetPath $targetPath
        [System.Windows.Forms.MessageBox]::Show(
            "Wrote $count entries to:`r`n$targetPath",
            "Parts Volumes List", "OK", "Information")
        return $true
    } catch {
        Write-Log "Show-PartsVolumesWizard: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed:`r`n$($_.Exception.Message)",
            "Parts Volumes List", "OK", "Error")
        return $false
    }
}

# ============================================================
# Sites list creator
# ============================================================

function New-SitesCsv {
    param(
        [string]$Html,
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($Html))       { throw "No HTML supplied." }
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { throw "No target path supplied." }

    # <option value="X" ...>Text [possibly with newlines] </option>
    # [\s\S]*? lets the display text span multiple lines; Singleline is
    # belt-and-braces in case the pattern is later edited to use .*?
    $regex = '<option\s+value="(?<siteid>[^"]+)"[^>]*>(?<fullname>[\s\S]*?)</option>'
    $optionMatches = [regex]::Matches($Html, $regex, 'IgnoreCase, Singleline')

    if ($optionMatches.Count -eq 0) {
        throw "No <option> elements matched. Confirm you pasted the entire <select> block."
    }

    $dir = Split-Path -Path $TargetPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Site ID,Full Name')

    $seen = @{}
    $written = 0
    foreach ($m in $optionMatches) {
        $siteId   = $m.Groups['siteid'].Value.Trim()
        $fullname = $m.Groups['fullname'].Value.Trim()

        if ([string]::IsNullOrWhiteSpace($siteId) -or [string]::IsNullOrWhiteSpace($fullname)) { continue }

        if ($seen.ContainsKey($siteId)) { continue }
        $seen[$siteId] = $true

        $safeName = $fullname.Replace('"', '""')
        $lines.Add("$siteId,""$safeName""")
        $written++
    }

    $lines | Out-File -LiteralPath $TargetPath -Encoding UTF8
    Write-Log "New-SitesCsv: wrote $written row(s) to $TargetPath"
    return $written
}

function Show-SitesWizard {
    $targetPath = Join-Path $script:config.DropdownCsvsDirectory 'Sites.csv'

    if (Test-Path $targetPath) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Sites.csv already exists:`r`n$targetPath`r`n`r`nReplace it?",
            "Sites List",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    try {
        Start-Process "http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_criteria"
    } catch {
        Write-Log "Sites wizard: could not open browser: $($_.Exception.Message)"
    }

    [System.Windows.Forms.MessageBox]::Show(
        "In the browser window that just opened:`r`n`r`n" +
        "  1. Navigate to the Site criteria page.`r`n" +
        "  2. Right-click the sites dropdown and choose Inspect.`r`n" +
        "  3. Copy the entire <select> element's HTML.`r`n" +
        "  4. Paste it into the next window.",
        "Sites List",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Paste Sites Dropdown HTML"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Paste the entire <select> block from the site criteria page."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 28
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(680, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Parse & Save"
    $okBtn.Location = New-Object System.Drawing.Point(780, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.Anchor = 'Top,Right'
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPaste.Text)) { return }
        $form.Tag = $txtPaste.Text
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    $txtPaste = New-Object System.Windows.Forms.TextBox
    $txtPaste.Multiline = $true
    $txtPaste.MaxLength = 0
    $txtPaste.Dock = 'Fill'
    $txtPaste.ScrollBars = 'Both'
    $txtPaste.WordWrap = $false
    $txtPaste.Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtPaste.AcceptsReturn = $true
    $txtPaste.Padding = New-Object System.Windows.Forms.Padding(8)
    $form.Controls.Add($txtPaste)
    $txtPaste.BringToFront()

    $form.CancelButton = $cancelBtn
    $form.ShowDialog() | Out-Null

    $html = $form.Tag
    if ([string]::IsNullOrWhiteSpace($html)) { return $false }

    try {
        $count = New-SitesCsv -Html $html -TargetPath $targetPath
        [System.Windows.Forms.MessageBox]::Show(
            "Wrote $count entries to:`r`n$targetPath",
            "Sites List", "OK", "Information")
        return $true
    } catch {
        Write-Log "Show-SitesWizard: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed:`r`n$($_.Exception.Message)",
            "Sites List", "OK", "Error")
        return $false
    }
}

# ============================================================
# Work Tracking — Historian (append-only JSONL)
# ============================================================

function Get-HistorianFilePath {
    $rel = $script:config.WorkTracking.HistorianFile
    if ([string]::IsNullOrWhiteSpace($rel)) { throw "WorkTracking.HistorianFile not set." }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    return Join-Path $script:config.RootDirectory $rel
}

function Add-HistorianEvent {
    # Appends one JSON object as a single line to Historian.jsonl.
    # Returns the event with .id and .capturedAt populated.
    param($Event)

    if ($null -eq $Event) { throw "Add-HistorianEvent: null event." }

    $path = Get-HistorianFilePath
    $dir  = Split-Path -Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    # Populate envelope fields if missing
    if (-not ($Event.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace("$($Event.id)")) {
        $Event | Add-Member -NotePropertyName 'id' -NotePropertyValue ([guid]::NewGuid().ToString()) -Force
    }
    if (-not ($Event.PSObject.Properties.Name -contains 'capturedAt') -or [string]::IsNullOrWhiteSpace("$($Event.capturedAt)")) {
        $Event | Add-Member -NotePropertyName 'capturedAt' -NotePropertyValue (Get-Date).ToString("o") -Force
    }
    if (-not ($Event.PSObject.Properties.Name -contains 'capturedBy')) {
        $tech = "$($script:config.WorkTracking.TechnicianName)".Trim()
        $Event | Add-Member -NotePropertyName 'capturedBy' -NotePropertyValue $tech -Force
    }

    # Serialize onto one line (ConvertTo-Json puts nested objects on multiple
    # lines, so we compact by removing all CR/LF pairs)
    $line = $Event | ConvertTo-Json -Depth 12 -Compress
    $line = $line -replace "[\r\n]+", ' '

    try {
        Add-Content -Path $path -Value $line -Encoding UTF8
        Write-Log "Historian: appended event kind=$($Event.kind) id=$($Event.id)"
        return $Event
    } catch {
        Write-Log "Historian append failed: $($_.Exception.Message)"
        throw
    }
}

function Send-JobToHistorian {
    # Returns $true on success, $false on validation failure.
    param($Job)

    if ($null -eq $Job) { return $false }

    $errs = Test-JobFields -Job $Job
    if (@($errs).Count -gt 0) {
        Write-Log "Send-JobToHistorian: '$($Job.id)' failed validation: $($errs -join '; ')"
        return $false
    }

    $hours = Get-JobHours $Job
    if ($null -eq $hours) { $hours = 0 }

    $label = if ($Job.kind -eq 'workorder') { 'Work order' } else { 'Reactive call' }
    $woLabel = if (-not [string]::IsNullOrWhiteSpace($Job.workOrderNo)) { " · WO $($Job.workOrderNo)" } else { '' }

    $detailLines = @()
    $detailLines += "Window: $($Job.startTime) → $($Job.endTime)  ($($hours.ToString('F2')) h)"
    if (-not [string]::IsNullOrWhiteSpace($Job.workOrderNo)) { $detailLines += "WO#:    $($Job.workOrderNo)" }
    if (-not [string]::IsNullOrWhiteSpace($Job.description)) { $detailLines += "Notes:  $($Job.description)" }
    if (@($Job.parts).Count -gt 0) {
        $detailLines += "Parts used:"
        foreach ($p in @($Job.parts)) {
            $detailLines += "  • $($p.PartNumber)  $($p.Description)  qty=$($p.Quantity)"
        }
    }

    $evt = [PSCustomObject]@{
        source     = 'work_tracking'
        kind       = $Job.kind
        machineId  = "$($Job.machineId)"
        eventTime  = $Job.createdAt
        summary    = "$label — $($Job.machineId) · $($hours.ToString('F2'))h$woLabel"
        detail     = ($detailLines -join "`n")
        payload    = [PSCustomObject]@{
            job = $Job
        }
    }

    try {
        $saved = Add-HistorianEvent -Event $evt
        return $saved.id
    } catch {
        Write-Log "Send-JobToHistorian: append failed for '$($Job.id)': $($_.Exception.Message)"
        return $false
    }
}

function Save-PmChecklistEvent {
    # Idempotent write for a pm.checklist event.  Any prior event with the
    # same (machineId, checklistNo, work date) is dropped from the file,
    # then this event is appended.  Other lines in Historian.jsonl are
    # preserved unchanged.
    param($Event)

    if ($null -eq $Event) { throw "Save-PmChecklistEvent: null event." }

    $path = Get-HistorianFilePath
    $dir  = Split-Path -Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    # Populate envelope fields. capturedAt is always refreshed so the
    # Historian shows when the current version was written.
    if (-not ($Event.PSObject.Properties.Name -contains 'id') -or
        [string]::IsNullOrWhiteSpace("$($Event.id)")) {
        $Event | Add-Member -NotePropertyName 'id' -NotePropertyValue ([guid]::NewGuid().ToString()) -Force
    }
    $Event | Add-Member -NotePropertyName 'capturedAt' -NotePropertyValue (Get-Date).ToString("o") -Force
    $tech = "$($script:config.WorkTracking.TechnicianName)".Trim()
    $Event | Add-Member -NotePropertyName 'capturedBy' -NotePropertyValue $tech -Force

    $machineId   = "$($Event.payload.machineId)"
    $checklistNo = "$($Event.payload.checklistNo)"
    $targetDate  = $null
    try { $targetDate = ([DateTime]$Event.payload.checklistDate).Date } catch { }

    $newLine = $Event | ConvertTo-Json -Depth 12 -Compress
    $newLine = $newLine -replace "[\r\n]+", ' '

    # Read existing lines; drop any prior pm.checklist for the same tuple.
    $kept = New-Object System.Collections.Generic.List[string]
    $dropped = 0
    if (Test-Path $path) {
        foreach ($line in (Get-Content -Path $path -Encoding UTF8)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $e = $null
            try { $e = $line | ConvertFrom-Json } catch { }
            if ($null -ne $e -and "$($e.kind)" -eq 'pm.checklist' -and $e.payload) {
                if ("$($e.machineId)" -eq $machineId -and
                    "$($e.payload.checklistNo)" -eq $checklistNo) {

                    $match = $true
                    if ($targetDate) {
                        $evDate = Get-HistorianEventWorkDate -Event $e
                        $match = ($evDate -and $evDate.Date -eq $targetDate)
                    }
                    if ($match) { $dropped++; continue }
                }
            }
            $kept.Add($line) | Out-Null
        }
    }

    $tmp = "$path.tmp"
    try {
        $kept.Add($newLine) | Out-Null
        $kept | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $path -Force
        Write-Log "Historian: saved pm.checklist machine=$machineId no=$checklistNo date=$($targetDate.ToString('yyyy-MM-dd')) (superseded $dropped prior)"
        return $Event
    } catch {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
}

function Send-PmChecklistToHistorian {
    # Writes one event per machine for the given date.  Any prior event for
    # the same machine + checklist number + date is replaced, so a second
    # send corrects the first rather than duplicating it.
    param([DateTime]$Date)

    $written = 0
    $blank   = 0

    foreach ($m in @($script:pmMachines)) {
        if (-not $m.SortDate -or $m.SortDate.Date -ne $Date.Date) { continue }
        if (-not $m.State) { continue }

        $selected = @()
        foreach ($t in @($m.Parsed.Tasks)) {
            $e = $m.State[$t.ItemNo]
            if (-not $e -or -not $e.Selected) { continue }
            $min = if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { [int]$t.EstTimeMin }
            $selected += [PSCustomObject]@{
                ItemNo        = $t.ItemNo
                Component     = $t.Component
                PowerState    = $t.PowerState
                DueDate       = $t.DueDate
                MinSkillLevel = $t.MinSkillLevel
                EstTimeMin    = $t.EstTimeMin
                ActualTimeMin = $min
                Override      = ($null -ne $e.CustomTimeMin)
            }
        }

        if ($selected.Count -eq 0) { $blank++; continue }

        $totalMin = ($selected | Measure-Object -Property ActualTimeMin -Sum).Sum
        $hours    = [Math]::Round($totalMin / 60.0, 2)

        $detailLines = @()
        $detailLines += "Checklist No: $($m.ChecklistNo)"
        $detailLines += "Date:         $($m.ChecklistDate)"
        $detailLines += "Est Open:     $($m.EstOpen) h"
        $detailLines += "Est Today:    $($m.EstDueToday) h"
        $detailLines += "Selected:     $($selected.Count) of $($m.TaskCount) tasks  ($hours h)"
        $detailLines += ""
        foreach ($s in $selected) {
            $flag = if ($s.Override) { '*' } else { ' ' }
            $detailLines += ("  {0} {1,-6} {2,-6} {3,4}min  {4}" -f $flag, $s.ItemNo, $s.PowerState, $s.ActualTimeMin, $s.Component)
        }

        $evt = [PSCustomObject]@{
            source    = 'work_tracking'
            kind      = 'pm.checklist'
            machineId = "$($m.MachineId)"
            eventTime = (Get-Date).ToString("o")
            summary   = "PM Checklist — $($m.MachineId) · $($selected.Count) task$(if ($selected.Count -ne 1) { 's' }) · $($hours.ToString('F2')) h"
            detail    = ($detailLines -join "`n")
            payload   = [PSCustomObject]@{
                machineId     = "$($m.MachineId)"
                checklistNo   = $m.ChecklistNo
                checklistDate = $m.ChecklistDate
                estOpenHours  = $m.EstOpen
                estDueToday   = $m.EstDueToday
                selectedTasks = $selected
                totalHours    = $hours
            }
        }

        try {
            Save-PmChecklistEvent -Event $evt | Out-Null
            $written++
        } catch {
            Write-Log "Send-PmChecklistToHistorian: write failed for $($m.MachineId): $($_.Exception.Message)"
        }
    }

    return [PSCustomObject]@{
        Written = $written
        Blank   = $blank
    }
}

function Send-WorksheetDayToHistorian {
    # Sends one event for the loaded worksheet day.
    # Skips rows still marked NeedsAudit and reports them in the summary.
    param([DateTime]$Date, $Worksheet)

    if ($null -eq $Worksheet -or $null -eq $Worksheet.Rows) { return $null }

    $rows = @($Worksheet.Rows)
    $audit = @($rows | Where-Object { $_.NeedsAudit })

    $sumEst = 0.0; $sumAct = 0.0
    foreach ($r in $rows) {
        $e = 0.0
        [void][double]::TryParse("$($r.EstimatedHrs)", [ref]$e)
        $sumEst += $e
        if ($null -ne $r.ActualTime) { $sumAct += [double]$r.ActualTime }
    }

    $detailLines = @()
    $detailLines += "Date:     $($Worksheet.Date)"
    $detailLines += "Rows:     $($rows.Count)"
    if ($audit.Count -gt 0) {
        $detailLines += "NeedsAudit: $($audit.Count) row(s) — NOT included as sent"
    }
    $detailLines += "Est:      $("{0:F2}" -f $sumEst) h"
    $detailLines += "Actual:   $("{0:F2}" -f $sumAct) h"
    $detailLines += ("Diff:     {0}{1:F2} h" -f $(if (($sumAct - $sumEst) -ge 0) { '+' } else { '' }), ($sumAct - $sumEst))
    $detailLines += ""
    foreach ($r in $rows) {
        $flag = if ($r.NeedsAudit) { 'AUDIT' } else { '     ' }
        $act  = if ($null -ne $r.ActualTime) { ("{0:F2}" -f [double]$r.ActualTime) } else { '  -- ' }
        $detailLines += ("  {0}  {1,-10}  {2,-5} {3,-4}  est {4,-5}  act {5}" -f `
            $flag, $r.WorkOrderNo, $r.Acronym, $r.Equipment, $r.EstimatedHrs, $act)
    }

    $evt = [PSCustomObject]@{
        source    = 'work_tracking'
        kind      = 'worksheet'
        machineId = 'WORKSHEET'
        eventTime = (Get-Date).ToString("o")
        summary   = "Worksheet — $($Date.ToString('yyyy-MM-dd')) · $($rows.Count) rows · est $("{0:F2}" -f $sumEst)h vs actual $("{0:F2}" -f $sumAct)h$(if ($audit.Count -gt 0) { " · $($audit.Count) needs audit" } else { '' })"
        detail    = ($detailLines -join "`n")
        payload   = [PSCustomObject]@{
            date       = $Worksheet.Date
            rowCount   = $rows.Count
            estHours   = [Math]::Round($sumEst, 2)
            actHours   = [Math]::Round($sumAct, 2)
            auditCount = $audit.Count
            rows       = $rows
        }
    }

    try {
        return Add-HistorianEvent -Event $evt
    } catch {
        Write-Log "Send-WorksheetDayToHistorian: $($_.Exception.Message)"
        return $null
    }
}

function Send-AllReadyJobs {
    # Returns a summary object with Sent/Skipped/Failed counts.

    $jobs = Get-Jobs
    if ($null -eq $jobs) { $jobs = @() }

    $sent = 0; $skipped = 0; $failed = 0
    $skippedReasons = @()

    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $job = $jobs[$i]
        if ($job.sentToHistorian) { continue }

        $errs = Test-JobFields -Job $job
        if (@($errs).Count -gt 0) {
            $skipped++
            $skippedReasons += "$($job.machineId) $($job.startTime): $($errs[0])"
            continue
        }

        $eventId = Send-JobToHistorian -Job $job
        if ($eventId) {
            $jobs[$i].sentToHistorian = $true
            $jobs[$i].sentAt          = (Get-Date).ToString("o")
            $jobs[$i] | Add-Member -NotePropertyName 'historianEventId' -NotePropertyValue $eventId -Force
            $sent++
        } else {
            $failed++
        }
    }

    if ($sent -gt 0) { Save-Jobs -Jobs $jobs | Out-Null }

    return [PSCustomObject]@{
        Sent    = $sent
        Skipped = $skipped
        Failed  = $failed
        Reasons = $skippedReasons
    }
}

function Get-HistorianEvents {
    # Reads the whole JSONL file, tolerates blank/garbled lines.
    # Returns an array of PSCustomObjects sorted newest first.
    $path = Get-HistorianFilePath
    if (-not (Test-Path $path)) { return ,@() }

    $events = @()
    $lineNo = 0
    foreach ($line in (Get-Content -Path $path -Encoding UTF8)) {
        $lineNo++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $events += ($line | ConvertFrom-Json)
        } catch {
            Write-Log "Historian: skipped unparseable line $lineNo"
        }
    }
    return ,@($events | Sort-Object -Property capturedAt -Descending)
}

function ConvertTo-FlatArray {
    param($Source)
    if ($null -eq $Source) { return ,@() }
    $out = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Stack
    $stack.Push($Source)
    while ($stack.Count -gt 0) {
        $x = $stack.Pop()
        if ($null -eq $x) { continue }
        if ($x -is [System.Array] -or $x -is [System.Collections.IList]) {
            for ($i = $x.Count - 1; $i -ge 0; $i--) { $stack.Push($x[$i]) }
        } else {
            [void]$out.Add($x)
        }
    }
    return ,$out.ToArray()
}

function Test-HistorianLayer {
    # Small standalone validation — run manually before trusting the layer.
    # Does NOT touch the real Historian.jsonl; uses a temp file.

    $savedPath = $script:config.WorkTracking.HistorianFile
    $tempFile  = Join-Path ([System.IO.Path]::GetTempPath()) ("hist_test_" + [guid]::NewGuid().ToString() + ".jsonl")
    try {
        $script:config.WorkTracking.HistorianFile = $tempFile

        $e1 = Add-HistorianEvent -Event ([PSCustomObject]@{
            kind = 'job'; machineId = 'DBCS 49'; eventTime = (Get-Date).ToString("o")
            summary = 'Reactive call — DBCS 49 · 0.50h'
        })
        $e2 = Add-HistorianEvent -Event ([PSCustomObject]@{
            kind = 'pm.checklist'; machineId = 'DBCS 49'; eventTime = (Get-Date).ToString("o")
            summary = 'PM Checklist — 3 tasks selected, 1.25h'
        })
        $e3 = Add-HistorianEvent -Event ([PSCustomObject]@{
            kind = 'worksheet'; machineId = 'WORKSHEET'; eventTime = (Get-Date).ToString("o")
            summary = 'Worksheet — 15 rows, est 8.20h vs actual 7.10h'
        })

        $all = Get-HistorianEvents
        $result = [PSCustomObject]@{
            AppendedCount = 3
            ReadCount     = @($all).Count
            Kinds         = @($all | ForEach-Object { $_.kind }) -join ', '
            HasIds        = (@($all | Where-Object { $_.id }) | Measure-Object).Count -eq 3
            HasCapturedAt = (@($all | Where-Object { $_.capturedAt }) | Measure-Object).Count -eq 3
            Path          = $tempFile
        }
        return $result
    } finally {
        $script:config.WorkTracking.HistorianFile = $savedPath
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
    }
}

# ============================================================
# Work Tracking — PM Checklist archive + import
# ============================================================

function Get-PmArchiveRoot {
    $rel = $script:config.WorkTracking.PmChecklistsDirectory
    if ([string]::IsNullOrWhiteSpace($rel)) { throw "WorkTracking.PmChecklistsDirectory not set." }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    return Join-Path $script:config.RootDirectory $rel
}

function Get-PmArchiveDir {
    # Flat archive: <Root>/Work Tracking/PM Checklists/  (no date subfolder)
    $root = Get-PmArchiveRoot
    if (-not (Test-Path $root)) {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
    }
    return $root
}

function Import-PmChecklist {
    # Parses pasted HTML, archives raw HTML + parsed JSON flat, returns
    # a summary object.  Throws on parse failure.
    param([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { throw "No HTML provided." }

    # Parse via a temp file (the existing parser is file-based)
    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("pm_import_" + [guid]::NewGuid().ToString() + ".html")
    try {
        [System.IO.File]::WriteAllText($temp, $Html, [System.Text.Encoding]::UTF8)
        $parsed = ConvertFrom-PmChecklist -HtmlPath $temp
    } finally {
        if (Test-Path $temp) { Remove-Item $temp -Force -ErrorAction SilentlyContinue }
    }

    # Determine the checklist date for the filename — prefer parsed, fall
    # back to today if the parse didn't yield a usable value.
    $dateStamp = (Get-Date).ToString("yyyy-MM-dd")
    if (-not [string]::IsNullOrWhiteSpace($parsed.Meta.ChecklistDate)) {
        try { $dateStamp = ([DateTime]$parsed.Meta.ChecklistDate).ToString("yyyy-MM-dd") } catch {}
    }

    $archiveDir = Get-PmArchiveDir   # flat

    # Safe filename components
    $acr = if ($parsed.Meta.Acronym)     { ($parsed.Meta.Acronym     -replace '[^\w]', '') } else { 'UNKNOWN' }
    $mno = if ($parsed.Meta.EquipmentNo) { ($parsed.Meta.EquipmentNo -replace '[^\w]', '') } else { '0' }
    $cn  = if ($parsed.Meta.ChecklistNo) { ($parsed.Meta.ChecklistNo -replace '[^\w]', '') } else { '0' }

    $baseName = "${acr}_${mno}_${cn}_${dateStamp}"

    $htmlPath = Join-Path $archiveDir "$baseName.html"
    $jsonPath = Join-Path $archiveDir "$baseName.json"

    [System.IO.File]::WriteAllText($htmlPath, $Html, [System.Text.Encoding]::UTF8)
    $parsed | ConvertTo-Json -Depth 12 | Set-Content -Path $jsonPath -Encoding UTF8

    Write-Log "Imported PM Checklist: $baseName -> $archiveDir"

    return [PSCustomObject]@{
        MachineId     = "$($parsed.Meta.Acronym) $($parsed.Meta.EquipmentNo)"
        Acronym       = $parsed.Meta.Acronym
        Mno           = $parsed.Meta.EquipmentNo
        ChecklistNo   = $parsed.Meta.ChecklistNo
        ChecklistDate = $parsed.Meta.ChecklistDate
        DateStamp     = $dateStamp
        GeneratedOn   = $parsed.Meta.GeneratedOn
        Skills        = $parsed.Meta.Skills
        EstOpen       = $parsed.Meta.EstOpenHours
        EstDueToday   = $parsed.Meta.EstDueTodayHours
        TaskCount     = @($parsed.Tasks).Count
        HtmlPath      = $htmlPath
        JsonPath      = $jsonPath
        Parsed        = $parsed
    }
}

function Load-PmMachines {
    # Scans the flat archive for all *.json, returns everything sorted by
    # checklist date desc (newest first).  Each entry carries a real
    # [DateTime] sortable field plus the raw Meta from the parsed file.
    $dir = Get-PmArchiveDir
    if (-not (Test-Path $dir)) { return ,@() }

    $machines = @()
    foreach ($jsonFile in @(Get-ChildItem -Path $dir -Filter "*.json" -File)) {
        try {
            $data = Get-Content -Path $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json

            $dt = [DateTime]::MinValue
            if (-not [string]::IsNullOrWhiteSpace($data.Meta.ChecklistDate)) {
                try { $dt = [DateTime]$data.Meta.ChecklistDate } catch {}
            }

            $machines += [PSCustomObject]@{
                MachineId     = "$($data.Meta.Acronym) $($data.Meta.EquipmentNo)"
                Acronym       = $data.Meta.Acronym
                Mno           = $data.Meta.EquipmentNo
                ChecklistNo   = $data.Meta.ChecklistNo
                ChecklistDate = $data.Meta.ChecklistDate
                DateStamp     = if ($dt -ne [DateTime]::MinValue) { $dt.ToString("yyyy-MM-dd") } else { '' }
                SortDate      = $dt
                GeneratedOn   = $data.Meta.GeneratedOn
                Skills        = $data.Meta.Skills
                EstOpen       = $data.Meta.EstOpenHours
                EstDueToday   = $data.Meta.EstDueTodayHours
                TaskCount     = @($data.Tasks).Count
                HtmlPath      = ($jsonFile.FullName -replace '\.json$', '.html')
                JsonPath      = $jsonFile.FullName
                Parsed        = $data
            }
        } catch {
            Write-Log "Error loading PM machine from $($jsonFile.Name): $($_.Exception.Message)"
        }
    }

    return ,@($machines | Sort-Object -Property SortDate -Descending)
}

function Get-PmSessionPath {
    param([string]$JsonPath)
    return ($JsonPath -replace '\.json$', '.session.json')
}

function Get-PmSession {
    param($Machine)
    $path = Get-PmSessionPath -JsonPath $Machine.JsonPath
    $empty = @{ Selections = @{} }
    if (-not (Test-Path $path)) { return $empty }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $empty }
        $data = $raw | ConvertFrom-Json
        $selections = @{}
        if ($data.Selections) {
            foreach ($prop in $data.Selections.PSObject.Properties) {
                $entry = $prop.Value
                $selections[$prop.Name] = @{
                    Selected      = [bool]$entry.Selected
                    CustomTimeMin = if ($null -ne $entry.CustomTimeMin) { [int]$entry.CustomTimeMin } else { $null }
                }
            }
        }
        return @{ Selections = $selections }
    } catch {
        Write-Log "Get-PmSession: read error: $($_.Exception.Message)"
        return $empty
    }
}

function Save-PmSession {
    param($Machine)
    if (-not $Machine -or -not $Machine.State) { return }
    $path = Get-PmSessionPath -JsonPath $Machine.JsonPath

    $selObj = [ordered]@{}
    foreach ($key in $Machine.State.Keys) {
        $e = $Machine.State[$key]
        $selObj[$key] = [PSCustomObject]@{
            Selected      = [bool]$e.Selected
            CustomTimeMin = if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { $null }
        }
    }

    $payload = [PSCustomObject]@{
        machineId     = $Machine.MachineId
        checklistDate = $Machine.ChecklistDate
        checklistNo   = $Machine.ChecklistNo
        updatedAt     = (Get-Date).ToString("o")
        selections    = $selObj
    }

    try { $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8 }
    catch { Write-Log "Save-PmSession: error: $($_.Exception.Message)" }
}

function Update-PmSelectedHoursTotal {
    # Counts selected tasks for machines whose checklist date == today.
    # Feeds the Work Budget panel.
    $today = (Get-Date).Date
    $total = 0.0
    foreach ($m in @($script:pmMachines)) {
        if (-not $m.State) { continue }
        if ($m.SortDate.Date -ne $today) { continue }
        foreach ($task in @($m.Parsed.Tasks)) {
            $e = $m.State[$task.ItemNo]
            if (-not $e -or -not $e.Selected) { continue }
            $min = if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { [int]$task.EstTimeMin }
            $total += ($min / 60.0)
        }
    }
    $script:pmTotalSelectedHours = [Math]::Round($total, 2)

    if (Get-Command Refresh-WorkBudget -ErrorAction SilentlyContinue) {
        try { Refresh-WorkBudget } catch { }
    }
    if ($script:pmGrandTotalLabel) {
        $count = 0
        foreach ($m in @($script:pmMachines)) {
            if (-not $m.State -or $m.SortDate.Date -ne $today) { continue }
            foreach ($t in @($m.Parsed.Tasks)) {
                $e = $m.State[$t.ItemNo]
                if ($e -and $e.Selected) { $count++ }
            }
        }
        $script:pmGrandTotalLabel.Text = "Today: $count task$(if ($count -ne 1) { 's' }) / $($script:pmTotalSelectedHours) h"
    }
}

function Show-PmTaskEditor {
    param($Machine, $Task)

    $state = $Machine.State[$Task.ItemNo]
    if (-not $state) { $state = @{ Selected = $false; CustomTimeMin = $null } }
    $defaultTime = if ($null -ne $state.CustomTimeMin) { [int]$state.CustomTimeMin } else { [int]$Task.EstTimeMin }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($Task.ItemNo) — $($Task.Component)"
    $form.Size = New-Object System.Drawing.Size(940, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 200), 10)
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Save"
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 104), 10)
    $okBtn.Anchor = 'Top,Right'
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Add_Click({
        $form.Tag = [PSCustomObject]@{
            Selected      = [bool]$chkDone.Checked
            CustomTimeMin = if ($chkOverride.Checked) { [int]$numTime.Value } else { $null }
        }
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    # Top: metadata block
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'
    $top.Height = 160
    $top.BackColor = [System.Drawing.Color]::White
    $top.Padding = New-Object System.Windows.Forms.Padding(14)
    $form.Controls.Add($top)

    $meta = New-Object System.Windows.Forms.Label
    $meta.Dock = 'Fill'
    $meta.Font = New-Object System.Drawing.Font("Consolas", 9)
    $meta.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $meta.Text = @(
        "Machine:      $($Machine.MachineId)"
        "Item No:      $($Task.ItemNo)"
        "Component:    $($Task.Component)"
        "Power State:  $($Task.PowerState)"
        "Min Skill:    $($Task.MinSkillLevel)"
        "Due Date:     $($Task.DueDate)"
        "Est. Time:    $($Task.EstTimeMin) min"
    ) -join "`r`n"
    $top.Controls.Add($meta)

    # Middle: edit controls
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Top'
    $mid.Height = 56
    $mid.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Controls.Add($mid)

    $chkDone = New-Object System.Windows.Forms.CheckBox
    $chkDone.Text = "Completed"
    $chkDone.Location = New-Object System.Drawing.Point(14, 16)
    $chkDone.Size = New-Object System.Drawing.Size(120, 24)
    $chkDone.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $chkDone.Checked = [bool]$state.Selected
    $mid.Controls.Add($chkDone)

    $chkOverride = New-Object System.Windows.Forms.CheckBox
    $chkOverride.Text = "Override est. time:"
    $chkOverride.Location = New-Object System.Drawing.Point(180, 16)
    $chkOverride.Size = New-Object System.Drawing.Size(160, 24)
    $chkOverride.Checked = ($null -ne $state.CustomTimeMin)
    $mid.Controls.Add($chkOverride)

    $numTime = New-Object System.Windows.Forms.NumericUpDown
    $numTime.Location = New-Object System.Drawing.Point(346, 17)
    $numTime.Size = New-Object System.Drawing.Size(80, 24)
    $numTime.Minimum = 0
    $numTime.Maximum = 9999
    $numTime.Value = [Math]::Min(9999, [Math]::Max(0, $defaultTime))
    $numTime.Enabled = $chkOverride.Checked
    $mid.Controls.Add($numTime)

    $lblMin = New-Object System.Windows.Forms.Label
    $lblMin.Text = "min"
    $lblMin.Location = New-Object System.Drawing.Point(432, 20)
    $lblMin.Size = New-Object System.Drawing.Size(30, 22)
    $mid.Controls.Add($lblMin)

    $chkOverride.Add_CheckedChanged({ $numTime.Enabled = $chkOverride.Checked })

    # Instructions
    $lblInst = New-Object System.Windows.Forms.Label
    $lblInst.Text = "Task Statement and Instruction"
    $lblInst.Dock = 'Top'
    $lblInst.Height = 24
    $lblInst.Padding = New-Object System.Windows.Forms.Padding(14, 4, 0, 0)
    $lblInst.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblInst.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblInst)

    $txtInst = New-Object System.Windows.Forms.TextBox
    $txtInst.Multiline = $true
    $txtInst.ReadOnly = $true
    $txtInst.ScrollBars = 'Both'
    $txtInst.WordWrap = $true
    $txtInst.Dock = 'Fill'
    $txtInst.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtInst.BackColor = [System.Drawing.Color]::White
    $txtInst.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $txtInst.Text = "$($Task.InstructionText)"
    $form.Controls.Add($txtInst)
    $txtInst.BringToFront()

    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Get-PmTaskRowColor {
    # Returns a [System.Drawing.Color] for the row, or $null to leave default.
    #  overdue + Power Off -> orange
    #  overdue + Power On  -> green
    #  future  (either)    -> yellow
    param($Task, [DateTime]$Today)

    if ([string]::IsNullOrWhiteSpace($Task.DueDate)) { return $null }

    $due = [DateTime]::MinValue
    try { $due = [DateTime]$Task.DueDate } catch { return $null }

    $overdue   = $due.Date -le $Today.Date
    $isPowerOn = ($Task.PowerState -eq 'Power On')

    if ($overdue) {
        if ($isPowerOn) { return [System.Drawing.Color]::FromArgb(200, 235, 200) }   # green
        else            { return [System.Drawing.Color]::FromArgb(255, 205, 155) }   # orange
    } else {
        return [System.Drawing.Color]::FromArgb(255, 245, 190)                        # yellow
    }
}

function Set-PmTaskRowColors {
    param($Item, $RowColor, [bool]$TimeIsOverridden)

    if ($null -ne $RowColor) {
        $Item.UseItemStyleForSubItems = $false
        foreach ($sub in $Item.SubItems) { $sub.BackColor = $RowColor }
    }
    if ($TimeIsOverridden) {
        $Item.UseItemStyleForSubItems = $false
        $Item.SubItems[6].ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
        $Item.SubItems[6].Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    }
}

function Sync-PmArchive {
    # Parses non-canonical HTML in the archive, renames to the canonical
    # schema, writes the companion .json.  Canonical files with a matching
    # .json are skipped.  Duplicates are collapsed onto the canonical name.
    # Returns a small summary object.
    $dir = Get-PmArchiveDir
    if (-not (Test-Path $dir)) { return [PSCustomObject]@{ Processed = 0; Removed = 0 } }

    $processed = 0
    $schemaPattern = '^([A-Za-z0-9]+)_([A-Za-z0-9]+)_([A-Za-z0-9]+)_(\d{4}-\d{2}-\d{2})\.html$'

    # --- Phase 1: parse any HTML that isn't yet canonically named ---
    foreach ($htmlFile in @(Get-ChildItem -Path $dir -Filter "*.html" -File)) {
        $companionJson = Join-Path $dir ($htmlFile.BaseName + ".json")
        $isCanonical = ($htmlFile.Name -match $schemaPattern) -and (Test-Path $companionJson)
        if ($isCanonical) { continue }

        try {
            Write-Log "Sync-PmArchive: parsing $($htmlFile.Name)"
            $parsed = ConvertFrom-PmChecklist -HtmlPath $htmlFile.FullName

            $dateStamp = (Get-Date).ToString("yyyy-MM-dd")
            if (-not [string]::IsNullOrWhiteSpace($parsed.Meta.ChecklistDate)) {
                try { $dateStamp = ([DateTime]$parsed.Meta.ChecklistDate).ToString("yyyy-MM-dd") } catch {}
            }
            $acr = if ($parsed.Meta.Acronym)     { ($parsed.Meta.Acronym     -replace '[^\w]', '') } else { 'UNKNOWN' }
            $mno = if ($parsed.Meta.EquipmentNo) { ($parsed.Meta.EquipmentNo -replace '[^\w]', '') } else { '0' }
            $cn  = if ($parsed.Meta.ChecklistNo) { ($parsed.Meta.ChecklistNo -replace '[^\w]', '') } else { '0' }
            $baseName = "${acr}_${mno}_${cn}_${dateStamp}"

            $targetHtml = Join-Path $dir "$baseName.html"
            $targetJson = Join-Path $dir "$baseName.json"

            if ($htmlFile.FullName -ne $targetHtml) {
                # Overwrite — no timestamp suffixes, ever.
                if (Test-Path $targetHtml) { Remove-Item -Path $targetHtml -Force }
                Move-Item -Path $htmlFile.FullName -Destination $targetHtml -Force
            }
            $parsed | ConvertTo-Json -Depth 12 | Set-Content -Path $targetJson -Encoding UTF8
            Write-Log "Sync-PmArchive: imported -> $baseName"
            $processed++
        } catch {
            Write-Log "Sync-PmArchive: FAILED on $($htmlFile.Name): $($_.Exception.Message)"
        }
    }

    # --- Phase 2: collapse duplicate JSONs that share a logical identity ---
    $removed = 0
    $groups = @{}

    foreach ($jf in @(Get-ChildItem -Path $dir -Filter "*.json" -File)) {
        try {
            $data = Get-Content -Path $jf.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $acr = ($data.Meta.Acronym     -replace '[^\w]', '')
            $mno = ($data.Meta.EquipmentNo -replace '[^\w]', '')
            $cn  = ($data.Meta.ChecklistNo -replace '[^\w]', '')
            $dateStamp = (Get-Date).ToString("yyyy-MM-dd")
            if (-not [string]::IsNullOrWhiteSpace($data.Meta.ChecklistDate)) {
                try { $dateStamp = ([DateTime]$data.Meta.ChecklistDate).ToString("yyyy-MM-dd") } catch {}
            }
            $key = "${acr}_${mno}_${cn}_${dateStamp}"
            if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
            $groups[$key] += $jf
        } catch {
            # Unparseable JSON — mark for deletion
            if (-not $groups.ContainsKey('__orphan__')) { $groups['__orphan__'] = @() }
            $groups['__orphan__'] += $jf
        }
    }

    foreach ($key in $groups.Keys) {
        $files = @($groups[$key] | Sort-Object -Property LastWriteTime -Descending)

        if ($key -eq '__orphan__') {
            foreach ($f in $files) {
                Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                $removed++
            }
            continue
        }

        $canonicalJson = Join-Path $dir "$key.json"
        $canonicalHtml = Join-Path $dir "$key.html"

        # Pick the newest JSON; if it isn't at the canonical path, move it there
        $newest = $files[0]
        if ($newest.FullName -ne $canonicalJson) {
            if (Test-Path $canonicalJson) { Remove-Item -Path $canonicalJson -Force }
            Move-Item -Path $newest.FullName -Destination $canonicalJson -Force
        }

        # Delete every other JSON in the group
        foreach ($f in $files | Select-Object -Skip 1) {
            Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
            $removed++
        }

        # Do the same for HTMLs that share this base name (with or without suffix)
        $htmlCandidates = @(Get-ChildItem -Path $dir -Filter "$key*.html" -File |
                            Sort-Object -Property LastWriteTime -Descending)
        if ($htmlCandidates.Count -gt 0) {
            $newestHtml = $htmlCandidates[0]
            if ($newestHtml.FullName -ne $canonicalHtml) {
                if (Test-Path $canonicalHtml) { Remove-Item -Path $canonicalHtml -Force }
                Move-Item -Path $newestHtml.FullName -Destination $canonicalHtml -Force
            }
            foreach ($f in $htmlCandidates | Select-Object -Skip 1) {
                Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                $removed++
            }
        }
    }

    Write-Log "Sync-PmArchive: processed=$processed removed=$removed"
    return [PSCustomObject]@{ Processed = $processed; Removed = $removed }
}

function Open-PmArchivedHtml {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        [System.Windows.Forms.MessageBox]::Show("Archived HTML not found:`n$Path", "Missing File", "OK", "Warning")
        return
    }
    Start-Process $Path
}

# ============================================================
# Work Tracking — Breaks + Work Budget
# ============================================================

# Initialised once; the PM Checklist sub-tab will later set this.
if ($null -eq $script:pmTotalSelectedHours) { $script:pmTotalSelectedHours = 0.0 }

$script:breakTypes = @(
    'Lunch',
    'Wash-up',
    'Startup',
    'Paid Break',
    'End-of-Day Wash-up',
    'Paperwork',
    'Other'
)

function Get-BreaksFilePath {
    $jobsPath = Get-JobsFilePath
    $dir = Split-Path $jobsPath -Parent
    return (Join-Path $dir "Breaks.json")
}

function Get-Breaks {
    $path = Get-BreaksFilePath
    if (-not (Test-Path $path)) { return ,@() }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return ,@() }
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return ,@() }
        return ,@($data)
    } catch {
        Write-Log "Error reading breaks: $($_.Exception.Message)"
        return ,@()
    }
}

function Save-Breaks {
    param([array]$Breaks)

    $path = Get-BreaksFilePath
    $dir  = Split-Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $tmp = "$path.tmp"
    try {
        $arr = if ($null -eq $Breaks) { @() } else { @($Breaks) }
        $json = if ($arr.Count -eq 0) { "[]" } else {
            $j = ConvertTo-Json -InputObject $arr -Depth 8
            if (-not $j.TrimStart().StartsWith('[')) { $j = "[$j]" }
            $j
        }
        $json | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $path -Force
        Write-Log "Save-Breaks: wrote $($arr.Count) break(s)"
        return $true
    } catch {
        Write-Log "Error saving breaks: $($_.Exception.Message)"
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        return $false
    }
}

function Get-TodaysBreaks {
    $today = (Get-Date).ToString("yyyy-MM-dd")
    $all = Get-Breaks
    return @($all | Where-Object { "$($_.date)" -eq $today })
}

function Get-TodaysJobHours {
    $today = (Get-Date).ToString("yyyy-MM-dd")
    $total = 0.0
    foreach ($job in @(Get-Jobs)) {
        if ($null -eq $job) { continue }
        $jd = ''
        if ($job.createdAt) {
            try { $jd = ([DateTime]$job.createdAt).ToString("yyyy-MM-dd") }
            catch { $jd = "$($job.createdAt)".Split('T')[0] }
        }
        if ($jd -ne $today) { continue }
        $h = Get-JobHours $job
        if ($null -ne $h) { $total += $h }
    }
    return [Math]::Round($total, 2)
}

function Refresh-BreaksList {
    if (-not $script:breaksListView) { return }
    $script:breaksListView.BeginUpdate()
    try {
        $script:breaksListView.Items.Clear()
        foreach ($b in @(Get-TodaysBreaks)) {
            $it = New-Object System.Windows.Forms.ListViewItem("$($b.type)")
            $it.SubItems.Add("$($b.startTime)")   | Out-Null
            $it.SubItems.Add("$($b.durationMin)") | Out-Null
            $it.SubItems.Add("$($b.notes)")       | Out-Null
            $it.Tag = $b
            $script:breaksListView.Items.Add($it) | Out-Null
        }
    } finally {
        $script:breaksListView.EndUpdate()
    }
}

function Refresh-WorkBudget {
    if (-not $script:wtBudgetLabel) { return }

    $wb = $script:config.WorkTracking.WorkBudget
    if (-not $wb) {
        $script:wtBudgetLabel.Text = "WorkBudget not configured."
        return
    }

    # --- Scheduled non-work (skip any component set to 0) ---
    $components = [ordered]@{
        'Lunch'             = [double]$wb.LunchMinutes
        'Wash-up (lunch)'   = [double]$wb.WashupMinutes
        'Startup'           = [double]$wb.StartupMinutes
        'Paid Break 1'      = if (@($wb.PaidBreaksMinutes).Count -ge 1) { [double]$wb.PaidBreaksMinutes[0] } else { 0 }
        'Paid Break 2'      = if (@($wb.PaidBreaksMinutes).Count -ge 2) { [double]$wb.PaidBreaksMinutes[1] } else { 0 }
        'Wash-up (end)'     = [double]$wb.EndOfDayWashupMinutes
        'Paperwork'         = [double]$wb.PaperworkMinutes
    }

    $nonWorkMin = 0
    foreach ($v in $components.Values) { $nonWorkMin += $v }
    $nonWorkH = $nonWorkMin / 60.0

    $target   = [double]$wb.WorkTargetHours
    $dayTotal = $target + $nonWorkH

    # --- Actuals for today ---
    $breakMin = 0
    foreach ($b in @(Get-TodaysBreaks)) { $breakMin += [int]$b.durationMin }
    $breakHrs = [Math]::Round($breakMin / 60.0, 2)

    $jobHours  = Get-TodaysJobHours
    $pmHours   = [double]$script:pmTotalSelectedHours
    $work      = [Math]::Round($jobHours + $pmHours, 2)
    $over      = [Math]::Round($work - $target, 2)
    $remaining = [Math]::Round($target - $work, 2)

    # --- Build text ---
    $lines = @()
    $lines += "Work Budget — $(Get-Date -Format 'yyyy-MM-dd')"
    $lines += ""
    $lines += "Scheduled non-work"
    foreach ($pair in $components.GetEnumerator()) {
        if ($pair.Value -le 0) {
            $lines += ("  {0,-18} {1,6}" -f $pair.Key, "  --")
        } else {
            $lines += ("  {0,-18} {1,6:F2} h" -f $pair.Key, ($pair.Value / 60.0))
        }
    }
    $lines += ("  {0,-18} {1,6:F2} h" -f "  Subtotal", $nonWorkH)
    $lines += ("  {0,-18} {1,6:F2} h" -f "  Work target", $target)
    $lines += ("  {0,-18} {1,6:F2} h" -f "  = Day total", $dayTotal)
    $lines += ""
    $lines += "Today so far"
    $lines += ("  PM tasks           {0,6:F2} h" -f $pmHours)
    $lines += ("  Reactive/WO        {0,6:F2} h" -f $jobHours)
    $lines += ("  Work total         {0,6:F2} h" -f $work)
    $lines += ("  Breaks logged      {0,6:F2} h" -f $breakHrs)
    $lines += "  -------------------------"
    if ($over -gt 0) {
        $lines += ("  OVER BY            {0,6:F2} h" -f $over)
    } else {
        $lines += ("  Remaining          {0,6:F2} h" -f $remaining)
    }

    # --- Grievance hint ---
    $grievance = $false
    if ($over -gt 0) {
        $grievance = $true
        $lines += ""
        $lines += "  ! Work has exceeded your $($target.ToString('F2'))h target."
        $lines += "    If non-work time was skipped to complete this"
        $lines += "    work, it may be grievance-worthy per your local"
        $lines += "    agreement. Note the skipped components and notify"
        $lines += "    your steward."
    }

    # --- Highlight over-day-length case (distinct from over-target) ---
    $grandTotal = $work + $breakHrs
    if ($grandTotal -gt $dayTotal) {
        $lines += ""
        $lines += ("  ! Total time {0:F2}h exceeds day length {1:F2}h." -f $grandTotal, $dayTotal)
    }

    $script:wtBudgetLabel.Text = ($lines -join "`r`n")
    if ($grievance) {
        $script:wtBudgetLabel.ForeColor = [System.Drawing.Color]::FromArgb(179,38,30)
    } else {
        $script:wtBudgetLabel.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    }
}

function Show-BreakEditor {
    param($Break = $null)

    if ($null -eq $Break) {
        $Break = [PSCustomObject]@{
            id = [guid]::NewGuid().ToString()
            date = (Get-Date).ToString("yyyy-MM-dd")
            startTime = (Get-Date).ToString("HH:mm")
            durationMin = 15
            type = 'Paid Break'
            notes = ''
        }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Break"
    $form.Size = New-Object System.Drawing.Size(420, 300)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $mkLabel = {
        param($t, $y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $t
        $l.Location = New-Object System.Drawing.Point(16, $y)
        $l.Size = New-Object System.Drawing.Size(100, 24)
        $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
        return $l
    }

    $form.Controls.Add((& $mkLabel "Type:" 20))
    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.Location = New-Object System.Drawing.Point(124, 20)
    $cmbType.Size = New-Object System.Drawing.Size(260, 24)
    $cmbType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($t in $script:breakTypes) { [void]$cmbType.Items.Add($t) }
    if ($cmbType.Items.Contains($Break.type)) { $cmbType.SelectedItem = $Break.type } else { $cmbType.SelectedIndex = 0 }
    $form.Controls.Add($cmbType)

    $form.Controls.Add((& $mkLabel "Start:" 60))
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpStart.CustomFormat = "HH:mm"
    $dtpStart.ShowUpDown = $true
    $dtpStart.Location = New-Object System.Drawing.Point(124, 60)
    $dtpStart.Size = New-Object System.Drawing.Size(90, 24)
    try { $dtpStart.Value = [DateTime]::ParseExact($Break.startTime, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { $dtpStart.Value = Get-Date }
    $form.Controls.Add($dtpStart)

    $btnNow = New-Object System.Windows.Forms.Button
    $btnNow.Text = "Now"
    $btnNow.Location = New-Object System.Drawing.Point(220, 60)
    $btnNow.Size = New-Object System.Drawing.Size(56, 24)
    $btnNow.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnNow.FlatAppearance.BorderSize = 0
    $btnNow.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnNow.ForeColor = [System.Drawing.Color]::White
    $btnNow.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnNow.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNow.Add_Click({ $dtpStart.Value = Get-Date }.GetNewClosure())
    $form.Controls.Add($btnNow)

    $form.Controls.Add((& $mkLabel "Duration (min):" 100))
    $numDur = New-Object System.Windows.Forms.NumericUpDown
    $numDur.Location = New-Object System.Drawing.Point(124, 100)
    $numDur.Size = New-Object System.Drawing.Size(90, 24)
    $numDur.Minimum = 1
    $numDur.Maximum = 480
    $numDur.Value = if ([int]$Break.durationMin -gt 0) { [int]$Break.durationMin } else { 15 }
    $form.Controls.Add($numDur)

    $form.Controls.Add((& $mkLabel "Notes:" 140))
    $txtNotes = New-Object System.Windows.Forms.TextBox
    $txtNotes.Location = New-Object System.Drawing.Point(124, 140)
    $txtNotes.Size = New-Object System.Drawing.Size(260, 60)
    $txtNotes.Multiline = $true
    $txtNotes.Text = "$($Break.notes)"
    $form.Controls.Add($txtNotes)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(180, 220)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 30)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $form.Controls.Add($cancelBtn)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Save"
    $saveBtn.Location = New-Object System.Drawing.Point(280, 220)
    $saveBtn.Size = New-Object System.Drawing.Size(100, 30)
    $saveBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveBtn.FlatAppearance.BorderSize = 0
    $saveBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $saveBtn.ForeColor = [System.Drawing.Color]::White
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $saveBtn.Add_Click({
        $form.Tag = [PSCustomObject]@{
            id          = $Break.id
            date        = $Break.date
            startTime   = $dtpStart.Value.ToString('HH:mm')
            durationMin = [int]$numDur.Value
            type        = "$($cmbType.SelectedItem)"
            notes       = $txtNotes.Text.Trim()
        }
        $form.Close()
    })
    $form.Controls.Add($saveBtn)

    $form.AcceptButton = $saveBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

# ============================================================
# Work Tracking — Job Log sub-tab
# ============================================================

function Get-JobHours {
    param($Job)
    if ($null -eq $Job) { return $null }

    if ($null -ne $Job.hoursOverride -and "$($Job.hoursOverride)" -ne '') {
        $h = 0.0
        if ([double]::TryParse("$($Job.hoursOverride)", [ref]$h)) { return $h }
    }

    if ([string]::IsNullOrWhiteSpace($Job.startTime) -or
        [string]::IsNullOrWhiteSpace($Job.endTime)) { return $null }

    try {
        $s = [DateTime]::ParseExact($Job.startTime, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        $e = [DateTime]::ParseExact($Job.endTime,   'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($e -lt $s) { $e = $e.AddDays(1) }
        return [Math]::Round(($e - $s).TotalHours, 2)
    } catch { return $null }
}

function Refresh-JobLogList {
    if (-not $script:jobLogListView) { return }

    $script:wtJobs = Get-Jobs
    $jobs = $script:wtJobs
    if ($null -eq $jobs) { $jobs = @() }

    $script:jobLogListView.BeginUpdate()
    try {
        $script:jobLogListView.Items.Clear()

        foreach ($job in $jobs) {
            $kindText = if ($job.kind -eq 'workorder') { 'WORK ORDER' } else { 'REACTIVE' }

            $dateText = ''
            if ($job.createdAt) {
                try { $dateText = ([DateTime]$job.createdAt).ToString("yyyy-MM-dd") }
                catch { $dateText = "$($job.createdAt)".Split('T')[0] }
            }

            $hrsText = ''
            $h = Get-JobHours $job
            if ($null -ne $h) { $hrsText = $h.ToString("F2") }

            $partsCount = 0
            if ($job.parts) { $partsCount = @($job.parts).Count }

            $statusText = if ($job.sentToHistorian) { 'SENT' } else { 'DRAFT' }

            $item = New-Object System.Windows.Forms.ListViewItem($kindText)
            $item.SubItems.Add("$($job.machineId)")   | Out-Null
            $item.SubItems.Add($dateText)             | Out-Null
            $item.SubItems.Add("$($job.startTime)")   | Out-Null
            $item.SubItems.Add("$($job.endTime)")     | Out-Null
            $item.SubItems.Add($hrsText)              | Out-Null
            $item.SubItems.Add("$($job.workOrderNo)") | Out-Null
            $item.SubItems.Add("$partsCount")         | Out-Null
            $item.SubItems.Add("$($job.description)") | Out-Null
            $item.SubItems.Add($statusText)           | Out-Null
            $item.Tag = $job
            $script:jobLogListView.Items.Add($item) | Out-Null
        }
    } finally {
        $script:jobLogListView.EndUpdate()
    }

    if ($script:jobLogStatusLabel) {
        $total = @($jobs).Count
        $drafts = @($jobs | Where-Object { -not $_.sentToHistorian }).Count
        $script:jobLogStatusLabel.Text = "$total job$(if ($total -ne 1) { 's' }) — $drafts draft$(if ($drafts -ne 1) { 's' })"
    }
}

function Setup-JobLogSubTab {
    param($parentTab)

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(0)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = 'Fill'
    $root.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $root.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($root)

    # Status bar (bottom-most)
    $statusBar = New-Object System.Windows.Forms.Panel
    $statusBar.Dock = 'Bottom'
    $statusBar.Height = 26
    $statusBar.BackColor = [System.Drawing.Color]::FromArgb(236,240,241)
    $statusBar.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)
    $root.Controls.Add($statusBar)

    $script:jobLogStatusLabel = New-Object System.Windows.Forms.Label
    $script:jobLogStatusLabel.Dock = 'Fill'
    $script:jobLogStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:jobLogStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:jobLogStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $script:jobLogStatusLabel.Text = "0 jobs"
    $statusBar.Controls.Add($script:jobLogStatusLabel)

    # Splitter: job list on top, breaks/budget on bottom
    $script:jobLogSplitInitialized = $false
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $split.SplitterWidth = 4
    $split.BackColor = [System.Drawing.Color]::FromArgb(220,225,232)

    # Give it a real height BEFORE setting the panel min-sizes, otherwise
    # WinForms' validation math sees Height - Panel2MinSize < Panel1MinSize
    # and throws.
    $split.Height = 600
    $split.Panel1MinSize = 220
    $split.Panel2MinSize = 240
    $split.SplitterDistance = 340   # any value between Panel1MinSize and (Height - Panel2MinSize)

    $root.Controls.Add($split)
    $split.BringToFront()

    $split.Add_SizeChanged({
        if (-not $script:jobLogSplitInitialized -and $this.Height -gt 100) {
            $this.SplitterDistance = [int]($this.Height * 0.58)
            $script:jobLogSplitInitialized = $true
        }
    })

    # ---- Top panel: toolbar + job list ----
    $topPanel = $split.Panel1
    $topPanel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $toolbar.Dock = 'Top'
    $toolbar.Height = 52
    $toolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $toolbar.WrapContents = $false
    $toolbar.BackColor = [System.Drawing.Color]::Transparent
    $toolbar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $topPanel.Controls.Add($toolbar)

    $newReactiveBtn = New-Button "+ New Reactive Call" {
        $result = Show-JobEditor -Job $null -Kind 'reactive'
        if ($null -ne $result -and $result.Action -eq 'save') {
            $jobs = Get-Jobs
            if ($null -eq $jobs) { $jobs = @() }
            $jobs += $result.Job
            Save-Jobs -Jobs $jobs | Out-Null
            Refresh-JobLogList
            Refresh-WorkBudget
        }
    } -Style 'Primary' -Width 180 -Height 32 -TextAlign MiddleCenter
    $newReactiveBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($newReactiveBtn)

    $newWoBtn = New-Button "+ New Work Order" {
        $result = Show-JobEditor -Job $null -Kind 'workorder'
        if ($null -ne $result -and $result.Action -eq 'save') {
            $jobs = Get-Jobs
            if ($null -eq $jobs) { $jobs = @() }
            $jobs += $result.Job
            Save-Jobs -Jobs $jobs | Out-Null
            Refresh-JobLogList
            Refresh-WorkBudget
        }
    } -Style 'Primary' -Width 170 -Height 32 -TextAlign MiddleCenter
    $newWoBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($newWoBtn)

    $refreshBtn = New-Button "Refresh" {
        Refresh-JobLogList
        Refresh-BreaksList
        Refresh-WorkBudget
    } -Style 'Ghost' -Width 100 -Height 32 -TextAlign MiddleCenter
    $refreshBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($refreshBtn)

    $sendAllBtn = New-Button "Send all ready" {
        $result = Send-AllReadyJobs
        Refresh-JobLogList

        $msg = "Sent: $($result.Sent)   Skipped: $($result.Skipped)   Failed: $($result.Failed)"
        if (@($result.Reasons).Count -gt 0) {
            $msg += "`r`n`r`nSkipped reasons:`r`n" + (($result.Reasons | Select-Object -First 10) -join "`r`n")
            if ($result.Reasons.Count -gt 10) { $msg += "`r`n... ($($result.Reasons.Count - 10) more)" }
        }
        [System.Windows.Forms.MessageBox]::Show($msg, "Send all ready", "OK", "Information")
    } -Style 'Success' -Width 150 -Height 32 -TextAlign MiddleCenter
    $sendAllBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($sendAllBtn)

    $script:jobLogListView = New-Object System.Windows.Forms.ListView
    $script:jobLogListView.Dock = 'Fill'
    $script:jobLogListView.Columns.Add("Kind", 90)         | Out-Null
    $script:jobLogListView.Columns.Add("Machine", 110)     | Out-Null
    $script:jobLogListView.Columns.Add("Date", 100)        | Out-Null
    $script:jobLogListView.Columns.Add("Start", 60)        | Out-Null
    $script:jobLogListView.Columns.Add("End", 60)          | Out-Null
    $script:jobLogListView.Columns.Add("Hrs", 50)          | Out-Null
    $script:jobLogListView.Columns.Add("WO#", 110)         | Out-Null
    $script:jobLogListView.Columns.Add("Parts", 50)        | Out-Null
    $script:jobLogListView.Columns.Add("Description", 320) | Out-Null
    $script:jobLogListView.Columns.Add("Status", 70)       | Out-Null
    Set-ListViewStyle -ListView $script:jobLogListView
    $topPanel.Controls.Add($script:jobLogListView)
    $script:jobLogListView.BringToFront()

    # ---- Bottom panel: budget (left) + breaks (right) ----
    $bottomPanel = $split.Panel2
    $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

    $bottomGrid = New-Object System.Windows.Forms.TableLayoutPanel
    $bottomGrid.Dock = 'Fill'
    $bottomGrid.ColumnCount = 2
    $bottomGrid.RowCount = 1
    $bottomGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 300))) | Out-Null
    $bottomGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100.0))) | Out-Null
    $bottomPanel.Controls.Add($bottomGrid)

    # Left: Work Budget
    $gbBudget = New-Object System.Windows.Forms.GroupBox
    $gbBudget.Text = "Work Budget (Today)"
    $gbBudget.Dock = 'Fill'
    $gbBudget.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gbBudget.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gbBudget.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)

    $script:wtBudgetLabel = New-Object System.Windows.Forms.Label
    $script:wtBudgetLabel.Dock = 'Fill'
    $script:wtBudgetLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:wtBudgetLabel.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $script:wtBudgetLabel.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    $script:wtBudgetLabel.Text = "Loading..."
    $gbBudget.Controls.Add($script:wtBudgetLabel)

    $bottomGrid.Controls.Add($gbBudget, 0, 0)

    # Right: Breaks
    $gbBreaks = New-Object System.Windows.Forms.GroupBox
    $gbBreaks.Text = "Breaks & Non-Work Time (Today)"
    $gbBreaks.Dock = 'Fill'
    $gbBreaks.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gbBreaks.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gbBreaks.Margin = New-Object System.Windows.Forms.Padding(0)

    $breaksToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $breaksToolbar.Dock = 'Bottom'
    $breaksToolbar.Height = 36
    $breaksToolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $breaksToolbar.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)
    $breaksToolbar.BackColor = [System.Drawing.Color]::Transparent
    $gbBreaks.Controls.Add($breaksToolbar)

    $addBreakBtn = New-Object System.Windows.Forms.Button
    $addBreakBtn.Text = "+ Add Break"
    $addBreakBtn.Size = New-Object System.Drawing.Size(110, 26)
    $addBreakBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addBreakBtn.FlatAppearance.BorderSize = 0
    $addBreakBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $addBreakBtn.ForeColor = [System.Drawing.Color]::White
    $addBreakBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $addBreakBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $addBreakBtn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    $breaksToolbar.Controls.Add($addBreakBtn)

    $rmBreakBtn = New-Object System.Windows.Forms.Button
    $rmBreakBtn.Text = "Remove"
    $rmBreakBtn.Size = New-Object System.Drawing.Size(90, 26)
    $rmBreakBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $rmBreakBtn.FlatAppearance.BorderSize = 0
    $rmBreakBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $rmBreakBtn.ForeColor = [System.Drawing.Color]::White
    $rmBreakBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $rmBreakBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $breaksToolbar.Controls.Add($rmBreakBtn)

    $script:breaksListView = New-Object System.Windows.Forms.ListView
    $script:breaksListView.Dock = 'Fill'
    $script:breaksListView.Columns.Add("Type", 130) | Out-Null
    $script:breaksListView.Columns.Add("Start", 60) | Out-Null
    $script:breaksListView.Columns.Add("Min", 50)   | Out-Null
    $script:breaksListView.Columns.Add("Notes", 260) | Out-Null
    Set-ListViewStyle -ListView $script:breaksListView
    $gbBreaks.Controls.Add($script:breaksListView)
    $script:breaksListView.BringToFront()

    $bottomGrid.Controls.Add($gbBreaks, 1, 0)

    # ---- Break editor wiring ----
    $addBreakBtn.Add_Click({
        $result = Show-BreakEditor
        if ($null -eq $result) { return }
        $all = Get-Breaks
        if ($null -eq $all) { $all = @() }
        $all += $result
        Save-Breaks -Breaks $all | Out-Null
        Refresh-BreaksList
        Refresh-WorkBudget
    })

    $rmBreakBtn.Add_Click({
        $sel = $script:breaksListView.SelectedItems
        if ($sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select a break to remove.", "Breaks", "OK", "Information")
            return
        }
        $breakId = $sel[0].Tag.id
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Remove this break?", "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $all = Get-Breaks
        $kept = @($all | Where-Object { $_.id -ne $breakId })
        Save-Breaks -Breaks $kept | Out-Null
        Refresh-BreaksList
        Refresh-WorkBudget
    })

    # Double-click job list — edit
    $script:jobLogListView.Add_DoubleClick({
        $sel = $script:jobLogListView.SelectedItems
        if ($sel.Count -eq 0) { return }

        $original = $sel[0].Tag
        $result = Show-JobEditor -Job $original -Kind $original.kind
        if ($null -eq $result) { return }

        $jobs = Get-Jobs
        if ($null -eq $jobs) { $jobs = @() }

        if ($result.Action -eq 'delete') {
            $jobs = @($jobs | Where-Object { $_.id -ne $original.id })
        } elseif ($result.Action -eq 'save') {
            $found = $false
            for ($i = 0; $i -lt $jobs.Count; $i++) {
                if ($jobs[$i].id -eq $original.id) {
                    $jobs[$i] = $result.Job
                    $found = $true
                    break
                }
            }
            if (-not $found) { $jobs += $result.Job }
        }

        Save-Jobs -Jobs $jobs | Out-Null
        Refresh-JobLogList
        Refresh-WorkBudget
    })

    # Initial loads
    Refresh-JobLogList
    Refresh-BreaksList
    Refresh-WorkBudget

    Write-Log "Job Log sub-tab setup completed."
}

function Show-PmChecklistPasteDialog {
    # Modal with a large paste box.  Returns the pasted HTML string, or $null.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Import PM Checklist"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Paste the raw HTML from the PM Checklist popup (Ctrl-U in the popup, Ctrl-A, Ctrl-C)."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 28
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(680, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Import"
    $okBtn.Location = New-Object System.Drawing.Point(780, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $okBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $okBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPaste.Text)) { return }
        $form.Tag = $txtPaste.Text
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    $txtPaste = New-Object System.Windows.Forms.TextBox
    $txtPaste.Multiline = $true
    $txtPaste.MaxLength = 0
    $txtPaste.Dock = 'Fill'
    $txtPaste.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtPaste.WordWrap = $false
    $txtPaste.Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtPaste.AcceptsReturn = $true
    $txtPaste.AcceptsTab = $true
    $txtPaste.Padding = New-Object System.Windows.Forms.Padding(8)
    $form.Controls.Add($txtPaste)
    $txtPaste.BringToFront()

    # Do NOT set AcceptButton — Enter must insert newlines
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Update-PmMachineTabTotals {
    param($Machine, [System.Windows.Forms.Control]$Tab)
    if (-not $Machine -or -not $Machine.State) { return }

    $d = 0; $mins = 0
    foreach ($t in @($Machine.Parsed.Tasks)) {
        $e = $Machine.State[$t.ItemNo]
        if ($e -and $e.Selected) {
            $d++
            $mins += if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { [int]$t.EstTimeMin }
        }
    }
    $tl = $Tab.Controls |
          Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Dock -eq 'Bottom' } |
          Select-Object -First 1
    if ($tl) {
        $tl.Text = "Done: $d of $($Machine.TaskCount)  •  $([Math]::Round($mins/60.0, 2)) h"
    }

    Update-PmSelectedHoursTotal
}

function Refresh-PmMachineTabs {
    if (-not $script:pmMachineTabs) { return }
    if (-not $script:pmDatePicker)   { return }

    $script:pmSuppressEvents = $true
    $script:pmMachineTabs.TabPages.Clear()

    $selected = $script:pmDatePicker.Value.Date
    $all = @($script:pmMachines)
    $shown = @()
    foreach ($m in $all) {
        if ($m.SortDate.Date -eq $selected) { $shown += $m }
    }

    if ($shown.Count -eq 0) {
        $emptyTab = New-Object System.Windows.Forms.TabPage
        $emptyTab.Text = "(none)"
        $emptyTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
        $dateStr = $selected.ToString("yyyy-MM-dd")
        $hint = if ($all.Count -gt 0) {
            "No PM Checklists for $dateStr.`r`n`r`n$($all.Count) archived checklist(s) exist on other dates.`r`nTry a different date, click 'Scan Archive', or '+ Import PM Checklist'."
        } else {
            "No PM Checklists for $dateStr.`r`n`r`nClick 'Scan Archive' to look for HTML files,`r`nor '+ Import PM Checklist' to paste one in."
        }
        $msg = New-Object System.Windows.Forms.Label
        $msg.Text = $hint
        $msg.Dock = 'Fill'
        $msg.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $msg.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Italic)
        $msg.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
        $emptyTab.Controls.Add($msg)
        $script:pmMachineTabs.TabPages.Add($emptyTab) | Out-Null
        $script:pmSuppressEvents = $false
        Update-PmSelectedHoursTotal
        return
    }

    foreach ($m in $shown) {
        # Load session on first render
        if (-not $m.State) {
            $sess = Get-PmSession -Machine $m
            $m | Add-Member -NotePropertyName State -NotePropertyValue $sess.Selections -Force
        }

        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = "$($m.MachineId)  ($($m.TaskCount))"
        $tab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
        $tab.Padding = New-Object System.Windows.Forms.Padding(12)

        # Top metadata
        $meta = New-Object System.Windows.Forms.Label
        $meta.Dock = 'Top'
        $meta.Height = 120
        $meta.Font = New-Object System.Drawing.Font("Consolas", 9)
        $meta.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
        $meta.Text = @(
            "Machine:        $($m.MachineId)"
            "Checklist No:   $($m.ChecklistNo)"
            "Checklist Date: $($m.ChecklistDate)"
            "Generated On:   $($m.GeneratedOn)"
            "Skills:         $($m.Skills)"
            "Tasks:          $($m.TaskCount)"
            "Est Open:       $($m.EstOpen) h"
            "Est Due Today:  $($m.EstDueToday) h"
        ) -join "`r`n"
        $tab.Controls.Add($meta)

        $openBtn = New-Object System.Windows.Forms.Button
        $openBtn.Text = "Open Original HTML"
        $openBtn.Size = New-Object System.Drawing.Size(160, 30)
        $openBtn.FlatStyle = 'Flat'
        $openBtn.FlatAppearance.BorderSize = 0
        $openBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
        $openBtn.ForeColor = [System.Drawing.Color]::White
        $openBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $openBtn.Cursor = 'Hand'
        $openBtn.Location = New-Object System.Drawing.Point(12, 128)
        $openBtn.Add_Click({ Open-PmArchivedHtml -Path $m.HtmlPath }.GetNewClosure())
        $tab.Controls.Add($openBtn)

        # Totals line at bottom
        $totals = New-Object System.Windows.Forms.Label
        $totals.Dock = 'Bottom'
        $totals.Height = 28
        $totals.Padding = New-Object System.Windows.Forms.Padding(6, 4, 0, 0)
        $totals.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
        $totals.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
        $totals.Text = ""
        $tab.Controls.Add($totals)

        # Task ListView
        $lv = New-Object System.Windows.Forms.ListView
        $lv.Location = New-Object System.Drawing.Point(12, 168)
        $lv.Size = New-Object System.Drawing.Size(900, 380)
        $lv.Anchor = 'Top,Left,Right,Bottom'
        $lv.CheckBoxes = $true
        $lv.Columns.Add("", 30)          | Out-Null
        $lv.Columns.Add("Item", 70)      | Out-Null
        $lv.Columns.Add("Component", 240)| Out-Null
        $lv.Columns.Add("Power", 70)     | Out-Null
        $lv.Columns.Add("Skill", 60)     | Out-Null
        $lv.Columns.Add("Est", 50)       | Out-Null
        $lv.Columns.Add("Time", 60)      | Out-Null
        $lv.Columns.Add("Due", 90)       | Out-Null
        Set-ListViewStyle -ListView $lv
        $lv.Tag = $m          # for event handlers

        $doneCount = 0
        $selectedMinutes = 0

        foreach ($task in @($m.Parsed.Tasks)) {
            $e = $m.State[$task.ItemNo]
            $isDone = if ($e) { [bool]$e.Selected } else { $false }
            $displayMin = if ($e -and $null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { [int]$task.EstTimeMin }
            $isOverridden = ($e -and $null -ne $e.CustomTimeMin)

            if ($isDone) {
                $doneCount++
                $selectedMinutes += $displayMin
            }

            $it = New-Object System.Windows.Forms.ListViewItem("")
            $it.Checked = $isDone
            $it.SubItems.Add("$($task.ItemNo)")        | Out-Null
            $it.SubItems.Add("$($task.Component)")     | Out-Null
            $it.SubItems.Add("$($task.PowerState)")    | Out-Null
            $it.SubItems.Add("$($task.MinSkillLevel)") | Out-Null
            $it.SubItems.Add("$($task.EstTimeMin)")    | Out-Null
            $it.SubItems.Add("$displayMin")            | Out-Null
            $it.SubItems.Add("$($task.DueDate)")       | Out-Null
            $it.Tag = $task

            $rowColor = Get-PmTaskRowColor -Task $task -Today (Get-Date).Date
            Set-PmTaskRowColors -Item $it -RowColor $rowColor -TimeIsOverridden $isOverridden

            $lv.Items.Add($it) | Out-Null
        }

        $totals.Text = "Done: $doneCount of $($m.TaskCount)  •  $([Math]::Round($selectedMinutes/60.0, 2)) h"

        $lv.Add_ItemChecked({
            param($sender, $e)
            if ($script:pmSuppressEvents) { return }

            $machine = $m
            if (-not $machine -or -not $machine.State) {
                Write-Log "PM checkbox: no machine bound, ignoring."
                return
            }

            $item = $e.Item
            if (-not $item) {
                Write-Log "PM checkbox: null e.Item, ignoring."
                return
            }
            $task = $item.Tag
            if (-not $task) {
                Write-Log "PM checkbox: null item.Tag, ignoring."
                return
            }

            $existing = $machine.State[$task.ItemNo]
            $custom = if ($existing) { $existing.CustomTimeMin } else { $null }
            $machine.State[$task.ItemNo] = @{ Selected = [bool]$item.Checked; CustomTimeMin = $custom }

            Write-Log "PM checkbox: $($machine.MachineId) $($task.ItemNo) selected=$($item.Checked) customMin=$custom"

            Save-PmSession -Machine $machine

            Update-PmMachineTabTotals -Machine $machine -Tab $sender.Parent
        }.GetNewClosure())

        $lv.Add_DoubleClick({
            param($sender, $e)
            if ($script:pmSuppressEvents) { return }

            $machine = $m
            if (-not $machine -or -not $machine.State) { return }

            $sel = $sender.SelectedItems
            if ($sel.Count -eq 0) { return }
            $item = $sel[0]
            $task = $item.Tag
            if (-not $task) { return }

            $result = Show-PmTaskEditor -Machine $machine -Task $task
            if ($null -eq $result) { return }

            $machine.State[$task.ItemNo] = @{
                Selected      = [bool]$result.Selected
                CustomTimeMin = $result.CustomTimeMin
            }
            Save-PmSession -Machine $machine

            # Update the row in place
            $item.Checked = [bool]$result.Selected
            $displayMin = if ($null -ne $result.CustomTimeMin) { [int]$result.CustomTimeMin } else { [int]$task.EstTimeMin }
            $item.SubItems[6].Text = "$displayMin"

            # Reapply row colors + override styling
            $rowColor = Get-PmTaskRowColor -Task $task -Today (Get-Date).Date
            Set-PmTaskRowColors -Item $item -RowColor $rowColor -TimeIsOverridden ($null -ne $result.CustomTimeMin)

            # Refresh per-machine totals
            Update-PmMachineTabTotals -Machine $machine -Tab $sender.Parent
        }.GetNewClosure())

        $tab.Controls.Add($lv)
        $script:pmMachineTabs.TabPages.Add($tab) | Out-Null
    }

    $script:pmSuppressEvents = $false
    Update-PmSelectedHoursTotal
}

function Setup-PmChecklistSubTab {
    param($parentTab)

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(0)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = 'Fill'
    $root.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $root.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($root)

    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $toolbar.Dock = 'Top'
    $toolbar.Height = 52
    $toolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $toolbar.WrapContents = $false
    $toolbar.BackColor = [System.Drawing.Color]::Transparent
    $toolbar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $root.Controls.Add($toolbar)

    # Import (paste)
    $importBtn = New-Button "+ Import PM Checklist" {
        $html = Show-PmChecklistPasteDialog
        if ($null -eq $html) { return }
        try {
            $summary = Import-PmChecklist -Html $html

            # Ensure the date picker matches the imported checklist's date
            $importDate = (Get-Date).Date
            if (-not [string]::IsNullOrWhiteSpace($summary.DateStamp)) {
                try { $importDate = [DateTime]$summary.DateStamp } catch {}
            }
            if ($script:pmDatePicker) { $script:pmDatePicker.Value = $importDate }

            $script:pmMachines = Load-PmMachines
            Refresh-PmMachineTabs

            [System.Windows.Forms.MessageBox]::Show(
                "Imported $($summary.MachineId)  ($($summary.DateStamp))`r`n$($summary.TaskCount) tasks`r`nArchived to:`r`n$($summary.HtmlPath)",
                "Import Complete", "OK", "Information")
        } catch {
            Write-Log "Import failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Import failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
    } -Style 'Primary' -Width 200 -Height 32 -TextAlign MiddleCenter
    $importBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($importBtn)

    # Scan archive for new raw HTML files
    $scanBtn = New-Button "Scan Archive" {
        try {
            $result = Sync-PmArchive
            $script:pmMachines = Load-PmMachines
            Refresh-PmMachineTabs
            $parts = @()
            if ($result.Processed -gt 0) { $parts += "$($result.Processed) new file$(if ($result.Processed -ne 1) { 's' }) imported" }
            if ($result.Removed   -gt 0) { $parts += "$($result.Removed) duplicate$(if ($result.Removed -ne 1) { 's' }) removed" }
            if ($parts.Count -eq 0)      { $parts += "Archive is clean" }
            [System.Windows.Forms.MessageBox]::Show(($parts -join ". ") + ".", "Scan Complete", "OK", "Information")
        } catch {
            Write-Log "Scan failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Scan failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
    } -Style 'Success' -Width 140 -Height 32 -TextAlign MiddleCenter
    $scanBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($scanBtn)

    # Reload (re-reads JSONs from disk without scanning for new files)
    $reloadBtn = New-Button "Reload" {
        $script:pmMachines = Load-PmMachines
        Refresh-PmMachineTabs
        Write-Log "PM Checklist: reloaded $($script:pmMachines.Count) machine(s)."
    } -Style 'Ghost' -Width 90 -Height 32 -TextAlign MiddleCenter
    $reloadBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($reloadBtn)

    $sendPmBtn = New-Button "Send Checklist" {
        $date   = $script:pmDatePicker.Value.Date
        $result = Send-PmChecklistToHistorian -Date $date

        if ($result.Written -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Nothing to send for $($date.ToString('yyyy-MM-dd')).`r`n`r`nNo tasks are marked complete for that date.",
                "Send Checklist", "OK", "Information")
        } else {
            $msg = "Sent $($result.Written) machine checklist$(if ($result.Written -ne 1) { 's' }) to Historian for $($date.ToString('yyyy-MM-dd'))."
            $msg += "`r`n`r`nSending again for the same date will replace the previous version."
            [System.Windows.Forms.MessageBox]::Show($msg, "Send Checklist", "OK", "Information")
        }
        Write-Log "PM Checklist: sent $($result.Written) machine(s) for $($date.ToString('yyyy-MM-dd'))."
    } -Style 'Success' -Width 150 -Height 32 -TextAlign MiddleCenter
    $sendPmBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($sendPmBtn)

    # Date picker
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Date:"
    $lblDate.Size = New-Object System.Drawing.Size(44, 32)
    $lblDate.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblDate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblDate.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblDate.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $toolbar.Controls.Add($lblDate)

    $script:pmDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $script:pmDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $script:pmDatePicker.Size = New-Object System.Drawing.Size(120, 26)
    $script:pmDatePicker.Value = (Get-Date).Date
    $script:pmDatePicker.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:pmDatePicker.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
    $script:pmDatePicker.Add_ValueChanged({ Refresh-PmMachineTabs })
    $toolbar.Controls.Add($script:pmDatePicker)

    $todayBtn = New-Button "Today" {
        $script:pmDatePicker.Value = (Get-Date).Date
    } -Style 'Ghost' -Width 70 -Height 32 -TextAlign MiddleCenter
    $todayBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($todayBtn)

    # Open archive folder
    $openFolderBtn = New-Button "Open Archive Folder" {
        $dir = Get-PmArchiveDir
        Start-Process $dir
    } -Style 'Ghost' -Width 170 -Height 32 -TextAlign MiddleCenter
    $openFolderBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($openFolderBtn)

    $script:pmGrandTotalLabel = New-Object System.Windows.Forms.Label
    $script:pmGrandTotalLabel.Text = "Today: 0 tasks / 0 h"
    $script:pmGrandTotalLabel.Size = New-Object System.Drawing.Size(200, 32)
    $script:pmGrandTotalLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:pmGrandTotalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:pmGrandTotalLabel.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $toolbar.Controls.Add($script:pmGrandTotalLabel)

    # Machine tabs
    $script:pmMachineTabs = New-Object System.Windows.Forms.TabControl
    $script:pmMachineTabs.Dock = 'Fill'
    $script:pmMachineTabs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $root.Controls.Add($script:pmMachineTabs)
    $script:pmMachineTabs.BringToFront()

    # Initial state: sync first, then load, then render today
    try {
        [void](Sync-PmArchive)
    } catch {
        Write-Log "Startup sync failed: $($_.Exception.Message)"
    }
    $script:pmMachines = Load-PmMachines
    Refresh-PmMachineTabs

    Write-Log "PM Checklist sub-tab setup completed."
}

function Refresh-WeeklyWorksheetList {
    if (-not $script:wsListView) { return }
    if (-not $script:wsDatePicker) { return }

    $script:wsListView.BeginUpdate()
    try {
        $script:wsListView.Items.Clear()

        $date = $script:wsDatePicker.Value.Date
        $ws = $script:wsWorksheet
        if ($null -eq $ws -or $null -eq $ws.Rows) {
            if ($script:wsStatusLabel) {
                $script:wsStatusLabel.Text = "No worksheet loaded for $($date.ToString('yyyy-MM-dd'))."
            }
            return
        }

        $shown = 0
        foreach ($row in @($ws.Rows)) {
            $actual = if ($null -ne $row.ActualTime) { ("{0:F2}" -f [double]$row.ActualTime) } else { '' }
            $isAuto = ($null -ne $row.AutoValue) -and (-not $row.ManualOverride)

            $it = New-Object System.Windows.Forms.ListViewItem("$($row.WorkOrderNo)")
            $it.SubItems.Add("$($row.Acronym)")      | Out-Null
            $it.SubItems.Add("$($row.ClassCode)")    | Out-Null
            $it.SubItems.Add("$($row.Equipment)")    | Out-Null
            $it.SubItems.Add("$($row.DueDate)")      | Out-Null
            $it.SubItems.Add("$($row.PmDescription)")| Out-Null
            $it.SubItems.Add("$($row.Description)")  | Out-Null
            $it.SubItems.Add("$($row.EstimatedHrs)") | Out-Null
            $flash = if ($isAuto) { [string][char]0x26A1 } else { '' }
            $auditFlag = if ($row.NeedsAudit) { 'AUDIT' } else { '' }
            $it.SubItems.Add([string]$actual)   | Out-Null
            $it.SubItems.Add($flash)            | Out-Null
            $it.SubItems.Add($auditFlag)        | Out-Null
            $it.Tag = $row

            if ($row.NeedsAudit) {
                $it.UseItemStyleForSubItems = $false
                foreach ($sub in $it.SubItems) {
                    $sub.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 190)
                }
            } elseif ($isAuto) {
                $it.UseItemStyleForSubItems = $false
                $it.SubItems[8].ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $it.SubItems[8].Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }

            $script:wsListView.Items.Add($it) | Out-Null
            $shown++
        }

        # Totals
        $sumEst = 0.0; $sumAct = 0.0; $audit = 0
        foreach ($row in @($ws.Rows)) {
            [void][double]::TryParse("$($row.EstimatedHrs)", [ref]$sumEst)
            if ($null -ne $row.ActualTime) { $sumAct += [double]$row.ActualTime }
            if ($row.NeedsAudit) { $audit++ }
        }
        $diff = $sumAct - $sumEst
        $diffStr = if ($diff -ge 0) { "+{0:F2}" -f $diff } else { "{0:F2}" -f $diff }

        if ($script:wsStatusLabel) {
            $script:wsStatusLabel.Text = "Rows: $shown   Est: $("{0:F2}" -f $sumEst) h   Actual: $("{0:F2}" -f $sumAct) h   Diff: $diffStr h" +
                $(if ($audit -gt 0) { "   |   $audit NEEDS AUDIT" } else { '' })
        }
    } finally {
        $script:wsListView.EndUpdate()
    }
}

function Show-WorksheetPasteDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Paste eDAC Worksheet HTML"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Paste the raw HTML from the eDAC Weekly Assignment Worksheet (Ctrl-U, Ctrl-A, Ctrl-C)."
    $lblHead.Dock = 'Top'
    $lblHead.Height = 28
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblHead)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(680, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Anchor = 'Top,Right'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "Import"
    $okBtn.Location = New-Object System.Drawing.Point(780, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = 'Hand'
    $okBtn.Anchor = 'Top,Right'
    $okBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtPaste.Text)) { return }
        $form.Tag = $txtPaste.Text
        $form.Close()
    })
    $btnBar.Controls.Add($okBtn)

    $txtPaste = New-Object System.Windows.Forms.TextBox
    $txtPaste.Multiline = $true
    $txtPaste.MaxLength = 0
    $txtPaste.Dock = 'Fill'
    $txtPaste.ScrollBars = 'Both'
    $txtPaste.WordWrap = $false
    $txtPaste.Font = New-Object System.Drawing.Font("Consolas", 8)
    $txtPaste.AcceptsReturn = $true
    $txtPaste.Padding = New-Object System.Windows.Forms.Padding(8)
    $form.Controls.Add($txtPaste)
    $txtPaste.BringToFront()

    $form.CancelButton = $cancelBtn
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Setup-WeeklyWorksheetSubTab {
    param($parentTab)

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(0)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = 'Fill'
    $root.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $root.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($root)

    # Bottom status bar
    $statusBar = New-Object System.Windows.Forms.Panel
    $statusBar.Dock = 'Bottom'
    $statusBar.Height = 26
    $statusBar.BackColor = [System.Drawing.Color]::FromArgb(236,240,241)
    $statusBar.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)
    $root.Controls.Add($statusBar)

    $script:wsStatusLabel = New-Object System.Windows.Forms.Label
    $script:wsStatusLabel.Dock = 'Fill'
    $script:wsStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:wsStatusLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:wsStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $script:wsStatusLabel.Text = "No worksheet loaded."
    $statusBar.Controls.Add($script:wsStatusLabel)

    # Toolbar
    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $toolbar.Dock = 'Top'
    $toolbar.Height = 52
    $toolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $toolbar.WrapContents = $false
    $toolbar.BackColor = [System.Drawing.Color]::Transparent
    $toolbar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $root.Controls.Add($toolbar)

    $fetchBtn = New-Button "Fetch from eDAC" {
        $url = "$($script:config.WorkTracking.eDacUrl)".Trim()
        if ([string]::IsNullOrWhiteSpace($url)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No eDAC URL configured.`r`n`r`nSet WorkTracking.eDacUrl in the Settings tab, or use Paste HTML.",
                "No URL", "OK", "Warning")
            return
        }
        try {
            $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
            $summary = New-WorksheetFromHtml -Html $resp.Content -Source 'url'
            $script:wsWorksheet = Get-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date
            Refresh-WeeklyWorksheetList
            Write-Log "Weekly Worksheet: imported $($summary.DatesParsed) date(s), $($summary.TotalRows) total row(s)."
            $msg = "Imported $($summary.DatesParsed) date(s):`r`n" +
                   (($summary.Saved | ForEach-Object { "  $($_.Date) — $($_.Rows) row(s)" }) -join "`r`n")
            [System.Windows.Forms.MessageBox]::Show($msg, "Fetch Complete", "OK", "Information")
        } catch {
            Write-Log "Weekly Worksheet fetch failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Fetch failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
    } -Style 'Primary' -Width 160 -Height 32 -TextAlign MiddleCenter
    $fetchBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($fetchBtn)

    $pasteBtn = New-Button "Paste HTML" {
        $html = Show-WorksheetPasteDialog
        if ($null -eq $html) { return }
        try {
            $summary = New-WorksheetFromHtml -Html $html -Source 'paste'
            $script:wsWorksheet = Get-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date
            Refresh-WeeklyWorksheetList
            Write-Log "Weekly Worksheet: imported $($summary.DatesParsed) date(s), $($summary.TotalRows) total row(s)."
            $msg = "Imported $($summary.DatesParsed) date(s):`r`n" +
                   (($summary.Saved | ForEach-Object { "  $($_.Date) — $($_.Rows) row(s)" }) -join "`r`n")
            [System.Windows.Forms.MessageBox]::Show($msg, "Import Complete", "OK", "Information")
        } catch {
            Write-Log "Weekly Worksheet paste import failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Import failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
    } -Style 'Primary' -Width 130 -Height 32 -TextAlign MiddleCenter
    $pasteBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($pasteBtn)

    $syncBtn = New-Button "Sync Actual Times" {
        if ($null -eq $script:wsWorksheet) {
            [System.Windows.Forms.MessageBox]::Show("No worksheet loaded.", "Sync", "OK", "Information")
            return
        }
        Sync-WorksheetActualTimes -Worksheet $script:wsWorksheet -Date $script:wsDatePicker.Value.Date
        Add-WorksheetPlaceholders -Worksheet $script:wsWorksheet -Date $script:wsDatePicker.Value.Date
        Save-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date -Worksheet $script:wsWorksheet | Out-Null
        Refresh-WeeklyWorksheetList
        Write-Log "Weekly Worksheet: manual sync."
    } -Style 'Success' -Width 160 -Height 32 -TextAlign MiddleCenter
    $syncBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($syncBtn)

    $sendDayBtn = New-Button "Send Day" {
        $date = $script:wsDatePicker.Value.Date
        $ws = $script:wsWorksheet
        if ($null -eq $ws) {
            [System.Windows.Forms.MessageBox]::Show("No worksheet loaded for $($date.ToString('yyyy-MM-dd')).", "Send Day", "OK", "Warning")
            return
        }
        $auditCount = @(@($ws.Rows) | Where-Object { $_.NeedsAudit }).Count
        $confirm = if ($auditCount -gt 0) {
            "Send worksheet for $($date.ToString('yyyy-MM-dd')) to Historian?`r`n`r`n$auditCount row(s) still need audit and will be excluded from the summary but included in the file with an AUDIT flag."
        } else {
            "Send worksheet for $($date.ToString('yyyy-MM-dd')) to Historian?"
        }
        $answer = [System.Windows.Forms.MessageBox]::Show($confirm, "Send Day",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $evt = Send-WorksheetDayToHistorian -Date $date -Worksheet $ws
        if ($null -eq $evt) {
            [System.Windows.Forms.MessageBox]::Show("Send failed. See log.", "Send Day", "OK", "Error")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Sent to Historian.", "Send Day", "OK", "Information")
        }
    } -Style 'Success' -Width 110 -Height 32 -TextAlign MiddleCenter
    $sendDayBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($sendDayBtn)

    $reloadBtn = New-Button "Reload" {
        $script:wsWorksheet = Get-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date
        Refresh-WeeklyWorksheetList
    } -Style 'Ghost' -Width 90 -Height 32 -TextAlign MiddleCenter
    $reloadBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($reloadBtn)

    # Date picker
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Date:"
    $lblDate.Size = New-Object System.Drawing.Size(44, 32)
    $lblDate.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblDate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblDate.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblDate.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $toolbar.Controls.Add($lblDate)

    $script:wsDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $script:wsDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $script:wsDatePicker.Size = New-Object System.Drawing.Size(120, 26)
    $script:wsDatePicker.Value = (Get-Date).Date
    $script:wsDatePicker.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:wsDatePicker.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
    $script:wsDatePicker.Add_ValueChanged({
        $script:wsWorksheet = Get-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date
        Refresh-WeeklyWorksheetList
    })
    $toolbar.Controls.Add($script:wsDatePicker)

    $todayBtn = New-Button "Today" {
        $script:wsDatePicker.Value = (Get-Date).Date
    } -Style 'Ghost' -Width 70 -Height 32 -TextAlign MiddleCenter
    $todayBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($todayBtn)

    # ListView
    $script:wsListView = New-Object System.Windows.Forms.ListView
    $script:wsListView.Dock = 'Fill'
    $script:wsListView.Columns.Add("W/O #", 90)          | Out-Null
    $script:wsListView.Columns.Add("Acronym", 70)        | Out-Null
    $script:wsListView.Columns.Add("Class", 60)          | Out-Null
    $script:wsListView.Columns.Add("Equip", 60)          | Out-Null
    $script:wsListView.Columns.Add("Due", 85)            | Out-Null
    $script:wsListView.Columns.Add("PM Description", 220)| Out-Null
    $script:wsListView.Columns.Add("Description", 220)   | Out-Null
    $script:wsListView.Columns.Add("Est", 55)            | Out-Null
    $script:wsListView.Columns.Add("Act", 60)            | Out-Null
    $script:wsListView.Columns.Add("", 26)               | Out-Null
    $script:wsListView.Columns.Add("Flag", 60)           | Out-Null
    Set-ListViewStyle -ListView $script:wsListView
    $root.Controls.Add($script:wsListView)
    $script:wsListView.BringToFront()

    $script:wsListView.Add_DoubleClick({
        param($sender, $e)
        $sel = $sender.SelectedItems
        if ($sel.Count -eq 0) { return }
        $row = $sel[0].Tag
        if (-not $row) { return }

        $result = Show-WorksheetRowEditor -Row $row
        if ($null -eq $result) { return }

        if ($result.Action -eq 'clear') {
            $row.ActualTime     = $null
            $row.AutoValue      = $null
            $row.ManualOverride = $false
            $row.NeedsAudit     = $false
        } elseif ($result.Action -eq 'save') {
            $row.ActualTime     = [double]$result.Value
            $row.ManualOverride = $true
            $row.NeedsAudit     = $false
        }

        Save-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date -Worksheet $script:wsWorksheet | Out-Null
        Refresh-WeeklyWorksheetList
    })

    # Initial load
    $script:wsWorksheet = Get-WeeklyWorksheet -Date $script:wsDatePicker.Value.Date
    Refresh-WeeklyWorksheetList

    Write-Log "Weekly Worksheet sub-tab setup completed."
}

# ============================================================
# Work Tracking — Historian sub-tab
# ============================================================

function Get-HistorianKindLabel {
    param([string]$Kind)
    switch ("$Kind") {
        'reactive'      { return 'Reactive' }
        'workorder'     { return 'Work Order' }
        'pm.checklist'  { return 'PM Checklist' }
        'worksheet'     { return 'Worksheet' }
        default         { return "$Kind" }
    }
}

function Refresh-HistorianList {
    if (-not $script:histListView) { return }

    $script:histListView.BeginUpdate()
    try {
        $script:histListView.Items.Clear()

        $events = Get-HistorianEvents
        if ($null -eq $events) { $events = @() }

        # Apply filters
        $kindFilter    = if ($script:histKindFilter)    { "$($script:histKindFilter.SelectedItem)" } else { 'All' }
        $machineFilter = if ($script:histMachineFilter -and $script:histMachineFilter.SelectedItem) { "$($script:histMachineFilter.SelectedItem)" } else { '(any)' }
        $textFilter    = if ($script:histTextFilter)    { "$($script:histTextFilter.Text)".Trim().ToLower() } else { '' }
        $today         = (Get-Date).Date

        $shown = 0
        foreach ($e in $events) {
            if ($kindFilter -ne 'All') {
                if ("$($e.kind)" -ne $kindFilter) { continue }
            }
            if ($machineFilter -ne '(any)') {
                if ("$($e.machineId)" -ne $machineFilter) { continue }
            }
            if ($textFilter) {
                $hay = "$($e.summary) $($e.detail)".ToLower()
                if (-not $hay.Contains($textFilter)) { continue }
            }

            # Parse capturedAt for column display
            $when = ''
            $whenDate = [DateTime]::MinValue
            if ($e.capturedAt) {
                try { $whenDate = [DateTime]$e.capturedAt; $when = $whenDate.ToString("yyyy-MM-dd HH:mm") }
                catch { $when = "$($e.capturedAt)" }
            }

            $evtTime = ''
            if ($e.eventTime) {
                try { $evtTime = ([DateTime]$e.eventTime).ToString("MM-dd HH:mm") }
                catch { $evtTime = "$($e.eventTime)" }
            }

            $it = New-Object System.Windows.Forms.ListViewItem($when)
            $it.SubItems.Add((Get-HistorianKindLabel -Kind "$($e.kind)")) | Out-Null
            $it.SubItems.Add("$($e.machineId)") | Out-Null
            $it.SubItems.Add($evtTime)          | Out-Null
            $it.SubItems.Add("$($e.capturedBy)")| Out-Null
            $it.SubItems.Add("$($e.summary)")   | Out-Null
            $it.Tag = $e

            # Color: today's events get a subtle accent
            if ($whenDate -ne [DateTime]::MinValue -and $whenDate.Date -eq $today) {
                $it.UseItemStyleForSubItems = $false
                $it.SubItems[0].ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
                $it.SubItems[0].Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }

            $script:histListView.Items.Add($it) | Out-Null
            $shown++
        }

        if ($script:histStatusLabel) {
            $script:histStatusLabel.Text = "$shown of $(@($events).Count) event$(if (@($events).Count -ne 1) { 's' }) shown."
        }
    } finally {
        $script:histListView.EndUpdate()
    }
}

function Refresh-HistorianMachineFilter {
    if (-not $script:histMachineFilter) { return }

    $previous = ''
    if ($script:histMachineFilter.SelectedItem) {
        $previous = "$($script:histMachineFilter.SelectedItem)"
    }

    $script:histMachineFilter.Items.Clear()
    [void]$script:histMachineFilter.Items.Add('(any)')

    $kindFilter = 'All'
    if ($script:histKindFilter -and $script:histKindFilter.SelectedItem) {
        $kindFilter = "$($script:histKindFilter.SelectedItem)"
    }

    $machines = @{}
    $eventItems = ConvertTo-FlatArray -Source (Get-HistorianEvents)
    foreach ($e in @($eventItems)) {        if ($null -eq $e) { continue }
        if ($kindFilter -ne 'All' -and "$($e.kind)" -ne $kindFilter) { continue }
        $m = "$($e.machineId)".Trim()
        if ([string]::IsNullOrWhiteSpace($m)) { continue }
        $machines[$m] = $true
    }

    $sorted = @($machines.Keys | Sort-Object)
    Write-Host "[Historian] machine filter: kind='$kindFilter' -> $($sorted.Count) machine(s): $($sorted -join ', ')" -ForegroundColor Cyan

    foreach ($m in $sorted) {
        [void]$script:histMachineFilter.Items.Add($m)
    }

    if ($previous -and $script:histMachineFilter.Items.Contains($previous)) {
        $script:histMachineFilter.SelectedItem = $previous
    } else {
        $script:histMachineFilter.SelectedIndex = 0
    }
}

function Show-HistorianDebug {
    Write-Host ""
    Write-Host "===== HISTORIAN DEBUG =====" -ForegroundColor Cyan

    Write-Host "Kind dropdown:"
    if ($script:histKindFilter) {
        Write-Host "  Items.Count = $($script:histKindFilter.Items.Count)"
        for ($i = 0; $i -lt $script:histKindFilter.Items.Count; $i++) {
            $it = $script:histKindFilter.Items[$i]
            $tn = if ($null -eq $it) { '<null>' } else { $it.GetType().Name }
            Write-Host "    [$i] '$it'  (type=$tn)"
        }
        Write-Host "  SelectedIndex = $($script:histKindFilter.SelectedIndex)"
        Write-Host "  SelectedItem  = '$($script:histKindFilter.SelectedItem)'"
        Write-Host "  SelectedItem type = $($script:histKindFilter.SelectedItem.GetType().Name)"
    } else {
        Write-Host "  (script:histKindFilter is NULL)"
    }

    Write-Host "Machine dropdown:"
    if ($script:histMachineFilter) {
        Write-Host "  Items.Count = $($script:histMachineFilter.Items.Count)"
        for ($i = 0; $i -lt $script:histMachineFilter.Items.Count; $i++) {
            $it = $script:histMachineFilter.Items[$i]
            $tn = if ($null -eq $it) { '<null>' } else { $it.GetType().Name }
            Write-Host "    [$i] '$it'  (type=$tn)"
        }
        Write-Host "  SelectedIndex = $($script:histMachineFilter.SelectedIndex)"
        Write-Host "  SelectedItem  = '$($script:histMachineFilter.SelectedItem)'"
    } else {
        Write-Host "  (script:histMachineFilter is NULL)"
    }

    Write-Host "Raw events from Get-HistorianEvents:"
    $evs = Get-HistorianEvents
    Write-Host "  Total: $(@($evs).Count)"
    $i = 0
    foreach ($e in @($evs)) {
        $i++
        $k = "$($e.kind)"
        $m = "$($e.machineId)"
        Write-Host "  [$i] kind='$k'  machineId='$m'"
    }

    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-HistorianEventDetail {
    param($Event)

    if ($null -eq $Event) { return }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Historian Event"
    $form.Size = New-Object System.Drawing.Size(820, 620)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Size = New-Object System.Drawing.Size(100, 32)
    $closeBtn.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 112), 10)
    $closeBtn.Anchor = 'Top,Right'
    $closeBtn.FlatStyle = 'Flat'
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Cursor = 'Hand'
    $closeBtn.Add_Click({ $form.Close() })
    $btnBar.Controls.Add($closeBtn)

    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'
    $top.Height = 110
    $top.BackColor = [System.Drawing.Color]::White
    $top.Padding = New-Object System.Windows.Forms.Padding(14)
    $form.Controls.Add($top)

    $meta = New-Object System.Windows.Forms.Label
    $meta.Dock = 'Fill'
    $meta.Font = New-Object System.Drawing.Font("Consolas", 9)
    $meta.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $meta.Text = @(
        "Kind:        $(Get-HistorianKindLabel -Kind "$($Event.kind)")"
        "Machine:     $($Event.machineId)"
        "Captured:    $($Event.capturedAt)  by $($Event.capturedBy)"
        "Event time:  $($Event.eventTime)"
        "Summary:     $($Event.summary)"
    ) -join "`r`n"
    $top.Controls.Add($meta)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = 'Both'
    $txt.WordWrap = $false
    $txt.Dock = 'Fill'
    $txt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txt.BackColor = [System.Drawing.Color]::White
    $txt.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $txt.Text = "$($Event.detail)"
    $form.Controls.Add($txt)
    $txt.BringToFront()

    $form.CancelButton = $closeBtn
    $form.ShowDialog() | Out-Null
}

function Get-HistorianEventWorkDate {
    # Returns the [DateTime] the event is "about" (work date), or $null.
    param($Event)
    if ($null -eq $Event) { return $null }

    $kind = "$($Event.kind)"

    if ($kind -eq 'worksheet' -and $Event.payload -and $Event.payload.date) {
        try { return ([DateTime]$Event.payload.date).Date } catch {}
    }
    if ($kind -eq 'pm.checklist' -and $Event.payload -and $Event.payload.checklistDate) {
        try { return ([DateTime]$Event.payload.checklistDate).Date } catch {}
    }
    if ($Event.eventTime) {
        try { return ([DateTime]$Event.eventTime).Date } catch {}
    }
    if ($Event.capturedAt) {
        try { return ([DateTime]$Event.capturedAt).Date } catch {}
    }
    return $null
}

function Format-HistorianTableRow {
    param(
        [string[]]$Cells,
        [int[]]$Widths,
        [string[]]$Aligns     # 'L' or 'R' per column
    )
    $parts = @()
    for ($i = 0; $i -lt $Cells.Count; $i++) {
        $c = "$($Cells[$i])"
        $w = $Widths[$i]
        if ($c.Length -gt $w) { $c = $c.Substring(0, [Math]::Max(0, $w - 3)) + '...' }
        if ($Aligns[$i] -eq 'R') { $parts += $c.PadLeft($w) }
        else                     { $parts += $c.PadRight($w) }
    }
    return ('  ' + ($parts -join '  '))
}

function Format-HistorianReport {
    param(
        [DateTime]$From,
        [DateTime]$To,
        [hashtable]$Sections,   # @{ Jobs=bool; Pm=bool; Worksheet=bool; Budget=bool }
        [array]$Events          # pre-filtered to range by caller
    )

    $W    = 100
    $sb   = New-Object System.Text.StringBuilder
    $eq   = '=' * $W
    $add  = { param($s) [void]$sb.AppendLine("$s") }

    # ---- Header ----
    & $add $eq
    $title = 'WORK TRACKING REPORT'
    $padL  = [Math]::Floor(($W - $title.Length) / 2)
    $padR  = $W - $title.Length - $padL
    & $add ((' ' * $padL) + $title + (' ' * $padR))
    & $add $eq

    $tech = "$($script:config.WorkTracking.TechnicianName)".Trim()
    if ([string]::IsNullOrWhiteSpace($tech)) { $tech = '(not set)' }

    & $add ("  Range:      {0}  to  {1}" -f $From.ToString('yyyy-MM-dd'), $To.ToString('yyyy-MM-dd'))
    & $add ("  Technician: {0}" -f $tech)
    & $add ("  Generated:  {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    & $add ("  Events:     {0}" -f @($Events).Count)
    & $add $eq
    & $add ''

    # ---------- [1] JOBS ----------
    if ($Sections.Jobs) {
        $jobEvents = @($Events | Where-Object { "$($_.kind)" -in @('reactive','workorder') })
        $reactive  = @($jobEvents | Where-Object { "$($_.kind)" -eq 'reactive' })
        $wos       = @($jobEvents | Where-Object { "$($_.kind)" -eq 'workorder' })

        & $add '[1] JOBS'
        & $add $eq
        & $add ''

        $widths = @(10, 12, 5, 5, 5, 8, 3, 38)
        $aligns = @('L','L','L','L','R','L','R','L')
        $head   = @('Date','Machine','Start','End','Hrs','WO#','Prt','Description')

        foreach ($pair in @(
            @{ Label = 'REACTIVE CALLS'; Items = $reactive }
            @{ Label = 'WORK ORDERS';    Items = $wos      }
        )) {
            if (@($pair.Items).Count -eq 0) { continue }
            & $add ("  $($pair.Label) ($(@($pair.Items).Count))")
            & $add ''
            & $add (Format-HistorianTableRow -Cells $head -Widths $widths -Aligns $aligns)
            & $add ('  ' + (($widths | ForEach-Object { '-' * $_ }) -join '  '))

            $sub = 0.0
            foreach ($e in $pair.Items) {
                $job  = if ($e.payload) { $e.payload.job } else { $null }
                $h    = if ($job) { Get-JobHours -Job $job } else { $null }
                if ($null -eq $h) { $h = 0 }
                $sub += $h

                $pdate = Get-HistorianEventWorkDate -Event $e
                $pdateStr = if ($pdate) { $pdate.ToString('yyyy-MM-dd') } else { '' }
                $prtCount = if ($job -and $job.parts) { @($job.parts).Count } else { 0 }

                & $add (Format-HistorianTableRow -Cells @(
                    $pdateStr,
                    "$($e.machineId)",
                    "$($job.startTime)",
                    "$($job.endTime)",
                    ("{0:F2}" -f $h),
                    "$($job.workOrderNo)",
                    "$prtCount",
                    "$($job.description)"
                ) -Widths $widths -Aligns $aligns)
            }
            & $add ''
            & $add ("  Subtotal:  {0} item{1}  ·  {2:F2} h" -f @($pair.Items).Count, $(if (@($pair.Items).Count -ne 1) { 's' } else { '' }), $sub)
            & $add ''
        }

        $reactiveH = 0.0
        foreach ($e in $reactive) {
            $h = Get-JobHours -Job $e.payload.job
            if ($null -ne $h) { $reactiveH += $h }
        }
        $woH = 0.0
        foreach ($e in $wos) {
            $h = Get-JobHours -Job $e.payload.job
            if ($null -ne $h) { $woH += $h }
        }
        & $add ("  JOBS TOTAL:  {0} item{1}  ·  {2:F2} h" -f $jobEvents.Count, $(if ($jobEvents.Count -ne 1) { 's' } else { '' }), ($reactiveH + $woH))
        & $add ''
    }

    # ---------- [2] PM CHECKLISTS ----------
    if ($Sections.Pm) {
        $pmEvents = @($Events | Where-Object { "$($_.kind)" -eq 'pm.checklist' })

        & $add '[2] PM CHECKLISTS'
        & $add $eq
        & $add ''

        if ($pmEvents.Count -eq 0) {
            & $add '  (none)'
            & $add ''
        } else {
            $widths = @(10, 12, 11, 7, 7)
            $aligns = @('L','L','L','R','R')
            $head   = @('Date','Machine','Checklist','Tasks','Hours')

            & $add (Format-HistorianTableRow -Cells $head -Widths $widths -Aligns $aligns)
            & $add ('  ' + (($widths | ForEach-Object { '-' * $_ }) -join '  '))

            $sub = 0.0
            foreach ($e in $pmEvents) {
                $p = $e.payload
                $dt = Get-HistorianEventWorkDate -Event $e
                $dtStr = if ($dt) { $dt.ToString('yyyy-MM-dd') } else { '' }
                $taskCount = if ($p -and $p.selectedTasks) { @($p.selectedTasks).Count } else { 0 }
                $hrs = if ($p -and $null -ne $p.totalHours) { [double]$p.totalHours } else { 0 }
                $sub += $hrs

                & $add (Format-HistorianTableRow -Cells @(
                    $dtStr, "$($e.machineId)", "$($p.checklistNo)", "$taskCount", ("{0:F2}" -f $hrs)
                ) -Widths $widths -Aligns $aligns)
            }
            & $add ''
            & $add ("  Subtotal:  {0} checklist{1}  ·  {2:F2} h" -f $pmEvents.Count, $(if ($pmEvents.Count -ne 1) { 's' } else { '' }), $sub)
            & $add ''
        }
    }

    # ---------- [3] WEEKLY WORKSHEETS ----------
    if ($Sections.Worksheet) {
        $wsEvents = @($Events | Where-Object { "$($_.kind)" -eq 'worksheet' })

        & $add '[3] WEEKLY WORKSHEETS'
        & $add $eq
        & $add ''

        if ($wsEvents.Count -eq 0) {
            & $add '  (none)'
            & $add ''
        } else {
            $widths = @(10, 6, 8, 8, 8, 6)
            $aligns = @('L','R','R','R','R','R')
            $head   = @('Date','Rows','Est','Act','Diff','Audit')

            & $add (Format-HistorianTableRow -Cells $head -Widths $widths -Aligns $aligns)
            & $add ('  ' + (($widths | ForEach-Object { '-' * $_ }) -join '  '))

            $sumE = 0.0; $sumA = 0.0
            foreach ($e in $wsEvents) {
                $p = $e.payload
                $dt = "$($p.date)"
                $diff = [double]$p.actHours - [double]$p.estHours
                $sumE += [double]$p.estHours
                $sumA += [double]$p.actHours
                & $add (Format-HistorianTableRow -Cells @(
                    $dt, "$($p.rowCount)",
                    ("{0:F2}" -f [double]$p.estHours),
                    ("{0:F2}" -f [double]$p.actHours),
                    ("{0}{1:F2}" -f $(if ($diff -ge 0) { '+' } else { '' }), $diff),
                    "$($p.auditCount)"
                ) -Widths $widths -Aligns $aligns)
            }
            & $add ''
            & $add ("  Subtotal:  {0} day{1}  ·  est {2:F2} h  ·  act {3:F2} h" -f $wsEvents.Count, $(if ($wsEvents.Count -ne 1) { 's' } else { '' }), $sumE, $sumA)
            & $add ''
        }
    }

    # ---------- [4] WORK BUDGET ROLL-UP ----------
    if ($Sections.Budget) {
        & $add '[4] WORK BUDGET ROLL-UP'
        & $add $eq
        & $add ''

        $wb = $script:config.WorkTracking.WorkBudget
        $target = if ($wb -and $wb.WorkTargetHours) { [double]$wb.WorkTargetHours } else { 6.5 }

        # Bucket job hours by date from source Jobs.json
        $jobsByDay = @{}
        $jobItems = ConvertTo-FlatArray -Source (Get-Jobs)
        foreach ($job in @($jobItems)) {
            if ($null -eq $job) { continue }
            $jd = ''
            if ($job.createdAt) {
                try { $jd = ([DateTime]$job.createdAt).Date.ToString('yyyy-MM-dd') }
                catch { $jd = "$($job.createdAt)".Split('T')[0] }
            }
            if ([string]::IsNullOrWhiteSpace($jd)) { continue }
            if (-not $jobsByDay.ContainsKey($jd)) { $jobsByDay[$jd] = 0.0 }
            $h = Get-JobHours -Job $job
            if ($null -ne $h) { $jobsByDay[$jd] += $h }
        }

        # Bucket PM hours by date from historian events
        $pmByDay = @{}
        foreach ($e in @($Events | Where-Object { "$($_.kind)" -eq 'pm.checklist' })) {
            $d = Get-HistorianEventWorkDate -Event $e
            if (-not $d) { continue }
            $key = $d.ToString('yyyy-MM-dd')
            if (-not $pmByDay.ContainsKey($key)) { $pmByDay[$key] = 0.0 }
            if ($e.payload -and $null -ne $e.payload.totalHours) {
                $pmByDay[$key] += [double]$e.payload.totalHours
            }
        }

        # Bucket breaks by date
        $breaksByDay = @{}
        $breakItems = ConvertTo-FlatArray -Source (Get-Breaks)
        foreach ($b in @($breakItems)) {
            if ($null -eq $b) { continue }
            $d = "$($b.date)"
            if ([string]::IsNullOrWhiteSpace($d)) { continue }

            $dm = $b.durationMin
            while ($dm -is [System.Array]) {
                if ($dm.Count -eq 0) { $dm = 0; break }
                $dm = $dm[0]
            }

            if (-not $breaksByDay.ContainsKey($d)) { $breaksByDay[$d] = 0 }
            $breaksByDay[$d] += [int]$dm
        }

        # Collect only days with activity in the range
        $activeDays = New-Object System.Collections.ArrayList
        $d = $From.Date
        while ($d -le $To.Date) {
            $key = $d.ToString('yyyy-MM-dd')
            $jH = if ($jobsByDay.ContainsKey($key))   { [Math]::Round($jobsByDay[$key], 2) }   else { 0.0 }
            $pH = if ($pmByDay.ContainsKey($key))     { [Math]::Round($pmByDay[$key], 2) }     else { 0.0 }
            $bM = if ($breaksByDay.ContainsKey($key)) { [int]$breaksByDay[$key] }              else { 0 }
            if ($jH -gt 0 -or $pH -gt 0 -or $bM -gt 0) {
                [void]$activeDays.Add([PSCustomObject]@{
                    Date       = $key
                    Jobs       = $jH
                    Pm         = $pH
                    BreakMin   = $bM
                    Work       = [Math]::Round($jH + $pH, 2)
                })
            }
            $d = $d.AddDays(1)
        }

        if ($activeDays.Count -eq 0) {
            & $add '  (no tracked work in this range)'
            & $add ''
        } else {
            $widths = @(10, 8, 7, 7, 7, 7, 7, 8)
            $aligns = @('L','R','R','R','R','R','R','L')
            $head   = @('Date','Jobs','PM','Work','Break','Target','Diff','Status')

            & $add (Format-HistorianTableRow -Cells $head -Widths $widths -Aligns $aligns)
            & $add ('  ' + (($widths | ForEach-Object { '-' * $_ }) -join '  '))

            $rangeWork = 0.0
            $rangeTarget = 0.0
            $rangeBreak = 0
            foreach ($row in $activeDays) {
                $diff = [Math]::Round($row.Work - $target, 2)
                $status = if ($diff -gt 0) { 'OVER' } else { 'OK' }
                $rangeWork   += $row.Work
                $rangeTarget += $target
                $rangeBreak  += $row.BreakMin

                & $add (Format-HistorianTableRow -Cells @(
                    $row.Date,
                    ("{0:F2}" -f $row.Jobs),
                    ("{0:F2}" -f $row.Pm),
                    ("{0:F2}" -f $row.Work),
                    ("{0:F2}" -f ($row.BreakMin / 60.0)),
                    ("{0:F2}" -f $target),
                    ("{0}{1:F2}" -f $(if ($diff -ge 0) { '+' } else { '' }), $diff),
                    $status
                ) -Widths $widths -Aligns $aligns)
            }

            & $add ''
            & $add ("  ACTIVE DAYS:  {0}   WORK: {1:F2} h   TARGET: {2:F2} h   DIFF: {3}{4:F2} h" -f `
                $activeDays.Count, $rangeWork, $rangeTarget,
                $(if (($rangeWork - $rangeTarget) -ge 0) { '+' } else { '' }),
                ($rangeWork - $rangeTarget))
            & $add ("  BREAKS LOGGED: {0:F2} h" -f ($rangeBreak / 60.0))
            & $add ''
        }
    }

    # ---- Footer ----
    & $add $eq
    $end = 'END OF REPORT'
    $padL = [Math]::Floor(($W - $end.Length) / 2)
    $padR = $W - $end.Length - $padL
    & $add ((' ' * $padL) + $end + (' ' * $padR))
    & $add $eq

    return $sb.ToString()
}

function Get-HistorianReportDirectory {
    $rel = $script:config.WorkTracking.ReportsDirectory
    if ([string]::IsNullOrWhiteSpace($rel)) { return (Join-Path $script:config.RootDirectory 'Work Tracking\Reports') }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    return (Join-Path $script:config.RootDirectory $rel)
}

function Show-HistorianReportDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Generate Report'
    $form.Size = New-Object System.Drawing.Size(1000, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $true
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ---------- bottom bar ----------
    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Size = New-Object System.Drawing.Size(90, 32)
    $closeBtn.Location = New-Object System.Drawing.Point(880, 10)
    $closeBtn.Anchor = 'Top,Right'
    $closeBtn.FlatStyle = 'Flat'
    $closeBtn.FlatAppearance.BorderSize = 0
    $closeBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $closeBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $closeBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeBtn.Cursor = 'Hand'
    $closeBtn.Add_Click({ $form.Close() })
    $btnBar.Controls.Add($closeBtn)

    $copyBtn = New-Object System.Windows.Forms.Button
    $copyBtn.Text = "Copy to Clipboard"
    $copyBtn.Size = New-Object System.Drawing.Size(150, 32)
    $copyBtn.Location = New-Object System.Drawing.Point(710, 10)
    $copyBtn.Anchor = 'Top,Right'
    $copyBtn.FlatStyle = 'Flat'
    $copyBtn.FlatAppearance.BorderSize = 0
    $copyBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $copyBtn.ForeColor = [System.Drawing.Color]::White
    $copyBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $copyBtn.Cursor = 'Hand'
    $btnBar.Controls.Add($copyBtn)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Save to File"
    $saveBtn.Size = New-Object System.Drawing.Size(140, 32)
    $saveBtn.Location = New-Object System.Drawing.Point(560, 10)
    $saveBtn.Anchor = 'Top,Right'
    $saveBtn.FlatStyle = 'Flat'
    $saveBtn.FlatAppearance.BorderSize = 0
    $saveBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $saveBtn.ForeColor = [System.Drawing.Color]::White
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.Cursor = 'Hand'
    $btnBar.Controls.Add($saveBtn)

    # ---------- top options ----------
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'
    $top.Height = 90
    $top.BackColor = [System.Drawing.Color]::White
    $top.Padding = New-Object System.Windows.Forms.Padding(12)
    $form.Controls.Add($top)

    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From:"
    $lblFrom.Location = New-Object System.Drawing.Point(12, 22)
    $lblFrom.Size = New-Object System.Drawing.Size(46, 22)
    $lblFrom.TextAlign = 'MiddleRight'
    $lblFrom.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $top.Controls.Add($lblFrom)

    $dtpFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtpFrom.Format = 'Short'
    $dtpFrom.Location = New-Object System.Drawing.Point(64, 20)
    $dtpFrom.Size = New-Object System.Drawing.Size(120, 24)
    $dtpFrom.Value = (Get-Date).Date.AddDays(-6)
    $top.Controls.Add($dtpFrom)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To:"
    $lblTo.Location = New-Object System.Drawing.Point(200, 22)
    $lblTo.Size = New-Object System.Drawing.Size(30, 22)
    $lblTo.TextAlign = 'MiddleRight'
    $lblTo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $top.Controls.Add($lblTo)

    $dtpTo = New-Object System.Windows.Forms.DateTimePicker
    $dtpTo.Format = 'Short'
    $dtpTo.Location = New-Object System.Drawing.Point(236, 20)
    $dtpTo.Size = New-Object System.Drawing.Size(120, 24)
    $dtpTo.Value = (Get-Date).Date
    $top.Controls.Add($dtpTo)

    $lblSections = New-Object System.Windows.Forms.Label
    $lblSections.Text = "Sections:"
    $lblSections.Location = New-Object System.Drawing.Point(380, 22)
    $lblSections.Size = New-Object System.Drawing.Size(70, 22)
    $lblSections.TextAlign = 'MiddleRight'
    $lblSections.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $top.Controls.Add($lblSections)

    $cbJobs = New-Object System.Windows.Forms.CheckBox
    $cbJobs.Text = "Jobs"; $cbJobs.Checked = $true
    $cbJobs.Location = New-Object System.Drawing.Point(456, 20)
    $cbJobs.Size = New-Object System.Drawing.Size(60, 24)
    $top.Controls.Add($cbJobs)

    $cbPm = New-Object System.Windows.Forms.CheckBox
    $cbPm.Text = "PM"; $cbPm.Checked = $true
    $cbPm.Location = New-Object System.Drawing.Point(520, 20)
    $cbPm.Size = New-Object System.Drawing.Size(60, 24)
    $top.Controls.Add($cbPm)

    $cbWs = New-Object System.Windows.Forms.CheckBox
    $cbWs.Text = "Worksheet"; $cbWs.Checked = $true
    $cbWs.Location = New-Object System.Drawing.Point(580, 20)
    $cbWs.Size = New-Object System.Drawing.Size(100, 24)
    $top.Controls.Add($cbWs)

    $cbBudget = New-Object System.Windows.Forms.CheckBox
    $cbBudget.Text = "Budget"; $cbBudget.Checked = $true
    $cbBudget.Location = New-Object System.Drawing.Point(690, 20)
    $cbBudget.Size = New-Object System.Drawing.Size(80, 24)
    $top.Controls.Add($cbBudget)

    $generateBtn = New-Object System.Windows.Forms.Button
    $generateBtn.Text = "Generate"
    $generateBtn.Location = New-Object System.Drawing.Point(64, 52)
    $generateBtn.Size = New-Object System.Drawing.Size(120, 28)
    $generateBtn.FlatStyle = 'Flat'
    $generateBtn.FlatAppearance.BorderSize = 0
    $generateBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $generateBtn.ForeColor = [System.Drawing.Color]::White
    $generateBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $generateBtn.Cursor = 'Hand'
    $top.Controls.Add($generateBtn)

    # ---------- preview ----------
    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Multiline = $true
    $txtPreview.ReadOnly = $true
    $txtPreview.ScrollBars = 'Both'
    $txtPreview.WordWrap = $false
    $txtPreview.MaxLength = 0
    $txtPreview.Dock = 'Fill'
    $txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtPreview.BackColor = [System.Drawing.Color]::White
    $txtPreview.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $txtPreview.Text = "(Click Generate)"
    $form.Controls.Add($txtPreview)
    $txtPreview.BringToFront()

    # ---------- state ----------
    $script:__hr_lastText = ''
    $script:__hr_lastRange = ''

    $doGenerate = {
        $from = $dtpFrom.Value.Date
        $to   = $dtpTo.Value.Date
        if ($to -lt $from) {
            [System.Windows.Forms.MessageBox]::Show("'To' must be on or after 'From'.", "Generate", "OK", "Warning")
            return
        }

        # Filter events to range
        $all = Get-HistorianEvents
        $filtered = @()
        foreach ($e in @($all)) {
            $d = Get-HistorianEventWorkDate -Event $e
            if ($null -eq $d) { continue }
            if ($d -ge $from -and $d -le $to) { $filtered += $e }
        }

        $sections = @{
            Jobs      = $cbJobs.Checked
            Pm        = $cbPm.Checked
            Worksheet = $cbWs.Checked
            Budget    = $cbBudget.Checked
        }

        $text = Format-HistorianReport -From $from -To $to -Sections $sections -Events $filtered
        $txtPreview.Text = $text
        $script:__hr_lastText = $text
        $script:__hr_lastRange = "$($from.ToString('yyyy-MM-dd'))_to_$($to.ToString('yyyy-MM-dd'))"
    }

    $generateBtn.Add_Click($doGenerate)

    $copyBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:__hr_lastText)) {
            [System.Windows.Forms.MessageBox]::Show("Generate a report first.", "Copy", "OK", "Warning")
            return
        }
        try {
            [System.Windows.Forms.Clipboard]::SetText($script:__hr_lastText)
            [System.Windows.Forms.MessageBox]::Show("Report copied to clipboard.", "Copy", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Copy failed: $($_.Exception.Message)", "Copy", "OK", "Error")
        }
    })

    $saveBtn.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:__hr_lastText)) {
            [System.Windows.Forms.MessageBox]::Show("Generate a report first.", "Save", "OK", "Warning")
            return
        }
        $dir = Get-HistorianReportDirectory
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $defaultName = "Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "Text files (*.txt)|*.txt"
        $dlg.InitialDirectory = $dir
        $dlg.FileName = $defaultName
        $dlg.Title = "Save Report"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $script:__hr_lastText | Set-Content -Path $dlg.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Saved:`r`n$($dlg.FileName)", "Save", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Save failed: $($_.Exception.Message)", "Save", "OK", "Error")
        }
    })

    $form.CancelButton = $closeBtn

    $form.Add_Shown({ & $doGenerate })

    $form.ShowDialog() | Out-Null
}

function Setup-HistorianSubTab {
    param($parentTab)

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(0)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = 'Fill'
    $root.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $root.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($root)

    # Status bar
    $statusBar = New-Object System.Windows.Forms.Panel
    $statusBar.Dock = 'Bottom'
    $statusBar.Height = 26
    $statusBar.BackColor = [System.Drawing.Color]::FromArgb(236,240,241)
    $statusBar.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)
    $root.Controls.Add($statusBar)

    $script:histStatusLabel = New-Object System.Windows.Forms.Label
    $script:histStatusLabel.Dock = 'Fill'
    $script:histStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:histStatusLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:histStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $script:histStatusLabel.Text = "0 events"
    $statusBar.Controls.Add($script:histStatusLabel)

    # Toolbar
    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $toolbar.Dock = 'Top'
    $toolbar.Height = 52
    $toolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $toolbar.WrapContents = $false
    $toolbar.BackColor = [System.Drawing.Color]::Transparent
    $toolbar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $root.Controls.Add($toolbar)

    $refreshBtn = New-Button "Refresh" {
        Refresh-HistorianMachineFilter
        Refresh-HistorianList
    } -Style 'Primary' -Width 100 -Height 32 -TextAlign MiddleCenter
    $refreshBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($refreshBtn)

    $openFileBtn = New-Button "Open Historian File" {
        $p = Get-HistorianFilePath
        if (Test-Path $p) { Start-Process $p }
        else { [System.Windows.Forms.MessageBox]::Show("Historian file not found.", "Historian", "OK", "Warning") }
    } -Style 'Ghost' -Width 160 -Height 32 -TextAlign MiddleCenter
    $openFileBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($openFileBtn)

    $lblKind = New-Object System.Windows.Forms.Label
    $lblKind.Text = "Kind:"
    $lblKind.Size = New-Object System.Drawing.Size(40, 32)
    $lblKind.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblKind.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblKind.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblKind.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $toolbar.Controls.Add($lblKind)

    $script:histKindFilter = New-Object System.Windows.Forms.ComboBox
    $script:histKindFilter.DropDownStyle = 'DropDownList'
    $script:histKindFilter.Size = New-Object System.Drawing.Size(130, 26)
    $script:histKindFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:histKindFilter.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
    foreach ($k in @('All','reactive','workorder','pm.checklist','worksheet')) {
        [void]$script:histKindFilter.Items.Add($k)
    }
    $script:histKindFilter.SelectedIndex = 0
    $script:histKindFilter.Add_SelectedIndexChanged({
        Refresh-HistorianMachineFilter
        Refresh-HistorianList
    })
    $toolbar.Controls.Add($script:histKindFilter)

    $lblMachine = New-Object System.Windows.Forms.Label
    $lblMachine.Text = "Machine:"
    $lblMachine.Size = New-Object System.Drawing.Size(66, 32)
    $lblMachine.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblMachine.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMachine.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblMachine.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $toolbar.Controls.Add($lblMachine)

    $script:histMachineFilter = New-Object System.Windows.Forms.ComboBox
    $script:histMachineFilter.DropDownStyle = 'DropDownList'
    $script:histMachineFilter.Size = New-Object System.Drawing.Size(140, 26)
    $script:histMachineFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:histMachineFilter.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
    [void]$script:histMachineFilter.Items.Add('(any)')
    $script:histMachineFilter.SelectedIndex = 0
    $script:histMachineFilter.Add_SelectedIndexChanged({ Refresh-HistorianList })
    $toolbar.Controls.Add($script:histMachineFilter)

    $lblText = New-Object System.Windows.Forms.Label
    $lblText.Text = "Search:"
    $lblText.Size = New-Object System.Drawing.Size(56, 32)
    $lblText.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblText.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblText.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblText.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $toolbar.Controls.Add($lblText)

    $script:histTextFilter = New-Object System.Windows.Forms.TextBox
    $script:histTextFilter.Size = New-Object System.Drawing.Size(220, 26)
    $script:histTextFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:histTextFilter.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
    $script:histTextFilter.Add_TextChanged({ Refresh-HistorianList })
    $toolbar.Controls.Add($script:histTextFilter)

    $debugBtn = New-Button "Debug Dump" {
        Show-HistorianDebug
    } -Style 'Ghost' -Width 110 -Height 32 -TextAlign MiddleCenter
    $debugBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($debugBtn)

    $reportBtn = New-Button "Generate Report" {
        Show-HistorianReportDialog
    } -Style 'Success' -Width 140 -Height 32 -TextAlign MiddleCenter
    $reportBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($reportBtn)

    $clearBtn = New-Button "Clear Filters" {
        $script:histKindFilter.SelectedIndex = 0
        $script:histMachineFilter.SelectedIndex = 0
        $script:histTextFilter.Text = ''
        Refresh-HistorianList
    } -Style 'Ghost' -Width 110 -Height 32 -TextAlign MiddleCenter
    $clearBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $toolbar.Controls.Add($clearBtn)

    # ListView
    $script:histListView = New-Object System.Windows.Forms.ListView
    $script:histListView.Dock = 'Fill'
    $script:histListView.Columns.Add("Captured", 140)    | Out-Null
    $script:histListView.Columns.Add("Kind", 110)        | Out-Null
    $script:histListView.Columns.Add("Machine", 110)     | Out-Null
    $script:histListView.Columns.Add("Event Time", 110)  | Out-Null
    $script:histListView.Columns.Add("By", 100)          | Out-Null
    $script:histListView.Columns.Add("Summary", 480)     | Out-Null
    Set-ListViewStyle -ListView $script:histListView
    $root.Controls.Add($script:histListView)
    $script:histListView.BringToFront()

    $script:histListView.Add_DoubleClick({
        param($sender, $e)
        $sel = $sender.SelectedItems
        if ($sel.Count -eq 0) { return }
        Show-HistorianEventDetail -Event $sel[0].Tag
    })

    # Initial load
    Refresh-HistorianMachineFilter
    Refresh-HistorianList

    Write-Log "Historian sub-tab setup completed."
}

function New-JobObject {
    param([string]$Kind)
    return [PSCustomObject]@{
        id              = [guid]::NewGuid().ToString()
        kind            = $Kind
        machineId       = ''
        startTime       = ''
        endTime         = ''
        hoursOverride   = $null
        description     = ''
        workOrderNo     = ''
        parts           = @()
        createdAt       = (Get-Date).ToString("o")
        sentToHistorian = $false
        sentAt          = $null
    }
}

function Test-JobFields {
    # Returns a string[] of validation errors (empty array = valid).
    param($Job)

    $errs = @()
    if ($null -eq $Job) { return @('No job to validate.') }

    $m = Format-MachineId "$($Job.machineId)"
    if ($null -eq $m) { $errs += 'Machine ID must look like "DBCS 49".' }

    if ([string]::IsNullOrWhiteSpace($Job.startTime)) { $errs += 'Start time is required.' }
    if ([string]::IsNullOrWhiteSpace($Job.endTime))   { $errs += 'End time is required.' }

    $h = Get-JobHours $Job
    if ($null -ne $h -and $h -le 0) { $errs += 'End time must be after start time.' }

    $needsDetail = ($null -ne $h -and $h -gt 0.5) -or (@($Job.parts).Count -gt 0)
    if ($needsDetail) {
        if ([string]::IsNullOrWhiteSpace($Job.description)) { $errs += 'Description required (>30 min or parts used).' }
        if ([string]::IsNullOrWhiteSpace($Job.workOrderNo)) { $errs += 'Work order # required (>30 min or parts used).' }
    }

    return ,$errs
}

function Show-JobEditor {
    # Returns:
    #   $null                              - cancelled
    #   [PSCustomObject]@{Action='save';   Job=$job} - save this job
    #   [PSCustomObject]@{Action='delete'; Job=$job} - delete this job
    param(
        [PSCustomObject]$Job,
        [string]$Kind
    )

    if ($null -eq $Job) { $Job = New-JobObject -Kind $Kind }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($Job.kind -eq 'workorder') { 'Work Order' } else { 'Reactive Call' }
    $form.Size = New-Object System.Drawing.Size(720, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 620)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ---------- bottom button bar ----------
    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $btnBar.BackColor = [System.Drawing.Color]::Transparent
    $btnBar.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 12)
    $form.Controls.Add($btnBar)

    $deleteBtn = New-Object System.Windows.Forms.Button
    $deleteBtn.Text = "Delete"
    $deleteBtn.Location = New-Object System.Drawing.Point(12, 10)
    $deleteBtn.Size = New-Object System.Drawing.Size(90, 32)
    $deleteBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $deleteBtn.ForeColor = [System.Drawing.Color]::White
    $deleteBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $deleteBtn.FlatAppearance.BorderSize = 0
    $deleteBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $deleteBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    if (-not [string]::IsNullOrWhiteSpace($Job.id) -and (Get-Jobs | Where-Object { $_.id -eq $Job.id })) {
        $btnBar.Controls.Add($deleteBtn)
    }

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(480, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(100, 32)
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnBar.Controls.Add($cancelBtn)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Save"
    $saveBtn.Location = New-Object System.Drawing.Point(590, 10)
    $saveBtn.Size = New-Object System.Drawing.Size(110, 32)
    $saveBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $saveBtn.ForeColor = [System.Drawing.Color]::White
    $saveBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveBtn.FlatAppearance.BorderSize = 0
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $saveBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnBar.Controls.Add($saveBtn)

    # ---------- top fields panel ----------
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Top'
    $top.Height = 340
    $top.BackColor = [System.Drawing.Color]::White
    $top.Padding = New-Object System.Windows.Forms.Padding(16)
    $form.Controls.Add($top)
    $top.BringToFront()

    $mklabel = {
        param($text, $y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point(16, ($y + 4))
        $l.Size = New-Object System.Drawing.Size(120, 22)
        $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
        return $l
    }

    # Machine ID
    $top.Controls.Add((& $mklabel "Machine ID:" 20))
    $txtMachine = New-Object System.Windows.Forms.TextBox
    $txtMachine.Location = New-Object System.Drawing.Point(144, 20)
    $txtMachine.Size = New-Object System.Drawing.Size(180, 24)
    $txtMachine.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtMachine.Text = "$($Job.machineId)"
    $top.Controls.Add($txtMachine)

    $lblParsed = New-Object System.Windows.Forms.Label
    $lblParsed.Location = New-Object System.Drawing.Point(334, 24)
    $lblParsed.Size = New-Object System.Drawing.Size(320, 22)
    $lblParsed.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblParsed.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $top.Controls.Add($lblParsed)

    # Start / End
    $top.Controls.Add((& $mklabel "Start:" 60))
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpStart.CustomFormat = "HH:mm"
    $dtpStart.ShowUpDown = $true
    $dtpStart.Location = New-Object System.Drawing.Point(144, 60)
    $dtpStart.Size = New-Object System.Drawing.Size(90, 24)
    $top.Controls.Add($dtpStart)

    $btnStartNow = New-Object System.Windows.Forms.Button
    $btnStartNow.Text = "Now"
    $btnStartNow.Location = New-Object System.Drawing.Point(240, 60)
    $btnStartNow.Size = New-Object System.Drawing.Size(50, 24)
    $btnStartNow.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStartNow.FlatAppearance.BorderSize = 0
    $btnStartNow.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnStartNow.ForeColor = [System.Drawing.Color]::White
    $btnStartNow.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnStartNow.Cursor = [System.Windows.Forms.Cursors]::Hand
    $top.Controls.Add($btnStartNow)

    $top.Controls.Add((& $mklabel "End:" 100))
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpEnd.CustomFormat = "HH:mm"
    $dtpEnd.ShowUpDown = $true
    $dtpEnd.Location = New-Object System.Drawing.Point(144, 100)
    $dtpEnd.Size = New-Object System.Drawing.Size(90, 24)
    $top.Controls.Add($dtpEnd)

    $btnEndNow = New-Object System.Windows.Forms.Button
    $btnEndNow.Text = "Now"
    $btnEndNow.Location = New-Object System.Drawing.Point(240, 100)
    $btnEndNow.Size = New-Object System.Drawing.Size(50, 24)
    $btnEndNow.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnEndNow.FlatAppearance.BorderSize = 0
    $btnEndNow.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnEndNow.ForeColor = [System.Drawing.Color]::White
    $btnEndNow.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $btnEndNow.Cursor = [System.Windows.Forms.Cursors]::Hand
    $top.Controls.Add($btnEndNow)

    # Hours (calculated, read-only)
    $top.Controls.Add((& $mklabel "Hours:" 140))
    $lblHours = New-Object System.Windows.Forms.Label
    $lblHours.Location = New-Object System.Drawing.Point(144, 144)
    $lblHours.Size = New-Object System.Drawing.Size(180, 22)
    $lblHours.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblHours.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $lblHours.Text = "--"
    $top.Controls.Add($lblHours)

    # Work Order #
    $top.Controls.Add((& $mklabel "Work Order #:" 180))
    $txtWo = New-Object System.Windows.Forms.TextBox
    $txtWo.Location = New-Object System.Drawing.Point(144, 180)
    $txtWo.Size = New-Object System.Drawing.Size(180, 24)
    $txtWo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtWo.Text = "$($Job.workOrderNo)"
    $top.Controls.Add($txtWo)

    $lblWoHint = New-Object System.Windows.Forms.Label
    $lblWoHint.Text = "required if >30 min or parts used"
    $lblWoHint.Location = New-Object System.Drawing.Point(334, 184)
    $lblWoHint.Size = New-Object System.Drawing.Size(340, 22)
    $lblWoHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblWoHint.ForeColor = [System.Drawing.Color]::FromArgb(160,170,185)
    $top.Controls.Add($lblWoHint)

    # Description
    $top.Controls.Add((& $mklabel "Description:" 220))
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Location = New-Object System.Drawing.Point(144, 220)
    $txtDesc.Size = New-Object System.Drawing.Size(530, 88)
    $txtDesc.Multiline = $true
    $txtDesc.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtDesc.Text = "$($Job.description)"
    $top.Controls.Add($txtDesc)

    # ---------- parts panel ----------
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Fill'
    $mid.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $mid.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 8)
    $form.Controls.Add($mid)
    $mid.BringToFront()

    $lblPartsHead = New-Object System.Windows.Forms.Label
    $lblPartsHead.Text = "Parts Used"
    $lblPartsHead.Dock = 'Top'
    $lblPartsHead.Height = 22
    $lblPartsHead.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblPartsHead.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $mid.Controls.Add($lblPartsHead)

    $partsToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $partsToolbar.Dock = 'Bottom'
    $partsToolbar.Height = 34
    $partsToolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $partsToolbar.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
    $mid.Controls.Add($partsToolbar)

    $addPartBtn = New-Object System.Windows.Forms.Button
    $addPartBtn.Text = "+ Add Part"
    $addPartBtn.Size = New-Object System.Drawing.Size(110, 26)
    $addPartBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addPartBtn.FlatAppearance.BorderSize = 0
    $addPartBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $addPartBtn.ForeColor = [System.Drawing.Color]::White
    $addPartBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $addPartBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $addPartBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,6,0)
    $partsToolbar.Controls.Add($addPartBtn)

    $rmPartBtn = New-Object System.Windows.Forms.Button
    $rmPartBtn.Text = "Remove"
    $rmPartBtn.Size = New-Object System.Drawing.Size(90, 26)
    $rmPartBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $rmPartBtn.FlatAppearance.BorderSize = 0
    $rmPartBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $rmPartBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $rmPartBtn.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $rmPartBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $partsToolbar.Controls.Add($rmPartBtn)

    $lvParts = New-Object System.Windows.Forms.ListView
    $lvParts.Dock = 'Fill'
    $lvParts.Columns.Add("Part Number", 140) | Out-Null
    $lvParts.Columns.Add("Description", 260) | Out-Null
    $lvParts.Columns.Add("Qty", 50)          | Out-Null
    $lvParts.Columns.Add("OEM", 120)         | Out-Null
    $lvParts.Columns.Add("Source", 140)      | Out-Null
    Set-ListViewStyle -ListView $lvParts
    $mid.Controls.Add($lvParts)
    $lvParts.BringToFront()

    # ---------- validation label ----------
    $lblVal = New-Object System.Windows.Forms.Label
    $lblVal.Location = New-Object System.Drawing.Point(120, 16)
    $lblVal.Size = New-Object System.Drawing.Size(340, 24)
    $lblVal.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblVal.ForeColor = [System.Drawing.Color]::FromArgb(179,38,30)
    $btnBar.Controls.Add($lblVal)

    # ---------- internal state ----------
    # ---------- internal state ----------
    $script:__je_parts = @()
    foreach ($p in @($Job.parts)) { $script:__je_parts += $p }

    $renderParts = {
        $lvParts.Items.Clear()
        foreach ($p in $script:__je_parts) {
            $it = New-Object System.Windows.Forms.ListViewItem("$($p.PartNumber)")
            $it.SubItems.Add("$($p.Description)") | Out-Null
            $it.SubItems.Add("$($p.Quantity)")    | Out-Null
            $it.SubItems.Add("$($p.OEMNumber)")   | Out-Null
            $it.SubItems.Add("$($p.Source)")      | Out-Null
            $lvParts.Items.Add($it) | Out-Null
        }
    }

    $recalcHours = {
        $tmp = [PSCustomObject]@{
            startTime = $dtpStart.Value.ToString('HH:mm')
            endTime   = $dtpEnd.Value.ToString('HH:mm')
            hoursOverride = $null
        }
        $h = Get-JobHours $tmp
        if ($null -ne $h) { $lblHours.Text = "$($h.ToString('F2')) hrs" } else { $lblHours.Text = "--" }
    }

    $updateParsed = {
        $m = Format-MachineId $txtMachine.Text
        if ($null -eq $m) {
            $lblParsed.Text = ''
            $lblParsed.ForeColor = [System.Drawing.Color]::FromArgb(160,170,185)
        } elseif ($m.IsNamed) {
            $lblParsed.Text = "-> $($m.Canonical)"
            $lblParsed.ForeColor = [System.Drawing.Color]::FromArgb(52,152,219)
        } else {
            $lblParsed.Text = "-> $($m.Canonical)  (unnamed MNO)"
            $lblParsed.ForeColor = [System.Drawing.Color]::FromArgb(230,126,34)
        }
    }

    # prime datetime pickers from the job (fall back to "now")
    $nowHHmm = Get-Date -Format 'HH:mm'
    $startStr = if (-not [string]::IsNullOrWhiteSpace($Job.startTime)) { $Job.startTime } else { $nowHHmm }
    $endStr   = if (-not [string]::IsNullOrWhiteSpace($Job.endTime))   { $Job.endTime }   else { $nowHHmm }
    try {
        $dtpStart.Value = [DateTime]::ParseExact($startStr, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {}
    try {
        $dtpEnd.Value = [DateTime]::ParseExact($endStr, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {}

    & $renderParts
    & $recalcHours
    & $updateParsed

    # ---------- events ----------
    $btnStartNow.Add_Click({ $dtpStart.Value = Get-Date }.GetNewClosure())
    $btnEndNow.Add_Click({ $dtpEnd.Value = Get-Date }.GetNewClosure())
    $dtpStart.Add_ValueChanged({ & $recalcHours })
    $dtpEnd.Add_ValueChanged({ & $recalcHours })
    $txtMachine.Add_TextChanged({ & $updateParsed })

    $addPartBtn.Add_Click({
        $picked = Show-PartPicker -InitialParts @()
        if ($null -eq $picked) { return }
        foreach ($p in @($picked)) { $script:__je_parts += $p }
        & $renderParts
    })

    $rmPartBtn.Add_Click({
        if ($lvParts.SelectedIndices.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select a part to remove.", "Job Editor", "OK", "Information")
            return
        }
        $idx = $lvParts.SelectedIndices[0]
        $newList = @()
        for ($i = 0; $i -lt $script:__je_parts.Count; $i++) {
            if ($i -ne $idx) { $newList += $script:__je_parts[$i] }
        }
        $script:__je_parts = $newList
        & $renderParts
    })

    $cancelBtn.Add_Click({
        $form.Tag = $null
        $form.Close()
    })

    $deleteBtn.Add_Click({
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Delete this job? This cannot be undone.",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $form.Tag = [PSCustomObject]@{ Action = 'delete'; Job = $Job }
        $form.Close()
    })

    $saveBtn.Add_Click({
        $candidate = [PSCustomObject]@{
            id              = $Job.id
            kind            = $Job.kind
            machineId       = $txtMachine.Text.Trim()
            startTime       = $dtpStart.Value.ToString('HH:mm')
            endTime         = $dtpEnd.Value.ToString('HH:mm')
            hoursOverride   = $Job.hoursOverride
            description     = $txtDesc.Text.Trim()
            workOrderNo     = $txtWo.Text.Trim()
            parts           = $script:__je_parts
            createdAt       = $Job.createdAt
            sentToHistorian = $Job.sentToHistorian
            sentAt          = $Job.sentAt
        }

        $errs = Test-JobFields -Job $candidate
        if ($errs.Count -gt 0) {
            $lblVal.Text = "! " + $errs[0]
            return
        }
        $lblVal.Text = ''

        $m = Format-MachineId $candidate.machineId
        $candidate.machineId = $m.Canonical
        if ($m.IsNamed) { Register-MachineIfNew -Acronym $m.Acronym -Mno $m.Mno }

        $form.Tag = [PSCustomObject]@{ Action = 'save'; Job = $candidate }
        $form.Close()
    })

    $form.AcceptButton = $saveBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-PartPicker {
    # Modal dialog. Search parts, pick rows, enter quantities, return an
    # array of part objects.  Returns $null on cancel, empty array on OK
    # with nothing chosen.
    param(
        [array]$InitialParts = @()
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Add Parts"
    $form.Size = New-Object System.Drawing.Size(880, 700)
    $form.MinimumSize = New-Object System.Drawing.Size(880, 700)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ---------- bottom bar ----------
    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $btnBar.BackColor = [System.Drawing.Color]::Transparent
    $form.Controls.Add($btnBar)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(650, 10)
    $cancelBtn.Size = New-Object System.Drawing.Size(100, 32)
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBar.Controls.Add($cancelBtn)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = "OK"
    $okBtn.Location = New-Object System.Drawing.Point(760, 10)
    $okBtn.Size = New-Object System.Drawing.Size(100, 32)
    $okBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnBar.Controls.Add($okBtn)

    # ---------- search panel ----------
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $searchPanel.Size = New-Object System.Drawing.Size(850, 70)
    $searchPanel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($searchPanel)

    $lblNSN = New-Object System.Windows.Forms.Label
    $lblNSN.Text = "Part/NSN:"
    $lblNSN.Location = New-Object System.Drawing.Point(14, 14)
    $lblNSN.Size = New-Object System.Drawing.Size(80, 22)
    $lblNSN.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $searchPanel.Controls.Add($lblNSN)

    $txtNSN = New-Object System.Windows.Forms.TextBox
    $txtNSN.Location = New-Object System.Drawing.Point(96, 12)
    $txtNSN.Size = New-Object System.Drawing.Size(180, 24)
    $searchPanel.Controls.Add($txtNSN)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(290, 14)
    $lblDesc.Size = New-Object System.Drawing.Size(84, 22)
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $searchPanel.Controls.Add($lblDesc)

    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Location = New-Object System.Drawing.Point(376, 12)
    $txtDesc.Size = New-Object System.Drawing.Size(260, 24)
    $searchPanel.Controls.Add($txtDesc)

    $searchBtn = New-Object System.Windows.Forms.Button
    $searchBtn.Text = "Search"
    $searchBtn.Location = New-Object System.Drawing.Point(650, 10)
    $searchBtn.Size = New-Object System.Drawing.Size(110, 28)
    $searchBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $searchBtn.ForeColor = [System.Drawing.Color]::White
    $searchBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $searchBtn.FlatAppearance.BorderSize = 0
    $searchBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $searchBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $searchPanel.Controls.Add($searchBtn)

    # ---------- results ----------
    $lblResults = New-Object System.Windows.Forms.Label
    $lblResults.Text = "Search Results"
    $lblResults.Location = New-Object System.Drawing.Point(14, 88)
    $lblResults.Size = New-Object System.Drawing.Size(200, 20)
    $lblResults.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblResults.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblResults)

    $lvResults = New-Object System.Windows.Forms.ListView
    $lvResults.Location = New-Object System.Drawing.Point(14, 110)
    $lvResults.Size = New-Object System.Drawing.Size(846, 220)
    $lvResults.CheckBoxes = $true
    $lvResults.Columns.Add("Part Number", 130) | Out-Null
    $lvResults.Columns.Add("Description", 260) | Out-Null
    $lvResults.Columns.Add("Qty", 55)          | Out-Null
    $lvResults.Columns.Add("Location", 110)    | Out-Null
    $lvResults.Columns.Add("OEM", 120)         | Out-Null
    $lvResults.Columns.Add("Source", 150)      | Out-Null
    Set-ListViewStyle -ListView $lvResults
    $form.Controls.Add($lvResults)

    # ---------- move buttons ----------
    $addSelBtn = New-Object System.Windows.Forms.Button
    $addSelBtn.Text = "Add Selected  v"
    $addSelBtn.Location = New-Object System.Drawing.Point(14, 338)
    $addSelBtn.Size = New-Object System.Drawing.Size(150, 30)
    $addSelBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $addSelBtn.ForeColor = [System.Drawing.Color]::White
    $addSelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addSelBtn.FlatAppearance.BorderSize = 0
    $addSelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $addSelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($addSelBtn)

    # ---------- chosen list ----------
    $lblChosen = New-Object System.Windows.Forms.Label
    $lblChosen.Text = "Selected Parts"
    $lblChosen.Location = New-Object System.Drawing.Point(14, 378)
    $lblChosen.Size = New-Object System.Drawing.Size(200, 20)
    $lblChosen.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblChosen.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $form.Controls.Add($lblChosen)

    $lvChosen = New-Object System.Windows.Forms.ListView
    $lvChosen.Location = New-Object System.Drawing.Point(14, 400)
    $lvChosen.Size = New-Object System.Drawing.Size(846, 200)
    $lvChosen.Columns.Add("Part Number", 130) | Out-Null
    $lvChosen.Columns.Add("Description", 260) | Out-Null
    $lvChosen.Columns.Add("Qty", 55)          | Out-Null
    $lvChosen.Columns.Add("Location", 110)    | Out-Null
    $lvChosen.Columns.Add("OEM", 120)         | Out-Null
    $lvChosen.Columns.Add("Source", 150)      | Out-Null
    Set-ListViewStyle -ListView $lvChosen
    $form.Controls.Add($lvChosen)

    $rmChosenBtn = New-Object System.Windows.Forms.Button
    $rmChosenBtn.Text = "Remove"
    $rmChosenBtn.Location = New-Object System.Drawing.Point(170, 338)
    $rmChosenBtn.Size = New-Object System.Drawing.Size(100, 30)
    $rmChosenBtn.BackColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $rmChosenBtn.ForeColor = [System.Drawing.Color]::White
    $rmChosenBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $rmChosenBtn.FlatAppearance.BorderSize = 0
    $rmChosenBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $rmChosenBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($rmChosenBtn)

    # ---------- state ----------
    $script:__pp_chosen = @()
    foreach ($p in @($InitialParts)) { $script:__pp_chosen += $p }

    $renderChosen = {
        $lvChosen.Items.Clear()
        foreach ($p in $script:__pp_chosen) {
            $it = New-Object System.Windows.Forms.ListViewItem("$($p.PartNumber)")
            $it.SubItems.Add("$($p.Description)") | Out-Null
            $it.SubItems.Add("$($p.Quantity)")    | Out-Null
            $it.SubItems.Add("$($p.Location)")    | Out-Null
            $it.SubItems.Add("$($p.OEMNumber)")   | Out-Null
            $it.SubItems.Add("$($p.Source)")      | Out-Null
            $lvChosen.Items.Add($it) | Out-Null
        }
    }

    & $renderChosen

    # ---------- events ----------
    $searchBtn.Add_Click({
        $lvResults.Items.Clear()
        $nsn = $txtNSN.Text.Trim()
        $dsc = $txtDesc.Text.Trim()
        try {
            $found = Search-Parts -NSN $nsn -Description $dsc
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Search failed: $($_.Exception.Message)", "Error", "OK", "Error")
            return
        }
        foreach ($part in @($found)) {
            $it = New-Object System.Windows.Forms.ListViewItem("$($part.PartNumber)")
            $it.SubItems.Add("$($part.Description)") | Out-Null
            $it.SubItems.Add("$($part.Quantity)")    | Out-Null
            $it.SubItems.Add("$($part.Location)")    | Out-Null
            $it.SubItems.Add("$($part.OEMNumber)")   | Out-Null
            $it.SubItems.Add("$($part.Source)")      | Out-Null
            $it.Tag = $part
            $lvResults.Items.Add($it) | Out-Null
        }
    })

    $addSelBtn.Add_Click({
        foreach ($it in $lvResults.CheckedItems) {
            $p = $it.Tag
            if ($null -eq $p) { continue }
            $qty = Get-PartQuantity -PartNumber $p.PartNumber -MaxQty ([int]$p.Quantity)
            if ($qty -le 0) { continue }

            $chosenPart = [PSCustomObject]@{
                PartNumber = $p.PartNumber
                Description = $p.Description
                Quantity   = $qty
                OEMNumber  = $p.OEMNumber
                Source     = $p.Source
                Location   = $p.Location
            }
            $script:__pp_chosen += $chosenPart
        }
        & $renderChosen
    })

    $rmChosenBtn.Add_Click({
        if ($lvChosen.SelectedIndices.Count -eq 0) { return }
        $idx = $lvChosen.SelectedIndices[0]
        $newList = @()
        for ($i = 0; $i -lt $script:__pp_chosen.Count; $i++) {
            if ($i -ne $idx) { $newList += $script:__pp_chosen[$i] }
        }
        $script:__pp_chosen = $newList
        & $renderChosen
    })

    $cancelBtn.Add_Click({
        $form.Tag = $null
        $form.Close()
    })

    $okBtn.Add_Click({
        $form.Tag = $script:__pp_chosen
        $form.Close()
    })

    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

# ============================================================
# Work Tracking mega-tab
# ============================================================

function New-SubTab {
    param(
        [string]$Title,
        [string]$PlaceholderText = "Coming soon."
    )
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $Title
    $tab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $tab.Padding = New-Object System.Windows.Forms.Padding(12)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $PlaceholderText
    $lbl.Dock = 'Fill'
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Italic)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $tab.Controls.Add($lbl)

    return $tab
}

# ============================================================
# Work Tracking — Weekly Worksheet (eDAC)
# ============================================================

function Format-eDacDate {
    param([DateTime]$Date)
    $months = @('JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC')
    return "$($Date.Day.ToString('00'))-$($months[$Date.Month - 1])-$($Date.ToString('yy'))"
}

function Get-WeeklyWorksheetDir {
    $rel = $script:config.WorkTracking.WeeklyWorksheetsDirectory
    if ([string]::IsNullOrWhiteSpace($rel)) { throw "WorkTracking.WeeklyWorksheetsDirectory not set." }
    if ([System.IO.Path]::IsPathRooted($rel)) { return $rel }
    $dir = Join-Path $script:config.RootDirectory $rel
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function Get-WeeklyWorksheetPath {
    param([DateTime]$Date)
    return (Join-Path (Get-WeeklyWorksheetDir) ($Date.ToString("yyyy-MM-dd") + ".json"))
}

function Get-WeeklyWorksheet {
    param([DateTime]$Date)

    $path = Get-WeeklyWorksheetPath -Date $Date
    if (-not (Test-Path $path)) { return $null }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        Write-Log "Get-WeeklyWorksheet: read error: $($_.Exception.Message)"
        return $null
    }
}

function Save-WeeklyWorksheet {
    param([DateTime]$Date, $Worksheet)

    $path = Get-WeeklyWorksheetPath -Date $Date
    $tmp  = "$path.tmp"
    try {
        $Worksheet | ConvertTo-Json -Depth 12 | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $path -Force
        return $true
    } catch {
        Write-Log "Save-WeeklyWorksheet: $($_.Exception.Message)"
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        return $false
    }
}

function Test-PmMachineExistsFor {
    param($Acronym, $Mno, [DateTime]$Date)
    if ([string]::IsNullOrWhiteSpace($Acronym)) { return $false }
    $mnoInt = 0
    if (-not [int]::TryParse("$Mno", [ref]$mnoInt)) { return $false }
    foreach ($m in @($script:pmMachines)) {
        if (-not $m.SortDate -or $m.SortDate.Date -ne $Date.Date) { continue }
        if ("$($m.Acronym)" -ne "$Acronym") { continue }
        $pmMno = 0
        if (-not [int]::TryParse("$($m.Mno)", [ref]$pmMno)) { continue }
        if ($pmMno -eq $mnoInt) { return $true }
    }
    return $false
}

function Get-PmSelectedHoursFor {
    param($Acronym, $Mno, [DateTime]$Date)
    $totalMin = 0.0
    $mnoInt = 0
    if (-not [int]::TryParse("$Mno", [ref]$mnoInt)) { return 0.0 }
    foreach ($m in @($script:pmMachines)) {
        if (-not $m.SortDate -or $m.SortDate.Date -ne $Date.Date) { continue }
        if ("$($m.Acronym)" -ne "$Acronym") { continue }
        $pmMno = 0
        if (-not [int]::TryParse("$($m.Mno)", [ref]$pmMno)) { continue }
        if ($pmMno -ne $mnoInt) { continue }
        foreach ($t in @($m.Parsed.Tasks)) {
            $e = $m.State[$t.ItemNo]
            if (-not $e -or -not $e.Selected) { continue }
            $min = if ($null -ne $e.CustomTimeMin) { [int]$e.CustomTimeMin } else { [int]$t.EstTimeMin }
            $totalMin += $min
        }
    }
    return [Math]::Round($totalMin / 60.0, 2)
}

function Sync-WorksheetActualTimes {
    # D1 rule:
    #   row.ChecklistNo == '000' and PM data exists -> sum selected PM hours
    #   no PM data for that machine -> copy estimated hours
    #   otherwise -> leave alone (manual entry stays)
    # Manual overrides are respected until user clicks Sync (which we call here).
    param($Worksheet, [DateTime]$Date)

    foreach ($row in @($Worksheet.Rows)) {
        $acr = "$($row.Acronym)"
        $mno = "$($row.Equipment)"
        $checkNo = "$($row.ChecklistNo)"

        $hasPm = Test-PmMachineExistsFor -Acronym $acr -Mno $mno -Date $Date

        if ($hasPm -and $checkNo -eq '000') {
            $hours = Get-PmSelectedHoursFor -Acronym $acr -Mno $mno -Date $Date
            $row.ActualTime    = $hours
            $row.AutoValue     = $hours
            $row.ManualOverride = $false
        }
        elseif (-not $hasPm -and "$($row.RowType)" -eq 'B') {
            $est = 0.0
            [void][double]::TryParse("$($row.EstimatedHrs)", [ref]$est)
            $row.ActualTime    = [Math]::Round($est, 2)
            $row.AutoValue     = [Math]::Round($est, 2)
            $row.ManualOverride = $false
        }
        elseif ($null -ne $row.AutoValue -and -not $row.ManualOverride) {
            # Auto source disappeared; clear unless user typed over
            $row.AutoValue     = $null
            $row.ActualTime    = $null
            $row.ManualOverride = $false
        }
    }
}

function Add-WorksheetPlaceholders {
    # For every PM machine on the given date with no matching worksheet row,
    # append a NeedsAudit placeholder row.
    param($Worksheet, [DateTime]$Date)

    $existing = @($Worksheet.Rows)
    foreach ($m in @($script:pmMachines)) {
        if (-not $m.SortDate -or $m.SortDate.Date -ne $Date.Date) { continue }
        $pmMno = 0
        if (-not [int]::TryParse("$($m.Mno)", [ref]$pmMno)) { continue }

        $found = $false
        foreach ($r in $existing) {
            if ("$($r.Acronym)" -ne "$($m.Acronym)") { continue }
            $rmno = 0
            if (-not [int]::TryParse("$($r.Equipment)", [ref]$rmno)) { continue }
            if ($rmno -eq $pmMno) { $found = $true; break }
        }
        if ($found) { continue }

        $existing += [PSCustomObject]@{
            RowType        = 'PM'
            WorkOrderNo    = ''
            WorkCode       = ''
            Acronym        = $m.Acronym
            ClassCode      = ''
            Equipment      = $m.Mno
            Route          = ''
            Frequency      = ''
            Priority       = ''
            DueDate        = $m.ChecklistDate
            Description    = "PM Checklist (no worksheet row)"
            ChecklistNo    = $m.ChecklistNo
            EstimatedHrs   = 0.0
            PmDescription  = ''
            ActualTime     = $null
            AutoValue      = $null
            ManualOverride = $false
            NeedsAudit     = $true
        }
    }
    $Worksheet.Rows = $existing
}

function New-WorksheetFromHtml {
    # Parses the full eDAC HTML, extracts EVERY Scheduled Date section,
    # saves one JSON per date, and returns a summary.
    param([string]$Html, [string]$Source)

    if ([string]::IsNullOrWhiteSpace($Html)) { throw "No HTML provided." }

    # Detect all Scheduled Date values in the file.
    #
    # The attribute-value pair on the <TD> is deliberately not matched:
    # eMARS renders some pages with align=left (unquoted) and others with
    # align="left" (quoted), and pinning the pattern to either form breaks
    # the other.  [^>]*> swallows whatever attributes are present.
    $datePattern = 'Scheduled Date:\s*</TH>\s*<TD[^>]*>\s*(\d{2}-[A-Z]{3}-\d{2})'
    $foundDates = @()
    $matches = [regex]::Matches($Html, $datePattern, 'IgnoreCase')
    Write-Log "New-WorksheetFromHtml: date pattern found $($matches.Count) match(es)"

    foreach ($m in $matches) {
        $v = $m.Groups[1].Value.Trim()
        if ($v -and $foundDates -notcontains $v) { $foundDates += $v }
    }

    if ($foundDates.Count -eq 0) {
        # Diagnostic: show what header text IS present so future breakage
        # is diagnosable from the log alone.
        $headerProbe = [regex]::Matches($Html, 'Scheduled Date[^<]{0,40}')
        Write-Log "New-WorksheetFromHtml: no dates matched. Header probe ($($headerProbe.Count) hit(s)):"
        foreach ($hp in ($headerProbe | Select-Object -First 3)) {
            Write-Log "  -> '$($hp.Value)'"
        }
        throw "No 'Scheduled Date' sections found in the HTML."
    }

    Write-Log "New-WorksheetFromHtml: found dates: $($foundDates -join ', ')"

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ("edac_" + [guid]::NewGuid().ToString() + ".html")
    $saved = @()
    try {
        [System.IO.File]::WriteAllText($temp, $Html, [System.Text.Encoding]::UTF8)

        foreach ($target in $foundDates) {
            # Convert 20-SEP-26 -> 2026-09-20
            $dt = $null
            try { $dt = [DateTime]::ParseExact($target, 'dd-MMM-yy', [System.Globalization.CultureInfo]::InvariantCulture) }
            catch {
                Write-Log "New-WorksheetFromHtml: could not parse date '$target' — skipping"
                continue
            }

            $rows = @(ConvertFrom-eDacWorksheet -HtmlPath $temp -TargetDate $target)
            Write-Log "New-WorksheetFromHtml: date $target -> $($rows.Count) data row(s)"

            $decorated = @()
            foreach ($r in $rows) {
                $decorated += [PSCustomObject]@{
                    RowType        = $r.RowType
                    WorkOrderNo    = $r.WorkOrderNo
                    WorkCode       = $r.WorkCode
                    Acronym        = $r.Acronym
                    ClassCode      = $r.ClassCode
                    Equipment      = $r.Equipment
                    Route          = $r.Route
                    Frequency      = $r.Frequency
                    Priority       = $r.Priority
                    DueDate        = $r.DueDate
                    Description    = $r.Description
                    ChecklistNo    = $r.ChecklistNo
                    EstimatedHrs   = $r.EstimatedHrs
                    PmDescription  = $r.PmDescription
                    ActualTime     = $null
                    AutoValue      = $null
                    ManualOverride = $false
                    NeedsAudit     = $false
                }
            }

            $ws = [PSCustomObject]@{
                Date      = $dt.ToString("yyyy-MM-dd")
                FetchedAt = (Get-Date).ToString("o")
                Source    = $Source
                Rows      = $decorated
            }

            Sync-WorksheetActualTimes -Worksheet $ws -Date $dt
            Add-WorksheetPlaceholders -Worksheet $ws -Date $dt

            Save-WeeklyWorksheet -Date $dt -Worksheet $ws | Out-Null
            $saved += [PSCustomObject]@{ Date = $dt.ToString("yyyy-MM-dd"); Rows = $decorated.Count }
        }
    } finally {
        if (Test-Path $temp) { Remove-Item $temp -Force -ErrorAction SilentlyContinue }
    }

    return [PSCustomObject]@{
        Saved        = $saved
        DatesParsed  = $foundDates.Count
        TotalRows    = ($saved | Measure-Object -Property Rows -Sum).Sum
    }
}

function Show-WorksheetRowEditor {
    param($Row)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Worksheet row — $($Row.WorkOrderNo)"
    $form.Size = New-Object System.Drawing.Size(560, 300)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $meta = New-Object System.Windows.Forms.Label
    $meta.Location = New-Object System.Drawing.Point(12, 12)
    $meta.Size = New-Object System.Drawing.Size(520, 90)
    $meta.Font = New-Object System.Drawing.Font("Consolas", 9)
    $meta.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $meta.Text = @(
        "Machine:   $($Row.Acronym) $($Row.Equipment)"
        "W/O #:     $($Row.WorkOrderNo)"
        "Due:       $($Row.DueDate)"
        "Est. hrs:  $($Row.EstimatedHrs)"
        "Description:"
        "  $($Row.Description)"
    ) -join "`r`n"
    $form.Controls.Add($meta)

    $lblAct = New-Object System.Windows.Forms.Label
    $lblAct.Text = "Actual Hours:"
    $lblAct.Location = New-Object System.Drawing.Point(12, 118)
    $lblAct.Size = New-Object System.Drawing.Size(100, 22)
    $lblAct.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblAct.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblAct)

    $numAct = New-Object System.Windows.Forms.NumericUpDown
    $numAct.Location = New-Object System.Drawing.Point(120, 118)
    $numAct.Size = New-Object System.Drawing.Size(90, 24)
    $numAct.Minimum = 0
    $numAct.Maximum = 24
    $numAct.DecimalPlaces = 2
    $numAct.Increment = 0.25
    if ($null -ne $Row.ActualTime) { $numAct.Value = [decimal]$Row.ActualTime }
    $form.Controls.Add($numAct)

    $lblAuto = New-Object System.Windows.Forms.Label
    $lblAuto.Location = New-Object System.Drawing.Point(220, 122)
    $lblAuto.Size = New-Object System.Drawing.Size(320, 22)
    $lblAuto.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblAuto.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    if ($null -ne $Row.AutoValue) {
        $lblAuto.Text = "auto value: $($Row.AutoValue) h"
    } else {
        $lblAuto.Text = "no auto value"
    }
    $form.Controls.Add($lblAuto)

    if ($Row.NeedsAudit) {
        $lblAudit = New-Object System.Windows.Forms.Label
        $lblAudit.Location = New-Object System.Drawing.Point(12, 150)
        $lblAudit.Size = New-Object System.Drawing.Size(520, 22)
        $lblAudit.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lblAudit.ForeColor = [System.Drawing.Color]::FromArgb(200,120,0)
        $lblAudit.Text = "! Needs Audit — this row was created from a PM checklist with no worksheet counterpart."
        $form.Controls.Add($lblAudit)
    }

    $clearAutoBtn = New-Object System.Windows.Forms.Button
    $clearAutoBtn.Text = "Clear Auto"
    $clearAutoBtn.Location = New-Object System.Drawing.Point(120, 180)
    $clearAutoBtn.Size = New-Object System.Drawing.Size(100, 30)
    $clearAutoBtn.FlatStyle = 'Flat'
    $clearAutoBtn.FlatAppearance.BorderSize = 0
    $clearAutoBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $clearAutoBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $clearAutoBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $clearAutoBtn.Cursor = 'Hand'
    $clearAutoBtn.Add_Click({
        $form.Tag = [PSCustomObject]@{ Action = 'clear' }
        $form.Close()
    })
    $form.Controls.Add($clearAutoBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(340, 230)
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 30)
    $cancelBtn.FlatStyle = 'Flat'
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = 'Hand'
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $form.Controls.Add($cancelBtn)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Save"
    $saveBtn.Location = New-Object System.Drawing.Point(440, 230)
    $saveBtn.Size = New-Object System.Drawing.Size(90, 30)
    $saveBtn.FlatStyle = 'Flat'
    $saveBtn.FlatAppearance.BorderSize = 0
    $saveBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $saveBtn.ForeColor = [System.Drawing.Color]::White
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.Cursor = 'Hand'
    $saveBtn.Add_Click({
        $form.Tag = [PSCustomObject]@{ Action = 'save'; Value = [double]$numAct.Value }
        $form.Close()
    })
    $form.Controls.Add($saveBtn)

    $form.AcceptButton = $saveBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Setup-WorkTrackingTab {
    param($parentTab)

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(6)

    $subTabs = New-Object System.Windows.Forms.TabControl
    $subTabs.Dock = 'Fill'
    $subTabs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subTabs.Padding = New-Object System.Drawing.Point(14, 4)
    $parentTab.Controls.Add($subTabs)

    $jobLogTab = New-Object System.Windows.Forms.TabPage
    $jobLogTab.Text = "Job Log"
    $subTabs.TabPages.Add($jobLogTab)
    Setup-JobLogSubTab -parentTab $jobLogTab
    $pmChecklistTab = New-Object System.Windows.Forms.TabPage
    $pmChecklistTab.Text = "PM Checklist"
    $subTabs.TabPages.Add($pmChecklistTab)
    Setup-PmChecklistSubTab -parentTab $pmChecklistTab

    $worksheetTab = New-Object System.Windows.Forms.TabPage
    $worksheetTab.Text = "Weekly Worksheet"
    $subTabs.TabPages.Add($worksheetTab)
    Setup-WeeklyWorksheetSubTab -parentTab $worksheetTab    
    $historianTab = New-Object System.Windows.Forms.TabPage
    $historianTab.Text = "Historian"
    $subTabs.TabPages.Add($historianTab)
    Setup-HistorianSubTab -parentTab $historianTab
}

# ============================================================
# Settings tab
# ============================================================

function New-SettingRow {
    param(
        [System.Windows.Forms.Control]$Parent,
        [int]$Top,
        [string]$LabelText,
        [System.Windows.Forms.Control]$InputControl
    )

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Location = New-Object System.Drawing.Point(12, ($Top + 4))
    $lbl.Size = New-Object System.Drawing.Size(240, 22)
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $Parent.Controls.Add($lbl)

    $InputControl.Location = New-Object System.Drawing.Point(260, $Top)
    $Parent.Controls.Add($InputControl)
}

# ============================================================
# Settings tab — computed Work Target helpers
# ============================================================

function Get-ComputedWorkTargetHours {
    # Target = day length − (all non-work components, minutes → hours)
    param($Settings)
    if (-not $Settings) { return 0.0 }

    $dayHours = [double]$Settings.DayLenBox.Value

    $nonWorkMinutes = 0.0
    foreach ($box in @(
        $Settings.LunchBox,
        $Settings.WashupBox,
        $Settings.StartupBox,
        $Settings.Break1Box,
        $Settings.Break2Box,
        $Settings.EndWashBox,
        $Settings.PaperBox
    )) {
        $nonWorkMinutes += [double]$box.Value
    }

    $target = $dayHours - ($nonWorkMinutes / 60.0)
    if ($target -lt 0)  { $target = 0 }
    if ($target -gt 24) { $target = 24 }
    return [Math]::Round($target, 2)
}

function Update-WorkTargetDisplay {
    if (-not $script:wtSettings) { return }
    if (-not $script:wtSettings.TargetBox) { return }
    $t = Get-ComputedWorkTargetHours -Settings $script:wtSettings
    $script:wtSettings.TargetBox.Value = [decimal]$t
}

function Load-WtSettingsValues {
    if (-not $script:wtSettings) { return }
    if (-not $script:config)     { return }

    $c  = $script:wtSettings
    $wt = $script:config.WorkTracking

    Write-Log "Load-WtSettingsValues: configPath='$script:configPath'"

    if (-not $wt) {
        Write-Log "Load-WtSettingsValues: WorkTracking is null — skipping."
        return
    }

    # --- Direct inputs ---
    $c.TechNameBox.Text   = if ($wt.TechnicianName)       { "$($wt.TechnicianName)" }       else { '' }
    $c.EmailBox.Text      = if ($script:config.SupervisorEmail) { "$($script:config.SupervisorEmail)" } else { '' }
    $c.EdacUrlBox.Text    = if ($wt.eDacUrl)              { "$($wt.eDacUrl)" }              else { '' }
    $c.ThresholdBox.Value = if ($wt.ReactiveToWorkOrderMinutes) { [int]$wt.ReactiveToWorkOrderMinutes } else { 15 }

    Write-Log "Load-WtSettingsValues: techName='$($c.TechNameBox.Text)' email='$($c.EmailBox.Text)'"

    # --- Work budget inputs ---
    $wb = $wt.WorkBudget
    if ($wb) {
        $c.DayLenBox.Value  = if ($wb.DayLengthHours) { [decimal]$wb.DayLengthHours } else { [decimal]8.5 }
        $c.LunchBox.Value   = if ($wb.LunchMinutes)   { [decimal]$wb.LunchMinutes }   else { [decimal]30 }
        $c.WashupBox.Value  = if ($wb.WashupMinutes)  { [decimal]$wb.WashupMinutes }  else { [decimal]15 }
        $c.StartupBox.Value = if ($wb.StartupMinutes) { [decimal]$wb.StartupMinutes } else { [decimal]15 }
        $pbm = @($wb.PaidBreaksMinutes)
        $c.Break1Box.Value  = if ($pbm.Count -ge 1 -and $pbm[0]) { [decimal]$pbm[0] } else { [decimal]15 }
        $c.Break2Box.Value  = if ($pbm.Count -ge 2 -and $pbm[1]) { [decimal]$pbm[1] } else { [decimal]15 }
        $c.EndWashBox.Value = if ($wb.EndOfDayWashupMinutes) { [decimal]$wb.EndOfDayWashupMinutes } else { [decimal]15 }
        $c.PaperBox.Value   = if ($wb.PaperworkMinutes)      { [decimal]$wb.PaperworkMinutes }      else { [decimal]15 }
    }

    # Work Target is derived — always recompute from loaded inputs.
    Update-WorkTargetDisplay
}

function Setup-SettingsTab {
    param($parentTab)

    Write-Log "Setting up Settings tab..."

    $parentTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $parentTab.Padding = New-Object System.Windows.Forms.Padding(0)

    $root = New-Object System.Windows.Forms.Panel
    $root.Dock = 'Fill'
    $root.AutoScroll = $true
    $root.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $root.Padding = New-Object System.Windows.Forms.Padding(16)
    $parentTab.Controls.Add($root)

    # --- Bottom save bar ---
    $saveBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $saveBar.Dock = 'Bottom'
    $saveBar.Height = 52
    $saveBar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $saveBar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $saveBar.BackColor = [System.Drawing.Color]::Transparent
    $root.Controls.Add($saveBar)

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Save Settings"
    $saveBtn.Size = New-Object System.Drawing.Size(160, 36)
    $saveBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $saveBtn.ForeColor = [System.Drawing.Color]::White
    $saveBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveBtn.FlatAppearance.BorderSize = 0
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $saveBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
    $saveBar.Controls.Add($saveBtn)

    $reloadBtn = New-Object System.Windows.Forms.Button
    $reloadBtn.Text = "Reload from File"
    $reloadBtn.Size = New-Object System.Drawing.Size(160, 36)
    $reloadBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $reloadBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $reloadBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $reloadBtn.FlatAppearance.BorderSize = 0
    $reloadBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $reloadBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $reloadBtn.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
    $saveBar.Controls.Add($reloadBtn)

    # --- Body ---
    $body = New-Object System.Windows.Forms.FlowLayoutPanel
    $body.Dock = 'Fill'
    $body.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $body.WrapContents = $false
    $body.AutoScroll = $true
    $body.BackColor = [System.Drawing.Color]::Transparent
    $body.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 60)
    $root.Controls.Add($body)
    $body.BringToFront()

    # --- Identity group ---
    $gIdentity = New-Object System.Windows.Forms.GroupBox
    $gIdentity.Text = "Identity"
    $gIdentity.Width = 720
    $gIdentity.Height = 110
    $gIdentity.Padding = New-Object System.Windows.Forms.Padding(12)
    $gIdentity.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gIdentity.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gIdentity.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
    $body.Controls.Add($gIdentity)

    $techNameBox = New-Object System.Windows.Forms.TextBox
    $techNameBox.Size = New-Object System.Drawing.Size(440, 24)
    New-SettingRow -Parent $gIdentity -Top 28 -LabelText "Technician Name:" -InputControl $techNameBox

    $emailBox = New-Object System.Windows.Forms.TextBox
    $emailBox.Size = New-Object System.Drawing.Size(440, 24)
    New-SettingRow -Parent $gIdentity -Top 62 -LabelText "Supervisor Email:" -InputControl $emailBox

    # --- Work Tracking group ---
    $gWorkTrack = New-Object System.Windows.Forms.GroupBox
    $gWorkTrack.Text = "Work Tracking"
    $gWorkTrack.Width = 720
    $gWorkTrack.Height = 110
    $gWorkTrack.Padding = New-Object System.Windows.Forms.Padding(12)
    $gWorkTrack.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gWorkTrack.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gWorkTrack.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
    $body.Controls.Add($gWorkTrack)

    $edacUrlBox = New-Object System.Windows.Forms.TextBox
    $edacUrlBox.Size = New-Object System.Drawing.Size(440, 24)
    New-SettingRow -Parent $gWorkTrack -Top 28 -LabelText "eDAC Worksheet URL:" -InputControl $edacUrlBox

    $thresholdBox = New-Object System.Windows.Forms.NumericUpDown
    $thresholdBox.Size = New-Object System.Drawing.Size(100, 24)
    $thresholdBox.Minimum = 1
    $thresholdBox.Maximum = 600
    New-SettingRow -Parent $gWorkTrack -Top 62 -LabelText "Reactive -> WO threshold (min):" -InputControl $thresholdBox

    # --- Work Budget group ---
    $gBudget = New-Object System.Windows.Forms.GroupBox
    $gBudget.Text = "Work Budget"
    $gBudget.Width = 720
    $gBudget.Height = 40 + (9 * 34)
    $gBudget.Padding = New-Object System.Windows.Forms.Padding(12)
    $gBudget.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gBudget.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gBudget.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
    $body.Controls.Add($gBudget)

    $newNum = {
        param([decimal]$Min, [decimal]$Max, [decimal]$Inc = 1, [int]$Decimals = 0)
        $n = New-Object System.Windows.Forms.NumericUpDown
        $n.Size = New-Object System.Drawing.Size(100, 24)
        $n.Minimum = $Min
        $n.Maximum = $Max
        $n.Increment = $Inc
        $n.DecimalPlaces = $Decimals
        return $n
    }

    $dayLenBox  = & $newNum 1 24 0.25 2
    $lunchBox   = & $newNum 0 120
    $washupBox  = & $newNum 0 60
    $startupBox = & $newNum 0 60
    $break1Box  = & $newNum 0 60
    $break2Box  = & $newNum 0 60
    $endWashBox = & $newNum 0 60
    $paperBox   = & $newNum 0 60
    $targetBox  = & $newNum 0 24 0.25 2

    # Work Target is derived, not user-editable.
    $targetBox.Enabled   = $false
    $targetBox.BackColor = [System.Drawing.Color]::FromArgb(236,240,241)

    New-SettingRow -Parent $gBudget -Top  28 -LabelText "Day Length (hours):"                  -InputControl $dayLenBox
    New-SettingRow -Parent $gBudget -Top  62 -LabelText "Lunch (min, 0=skip):"                 -InputControl $lunchBox
    New-SettingRow -Parent $gBudget -Top  96 -LabelText "Wash-up at Lunch (min):"              -InputControl $washupBox
    New-SettingRow -Parent $gBudget -Top 130 -LabelText "Startup (min):"                       -InputControl $startupBox
    New-SettingRow -Parent $gBudget -Top 164 -LabelText "Paid Break 1 (min):"                  -InputControl $break1Box
    New-SettingRow -Parent $gBudget -Top 198 -LabelText "Paid Break 2 (min):"                  -InputControl $break2Box
    New-SettingRow -Parent $gBudget -Top 232 -LabelText "Wash-up End-of-Day (min):"            -InputControl $endWashBox
    New-SettingRow -Parent $gBudget -Top 266 -LabelText "Paperwork (min):"                     -InputControl $paperBox
    New-SettingRow -Parent $gBudget -Top 300 -LabelText "Work Target (hours, auto-computed):"  -InputControl $targetBox

    # --- Installation group ---
    $gInstall = New-Object System.Windows.Forms.GroupBox
    $gInstall.Text = "Installation"
    $gInstall.Width = 720
    $gInstall.Height = 168
    $gInstall.Padding = New-Object System.Windows.Forms.Padding(12)
    $gInstall.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $gInstall.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $gInstall.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
    $body.Controls.Add($gInstall)

    $currentRootLabel = New-Object System.Windows.Forms.Label
    $currentRootLabel.Location = New-Object System.Drawing.Point(12, 28)
    $currentRootLabel.Size = New-Object System.Drawing.Size(690, 22)
    $currentRootLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
    $currentRootLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,100,115)
    $currentRootLabel.Text = "Current root: $($script:config.RootDirectory)"
    $gInstall.Controls.Add($currentRootLabel)

    $installColX  = @(12, 240, 468)
    $installRowY  = @(58, 98)
    $installBtnW  = 220
    $installBtnH  = 32

    $btnSetupWizard = New-Object System.Windows.Forms.Button
    $btnSetupWizard.Text = "Run Installation Wizard"
    $btnSetupWizard.Location = New-Object System.Drawing.Point($installColX[0], $installRowY[0])
    $btnSetupWizard.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnSetupWizard.FlatStyle = 'Flat'
    $btnSetupWizard.FlatAppearance.BorderSize = 0
    $btnSetupWizard.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnSetupWizard.ForeColor = [System.Drawing.Color]::White
    $btnSetupWizard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSetupWizard.Cursor = 'Hand'
    $btnSetupWizard.Add_Click({
        $initial = Join-Path $PSScriptRoot 'PartsMgmt'
        $ok = Show-InstallationWizard -InitialRoot $initial
        if ($ok) {
            try {
                $script:config = Get-Content -Path $script:configPath -Raw | ConvertFrom-Json
                $currentRootLabel.Text = "Current root: $($script:config.RootDirectory)"
                Load-WtSettingsValues
            } catch {
                Write-Log "Post-wizard config reload failed: $($_.Exception.Message)"
            }
        }
    })
    $gInstall.Controls.Add($btnSetupWizard)

    $btnOpenRoot = New-Object System.Windows.Forms.Button
    $btnOpenRoot.Text = "Open Root Folder"
    $btnOpenRoot.Location = New-Object System.Drawing.Point($installColX[1], $installRowY[0])
    $btnOpenRoot.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnOpenRoot.FlatStyle = 'Flat'
    $btnOpenRoot.FlatAppearance.BorderSize = 0
    $btnOpenRoot.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnOpenRoot.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnOpenRoot.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOpenRoot.Cursor = 'Hand'
    $btnOpenRoot.Add_Click({
        $r = "$($script:config.RootDirectory)"
        if (Test-Path $r) { Start-Process $r }
        else { [System.Windows.Forms.MessageBox]::Show("Root folder not found: $r", "Settings", "OK", "Warning") }
    })
    $gInstall.Controls.Add($btnOpenRoot)

    $btnOpenScripts = New-Object System.Windows.Forms.Button
    $btnOpenScripts.Text = "Open Scripts Folder"
    $btnOpenScripts.Location = New-Object System.Drawing.Point($installColX[2], $installRowY[0])
    $btnOpenScripts.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnOpenScripts.FlatStyle = 'Flat'
    $btnOpenScripts.FlatAppearance.BorderSize = 0
    $btnOpenScripts.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $btnOpenScripts.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $btnOpenScripts.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOpenScripts.Cursor = 'Hand'
    $btnOpenScripts.Add_Click({
        if (Test-Path $PSScriptRoot) { Start-Process $PSScriptRoot }
    })
    $gInstall.Controls.Add($btnOpenScripts)

    $btnPartsVolumes = New-Object System.Windows.Forms.Button
    $btnPartsVolumes.Text = "Regenerate Parts Volumes List"
    $btnPartsVolumes.Location = New-Object System.Drawing.Point($installColX[0], $installRowY[1])
    $btnPartsVolumes.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnPartsVolumes.FlatStyle = 'Flat'
    $btnPartsVolumes.FlatAppearance.BorderSize = 0
    $btnPartsVolumes.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnPartsVolumes.ForeColor = [System.Drawing.Color]::White
    $btnPartsVolumes.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnPartsVolumes.Cursor = 'Hand'
    $btnPartsVolumes.Add_Click({ [void](Show-PartsVolumesWizard) })
    $gInstall.Controls.Add($btnPartsVolumes)

    $btnSitesWizard = New-Object System.Windows.Forms.Button
    $btnSitesWizard.Text = "Regenerate Sites List"
    $btnSitesWizard.Location = New-Object System.Drawing.Point($installColX[1], $installRowY[1])
    $btnSitesWizard.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnSitesWizard.FlatStyle = 'Flat'
    $btnSitesWizard.FlatAppearance.BorderSize = 0
    $btnSitesWizard.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnSitesWizard.ForeColor = [System.Drawing.Color]::White
    $btnSitesWizard.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSitesWizard.Cursor = 'Hand'
    $btnSitesWizard.Add_Click({ [void](Show-SitesWizard) })
    $gInstall.Controls.Add($btnSitesWizard)

    $btnFigureConvert = New-Object System.Windows.Forms.Button
    $btnFigureConvert.Text = "Convert Figure HTML -> CSV"
    $btnFigureConvert.Location = New-Object System.Drawing.Point($installColX[2], $installRowY[1])
    $btnFigureConvert.Size = New-Object System.Drawing.Size($installBtnW, $installBtnH)
    $btnFigureConvert.FlatStyle = 'Flat'
    $btnFigureConvert.FlatAppearance.BorderSize = 0
    $btnFigureConvert.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnFigureConvert.ForeColor = [System.Drawing.Color]::White
    $btnFigureConvert.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnFigureConvert.Cursor = 'Hand'
    $btnFigureConvert.Add_Click({ Show-FigureConvertDialog })
    $gInstall.Controls.Add($btnFigureConvert)

    # --- Store control references ---
    $script:wtSettings = @{
        TechNameBox  = $techNameBox
        EmailBox     = $emailBox
        EdacUrlBox   = $edacUrlBox
        ThresholdBox = $thresholdBox
        DayLenBox    = $dayLenBox
        LunchBox     = $lunchBox
        WashupBox    = $washupBox
        StartupBox   = $startupBox
        Break1Box    = $break1Box
        Break2Box    = $break2Box
        EndWashBox   = $endWashBox
        PaperBox     = $paperBox
        TargetBox    = $targetBox
    }

    # Wire auto-recompute on any budget input.
    foreach ($box in @($dayLenBox, $lunchBox, $washupBox, $startupBox,
                       $break1Box, $break2Box, $endWashBox, $paperBox)) {
        $box.Add_ValueChanged({ Update-WorkTargetDisplay })
    }

    # --- Initial load ---
    Load-WtSettingsValues

    # --- Save handler ---
    $saveBtn.Add_Click({
        try {
            $c   = $script:wtSettings
            $cfg = $script:config
            if (-not $cfg.WorkTracking) { throw "WorkTracking section missing from config." }

            $techName  = "$($c.TechNameBox.Text)".Trim()
            $email     = "$($c.EmailBox.Text)".Trim()
            $edacUrl   = "$($c.EdacUrlBox.Text)".Trim()
            $threshold = [int]$c.ThresholdBox.Value

            Write-Log "Save: path='$script:configPath' techName='$techName' email='$email'"

            $cfg.WorkTracking.TechnicianName             = $techName
            $cfg.SupervisorEmail                         = $email
            $cfg.WorkTracking.eDacUrl                    = $edacUrl
            $cfg.WorkTracking.ReactiveToWorkOrderMinutes = $threshold

            $cfg.WorkTracking.WorkBudget.DayLengthHours        = [double]$c.DayLenBox.Value
            $cfg.WorkTracking.WorkBudget.LunchMinutes          = [int]$c.LunchBox.Value
            $cfg.WorkTracking.WorkBudget.WashupMinutes         = [int]$c.WashupBox.Value
            $cfg.WorkTracking.WorkBudget.StartupMinutes        = [int]$c.StartupBox.Value
            $cfg.WorkTracking.WorkBudget.PaidBreaksMinutes     = @([int]$c.Break1Box.Value, [int]$c.Break2Box.Value)
            $cfg.WorkTracking.WorkBudget.EndOfDayWashupMinutes = [int]$c.EndWashBox.Value
            $cfg.WorkTracking.WorkBudget.PaperworkMinutes      = [int]$c.PaperBox.Value
            $cfg.WorkTracking.WorkBudget.WorkTargetHours       = [double]$c.TargetBox.Value

            $json = $cfg | ConvertTo-Json -Depth 12
            $json | Set-Content -Path $script:configPath -Encoding UTF8
            $script:config = $cfg

            Write-Log "Save: wrote $($json.Length) chars to $script:configPath"

            [System.Windows.Forms.MessageBox]::Show("Settings saved.", "Settings", "OK", "Information")
        } catch {
            Write-Log "Error saving settings: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error saving settings: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    # --- Reload handler ---
    $reloadBtn.Add_Click({
        try {
            Write-Log "Reload: path='$script:configPath' exists=$(Test-Path $script:configPath)"

            if (-not (Test-Path $script:configPath)) {
                throw "Config file not found at $script:configPath"
            }

            $raw = Get-Content -Path $script:configPath -Raw -Encoding UTF8
            Write-Log "Reload: read $($raw.Length) chars"

            $fresh = $raw | ConvertFrom-Json
            $script:config = $fresh

            $techRead = "$($script:config.WorkTracking.TechnicianName)"
            Write-Log "Reload: parsed techName='$techRead'"

            Load-WtSettingsValues

            Write-Log "Reload: UI updated"
            [System.Windows.Forms.MessageBox]::Show("Settings reloaded.", "Settings", "OK", "Information")
        } catch {
            Write-Log "Error reloading settings: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error reloading settings: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    Write-Log "Settings tab setup completed."
}

# Function to set up the Search tab with enhanced debugging
function Setup-SearchTab {
    param($parentTab, $config)

    Write-Log "Setting up Search tab..."

    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock = 'Fill'
    $searchPanel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $searchPanel.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($searchPanel)

    # ---- Bottom bar (Dock=Bottom) ----
    $bottomBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottomBar.Dock = 'Bottom'
    $bottomBar.Height = 52
    $bottomBar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $bottomBar.WrapContents = $false
    $bottomBar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $bottomBar.BackColor = [System.Drawing.Color]::Transparent
    $searchPanel.Controls.Add($bottomBar)

    # ---- Top search controls (Dock=Top) ----
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.Height = 130
    $topPanel.BackColor = [System.Drawing.Color]::White
    $topPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    $searchPanel.Controls.Add($topPanel)

    # NSN
    $labelNSN = New-Object System.Windows.Forms.Label
    $labelNSN.Text = "NSN:"
    $labelNSN.Location = New-Object System.Drawing.Point(15, 18)
    $labelNSN.Size = New-Object System.Drawing.Size(90, 22)
    $labelNSN.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelNSN.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $topPanel.Controls.Add($labelNSN)

    $script:textBoxNSN = New-Object System.Windows.Forms.TextBox
    $script:textBoxNSN.Location = New-Object System.Drawing.Point(110, 15)
    $script:textBoxNSN.Size = New-Object System.Drawing.Size(250, 25)
    $script:textBoxNSN.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($script:textBoxNSN)

    # OEM
    $labelOEM = New-Object System.Windows.Forms.Label
    $labelOEM.Text = "OEM:"
    $labelOEM.Location = New-Object System.Drawing.Point(15, 52)
    $labelOEM.Size = New-Object System.Drawing.Size(90, 22)
    $labelOEM.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelOEM.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $topPanel.Controls.Add($labelOEM)

    $script:textBoxOEM = New-Object System.Windows.Forms.TextBox
    $script:textBoxOEM.Location = New-Object System.Drawing.Point(110, 49)
    $script:textBoxOEM.Size = New-Object System.Drawing.Size(250, 25)
    $script:textBoxOEM.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($script:textBoxOEM)

    # Description
    $labelDescription = New-Object System.Windows.Forms.Label
    $labelDescription.Text = "Description:"
    $labelDescription.Location = New-Object System.Drawing.Point(15, 86)
    $labelDescription.Size = New-Object System.Drawing.Size(90, 22)
    $labelDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelDescription.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $topPanel.Controls.Add($labelDescription)

    $script:textBoxDescription = New-Object System.Windows.Forms.TextBox
    $script:textBoxDescription.Location = New-Object System.Drawing.Point(110, 83)
    $script:textBoxDescription.Size = New-Object System.Drawing.Size(250, 25)
    $script:textBoxDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($script:textBoxDescription)

    # Search button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(380, 47)
    $searchButton.Size = New-Object System.Drawing.Size(120, 36)
    $searchButton.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $searchButton.ForeColor = [System.Drawing.Color]::White
    $searchButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $searchButton.FlatAppearance.BorderSize = 0
    $searchButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $searchButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $topPanel.Controls.Add($searchButton)

    # Tooltips
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($script:textBoxNSN, "Use * as a wildcard. E.g., 1234*567")
    $tooltip.SetToolTip($script:textBoxOEM, "Use * as a wildcard. E.g., ABC*123")
    $tooltip.SetToolTip($script:textBoxDescription, "Use * as a wildcard in your description search.")

    # ---- Middle: 3 stacked list views via TableLayoutPanel (Dock=Fill) ----
    $listContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $listContainer.Dock = 'Fill'
    $listContainer.ColumnCount = 1
    $listContainer.RowCount = 3
    $listContainer.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 10)
    $listContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.0))) | Out-Null
    $listContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.0))) | Out-Null
    $listContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 34.0))) | Out-Null
    $searchPanel.Controls.Add($listContainer)
    $listContainer.BringToFront()

    # Availability
    $availPanel = New-Object System.Windows.Forms.Panel
    $availPanel.Dock = 'Fill'
    $availPanel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $labelAvailability = New-Object System.Windows.Forms.Label
    $labelAvailability.Text = "Availability"
    $labelAvailability.Dock = 'Top'
    $labelAvailability.Height = 22
    $labelAvailability.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelAvailability.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $availPanel.Controls.Add($labelAvailability)

    $script:listViewAvailability = New-Object System.Windows.Forms.ListView
    $script:listViewAvailability.Dock = 'Fill'
    $script:listViewAvailability.CheckBoxes = $true
    $script:listViewAvailability.Columns.Add("Part (NSN)", 110)          | Out-Null
    $script:listViewAvailability.Columns.Add("Description", 220)         | Out-Null
    $script:listViewAvailability.Columns.Add("QTY", 55)                  | Out-Null
    $script:listViewAvailability.Columns.Add("13 Period Usage", 110)     | Out-Null
    $script:listViewAvailability.Columns.Add("Location", 110)            | Out-Null
    $script:listViewAvailability.Columns.Add("OEM 1", 110)               | Out-Null
    $script:listViewAvailability.Columns.Add("OEM 2", 110)               | Out-Null
    $script:listViewAvailability.Columns.Add("OEM 3", 110)               | Out-Null
    $script:listViewAvailability.Columns.Add("Changed Part (NSN)", 130)  | Out-Null
    Set-ListViewStyle -ListView $script:listViewAvailability
    $availPanel.Controls.Add($script:listViewAvailability)
    $script:listViewAvailability.BringToFront()
    $listContainer.Controls.Add($availPanel, 0, 0)

    # Same Day
    $sameDayPanel = New-Object System.Windows.Forms.Panel
    $sameDayPanel.Dock = 'Fill'
    $sameDayPanel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $labelSameDayAvailability = New-Object System.Windows.Forms.Label
    $labelSameDayAvailability.Text = "Same Day Parts Availability"
    $labelSameDayAvailability.Dock = 'Top'
    $labelSameDayAvailability.Height = 22
    $labelSameDayAvailability.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelSameDayAvailability.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $sameDayPanel.Controls.Add($labelSameDayAvailability)

    $script:listViewSameDayAvailability = New-Object System.Windows.Forms.ListView
    $script:listViewSameDayAvailability.Dock = 'Fill'
    $script:listViewSameDayAvailability.CheckBoxes = $true
    $script:listViewSameDayAvailability.Columns.Add("Part (NSN)", 110)          | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("Description", 220)         | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("QTY", 55)                  | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("13 Period Usage", 110)     | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("Location", 110)            | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("OEM 1", 110)               | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("OEM 2", 110)               | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("OEM 3", 110)               | Out-Null
    $script:listViewSameDayAvailability.Columns.Add("Site Name", 130)           | Out-Null
    Set-ListViewStyle -ListView $script:listViewSameDayAvailability
    $sameDayPanel.Controls.Add($script:listViewSameDayAvailability)
    $script:listViewSameDayAvailability.BringToFront()
    $listContainer.Controls.Add($sameDayPanel, 0, 1)

    # Cross Reference
    $crossRefPanel = New-Object System.Windows.Forms.Panel
    $crossRefPanel.Dock = 'Fill'
    $crossRefPanel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)

    $labelCrossRef = New-Object System.Windows.Forms.Label
    $labelCrossRef.Text = "Cross Reference"
    $labelCrossRef.Dock = 'Top'
    $labelCrossRef.Height = 22
    $labelCrossRef.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $labelCrossRef.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $crossRefPanel.Controls.Add($labelCrossRef)

    $script:listViewCrossRef = New-Object System.Windows.Forms.ListView
    $script:listViewCrossRef.Dock = 'Fill'
    $script:listViewCrossRef.CheckBoxes = $true
    $script:listViewCrossRef.Columns.Add("Handbook", 110)         | Out-Null
    $script:listViewCrossRef.Columns.Add("Section Name", 160)     | Out-Null
    $script:listViewCrossRef.Columns.Add("NO.", 50)               | Out-Null
    $script:listViewCrossRef.Columns.Add("PART DESCRIPTION", 220) | Out-Null
    $script:listViewCrossRef.Columns.Add("REF.", 100)             | Out-Null
    $script:listViewCrossRef.Columns.Add("STOCK NO.", 110)        | Out-Null
    $script:listViewCrossRef.Columns.Add("PART NO.", 110)         | Out-Null
    $script:listViewCrossRef.Columns.Add("CAGE", 60)              | Out-Null
    $script:listViewCrossRef.Columns.Add("Location", 110)         | Out-Null
    $script:listViewCrossRef.Columns.Add("QTY", 55)               | Out-Null
    Set-ListViewStyle -ListView $script:listViewCrossRef
    $crossRefPanel.Controls.Add($script:listViewCrossRef)
    $script:listViewCrossRef.BringToFront()
    $listContainer.Controls.Add($crossRefPanel, 0, 2)

    # ---- Bottom buttons ----
    $script:openFiguresButton = New-Object System.Windows.Forms.Button
    $script:openFiguresButton.Text = "Open Figures"
    $script:openFiguresButton.Size = New-Object System.Drawing.Size(140, 36)
    $script:openFiguresButton.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $script:openFiguresButton.ForeColor = [System.Drawing.Color]::White
    $script:openFiguresButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:openFiguresButton.FlatAppearance.BorderSize = 0
    $script:openFiguresButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:openFiguresButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:openFiguresButton.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
    $bottomBar.Controls.Add($script:openFiguresButton)

    $script:takePartOutButton = New-Object System.Windows.Forms.Button
    $script:takePartOutButton.Text = "Take Part(s) Out"
    $script:takePartOutButton.Size = New-Object System.Drawing.Size(140, 36)
    $script:takePartOutButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $script:takePartOutButton.ForeColor = [System.Drawing.Color]::White
    $script:takePartOutButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:takePartOutButton.FlatAppearance.BorderSize = 0
    $script:takePartOutButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:takePartOutButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:takePartOutButton.Enabled = $false
    $bottomBar.Controls.Add($script:takePartOutButton)

    # ---------- Event handlers ----------
    $takePartOutButton = $script:takePartOutButton
    $openFiguresButton = $script:openFiguresButton

    $takePartOutButton.Add_Click({
        $selectedParts = $script:listViewAvailability.CheckedItems + $script:listViewSameDayAvailability.CheckedItems | ForEach-Object { $_.SubItems[0].Text }
        if ($selectedParts) {
            Write-Log "Selected parts to take out: $($selectedParts -join ', ')"
            [System.Windows.Forms.MessageBox]::Show("Selected parts to take out: $($selectedParts -join ', ')", "Parts Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            Write-Log "No parts selected to take out."
            [System.Windows.Forms.MessageBox]::Show("No parts selected to take out.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $openFiguresButton.Add_Click({
        $checkedItems = $script:listViewCrossRef.CheckedItems
        if ($checkedItems.Count -gt 0) {
            foreach ($item in $checkedItems) {
                $handbook = $item.SubItems[0].Text
                $ref = $item.SubItems[4].Text

                if ($handbook -and $ref) {
                    $bookProp = $config.Books.PSObject.Properties[$handbook]
                    if ($bookProp -and $bookProp.Value.VolumesToUrlCsvPath) {
                        $bookDir = Split-Path -Path $bookProp.Value.VolumesToUrlCsvPath -Parent
                    } else {
                        $bookDir = Join-Path $config.PartsBooksDirectory $handbook
                    }
                    $htmlFilePath = Join-Path $bookDir "HTML and CSV Files\$ref.html"

                    if (Test-Path $htmlFilePath) {
                        Start-Process $htmlFilePath
                        Write-Log "Opened figure: $htmlFilePath"
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Figure file not found: $htmlFilePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        Write-Log "Figure file not found: $htmlFilePath"
                    }
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one item to open figures.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    $script:listViewAvailability.Add_ItemChecked({
        $script:takePartOutButton.Enabled = ($script:listViewAvailability.CheckedItems.Count -gt 0 -or $script:listViewSameDayAvailability.CheckedItems.Count -gt 0)
    })

    $script:listViewSameDayAvailability.Add_ItemChecked({
        $script:takePartOutButton.Enabled = ($script:listViewAvailability.CheckedItems.Count -gt 0 -or $script:listViewSameDayAvailability.CheckedItems.Count -gt 0)
    })

    $script:listViewCrossRef.Add_ItemChecked({
        $script:openFiguresButton.Enabled = ($script:listViewCrossRef.CheckedItems.Count -gt 0)
    })

    # ---------------- Search logic (unchanged) ----------------
    $searchButton.Add_Click({
        Write-Log "Performing part search..."

        $nsnSearch         = $script:textBoxNSN.Text.Trim()
        $oemSearch         = $script:textBoxOEM.Text.Trim()
        $descriptionSearch = $script:textBoxDescription.Text.Trim()
        Write-Log "Search criteria - NSN: $nsnSearch, OEM: $oemSearch, Description: $descriptionSearch"

        # NSN pattern
        $nsnActive  = $false
        $nsnNoMatch = $false
        $nsnPattern = $null
        if (-not [string]::IsNullOrWhiteSpace($nsnSearch)) {
            $nsnCleaned = $nsnSearch -replace '[^0-9*]', ''
            if ([string]::IsNullOrEmpty($nsnCleaned) -or $nsnCleaned -eq '*') {
                $nsnNoMatch = $true
            } else {
                $nsnPattern = ([regex]::Escape($nsnCleaned)) -replace '\\\*', '.*'
                $nsnActive  = $true
            }
        }

        # OEM pattern
        $oemActive  = $false
        $oemNoMatch = $false
        $oemPattern = $null
        if (-not [string]::IsNullOrWhiteSpace($oemSearch)) {
            $oemCleaned = ($oemSearch -replace '[^A-Za-z0-9*]', '').ToUpper()
            if ([string]::IsNullOrEmpty($oemCleaned) -or $oemCleaned -eq '*') {
                $oemNoMatch = $true
            } else {
                $oemPattern = ([regex]::Escape($oemCleaned)) -replace '\\\*', '.*'
                $oemActive  = $true
            }
        }

        # Description tokens
        $descTokens = @()
        if (-not [string]::IsNullOrWhiteSpace($descriptionSearch)) {
            $descTokens = @(
                ($descriptionSearch -replace '[^A-Za-z0-9]', ' ').ToLower() -split '\s+' |
                Where-Object { $_ -ne '' }
            )
        }
        $descActive = $descTokens.Count -gt 0

        Write-Log "Patterns: NSN='$nsnPattern' active=$nsnActive noMatch=$nsnNoMatch ; OEM='$oemPattern' active=$oemActive noMatch=$oemNoMatch ; DescTokens=$($descTokens -join ',')"

        # ---- 1) AVAILABILITY ----
        $csvFiles = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.csv" -File
        if ($csvFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No CSV files found in $($config.PartsRoomDirectory).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "No CSV files found in $($config.PartsRoomDirectory)."
            return
        } elseif ($csvFiles.Count -gt 1) {
            [System.Windows.Forms.MessageBox]::Show("Multiple CSV files found in $($config.PartsRoomDirectory). Please ensure only one CSV file is present.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Multiple CSV files found in $($config.PartsRoomDirectory)."
            return
        } else {
            $csvFilePath = $csvFiles[0].FullName
            Write-Log "Using CSV file: $csvFilePath"
        }

        try {
            $data = @(Import-Csv -Path $csvFilePath)
            Write-Log "CSV file loaded successfully. Row count: $($data.Count)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to read CSV file. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Failed to read CSV file. Error: $_"
            return
        }

        if ($nsnNoMatch -or $oemNoMatch) {
            $filteredData = @()
        } else {
            $filteredData = @($data | Where-Object {
                $nsnOk = $true
                if ($nsnActive) {
                    $rowNSN = ([string]$_.'Part (NSN)') -replace '[^0-9]', ''
                    $nsnOk = $rowNSN -match $nsnPattern
                }
                $oemOk = $true
                if ($oemActive) {
                    $oemOk = $false
                    foreach ($f in 'OEM 1','OEM 2','OEM 3') {
                        $rowOem = ([string]$_.$f -replace '[^A-Za-z0-9]', '').ToUpper()
                        if ($rowOem -and $rowOem -match $oemPattern) { $oemOk = $true; break }
                    }
                }
                $descOk = $true
                if ($descActive) {
                    $rowDesc = (([string]$_.Description -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                    foreach ($token in $descTokens) {
                        if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                    }
                }
                $nsnOk -and $oemOk -and $descOk
            })
        }
        if ($null -eq $filteredData) { $filteredData = @() }
        Write-Log "Found $($filteredData.Count) matching records in Availability."

        $script:listViewAvailability.Items.Clear()
        foreach ($row in $filteredData) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$row.'Part (NSN)')
            $item.SubItems.Add([string]$row.Description)
            $item.SubItems.Add([string]$row.QTY)
            $item.SubItems.Add([string]$row.'13 Period Usage')
            $item.SubItems.Add([string]$row.Location)
            $item.SubItems.Add([string]$row.'OEM 1')
            $item.SubItems.Add([string]$row.'OEM 2')
            $item.SubItems.Add([string]$row.'OEM 3')
            if ($row.PSObject.Properties.Name -contains 'Changed Part (NSN)') {
                $item.SubItems.Add([string]$row.'Changed Part (NSN)')
            } else {
                $item.SubItems.Add("")
            }
            $script:listViewAvailability.Items.Add($item) | Out-Null
        }

        # ---- 2) SAME DAY ----
        $sameDayPartsDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
        if (Test-Path $sameDayPartsDir) {
            $sameDayCsvFiles = Get-ChildItem -Path $sameDayPartsDir -Filter "*.csv" -File
            Write-Log "Found $($sameDayCsvFiles.Count) CSV files in Same Day Parts Room."

            $sameDayData = @()
            foreach ($csvFile in $sameDayCsvFiles) {
                $siteName = [IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
                try {
                    $csvData = Import-Csv -Path $csvFile.FullName
                    foreach ($row in $csvData) {
                        $row | Add-Member -NotePropertyName 'SiteName' -NotePropertyValue $siteName -Force
                        $sameDayData += $row
                    }
                } catch {
                    Write-Log "Failed to read CSV file $($csvFile.FullName). Error: $_"
                }
            }

            if ($nsnNoMatch -or $oemNoMatch) {
                $filteredSameDayData = @()
            } else {
                $filteredSameDayData = @($sameDayData | Where-Object {
                    $nsnOk = $true
                    if ($nsnActive) {
                        $rowNSN = ([string]$_.'Part (NSN)') -replace '[^0-9]', ''
                        $nsnOk = $rowNSN -match $nsnPattern
                    }
                    $oemOk = $true
                    if ($oemActive) {
                        $oemOk = $false
                        foreach ($f in 'OEM 1','OEM 2','OEM 3') {
                            $rowOem = ([string]$_.$f -replace '[^A-Za-z0-9]', '').ToUpper()
                            if ($rowOem -and $rowOem -match $oemPattern) { $oemOk = $true; break }
                        }
                    }
                    $descOk = $true
                    if ($descActive) {
                        $rowDesc = (([string]$_.Description -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                        foreach ($token in $descTokens) {
                            if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                        }
                    }
                    $nsnOk -and $oemOk -and $descOk
                })
            }
            if ($null -eq $filteredSameDayData) { $filteredSameDayData = @() }
            Write-Log "Found $($filteredSameDayData.Count) matching records in Same Day Parts Availability."

            $script:listViewSameDayAvailability.Items.Clear()
            foreach ($row in $filteredSameDayData) {
                $item = New-Object System.Windows.Forms.ListViewItem([string]$row.'Part (NSN)')
                $item.SubItems.Add([string]$row.Description)
                $item.SubItems.Add([string]$row.QTY)
                $item.SubItems.Add([string]$row.'13 Period Usage')
                $item.SubItems.Add([string]$row.Location)
                $item.SubItems.Add([string]$row.'OEM 1')
                $item.SubItems.Add([string]$row.'OEM 2')
                $item.SubItems.Add([string]$row.'OEM 3')
                $item.SubItems.Add([string]$row.SiteName)
                $script:listViewSameDayAvailability.Items.Add($item) | Out-Null
            }
        } else {
            Write-Log "Same Day Parts Room directory not found at $sameDayPartsDir"
        }

        # ---- 3) CROSS REFERENCE ----
        $crossRefResults = @()
        if ($config.Books) {
            foreach ($book in $config.Books.PSObject.Properties) {
                $bookName = $book.Name
                $volumesCsvPath = $book.Value.VolumesToUrlCsvPath
                if ($volumesCsvPath) {
                    $bookDir = Split-Path -Path $volumesCsvPath -Parent
                } else {
                    $bookDir = Join-Path $config.PartsBooksDirectory $bookName
                }

                $combinedSectionsDir = Join-Path $bookDir "CombinedSections"
                $sectionNamesFile    = Join-Path $bookDir "SectionNames.txt"

                $sectionNameMapping = @{}
                if (Test-Path $sectionNamesFile) {
                    $sectionNames = Get-Content -Path $sectionNamesFile
                    foreach ($line in $sectionNames) {
                        if ($line -match '^Section\s+(\d+)\s+(.*)$') {
                            $sectionNumber   = $Matches[1]
                            $sectionFullName = $line.Trim()
                            $sectionNameMapping["Section $sectionNumber"] = $sectionFullName
                        }
                    }
                    Write-Log "Loaded $($sectionNameMapping.Count) section names for $bookName"
                } else {
                    Write-Log "SectionNames.txt not found for $bookName"
                }

                if (-not (Test-Path $combinedSectionsDir)) {
                    Write-Log "CombinedSections directory not found for $bookName"
                    continue
                }

                $sectionCsvFiles = Get-ChildItem -Path $combinedSectionsDir -Filter "*.csv" -File
                Write-Log "Found $($sectionCsvFiles.Count) CSV files in $bookName"

                foreach ($csvFile in $sectionCsvFiles) {
                    $sectionFileName = [IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
                    $csvFilePath = $csvFile.FullName

                    if ($sectionFileName -match '^Section\s+(\d+)$') {
                        $sectionNumber = $Matches[1]
                        if ($sectionNameMapping.ContainsKey("Section $sectionNumber")) {
                            $sectionName = $sectionNameMapping["Section $sectionNumber"]
                        } else {
                            $sectionName = "Section $sectionNumber"
                        }
                    } else {
                        $sectionName = $sectionFileName
                    }

                    try {
                        $sectionData = Import-Csv -Path $csvFilePath
                        Write-Log "Processed $($sectionData.Count) rows from $($csvFile.Name)"
                    } catch {
                        Write-Log "Failed to read CSV file $csvFilePath. Error: $_"
                        continue
                    }

                    if ($nsnNoMatch -or $oemNoMatch) {
                        $filteredSectionData = @()
                    } else {
                        $filteredSectionData = @($sectionData | Where-Object {
                            $nsnOk = $true
                            if ($nsnActive) {
                                $rowStock = ([string]$_.'STOCK NO.') -replace '[^0-9]', ''
                                $nsnOk = $rowStock -match $nsnPattern
                            }
                            $oemOk = $true
                            if ($oemActive) {
                                $rowPart = ([string]$_.'PART NO.' -replace '[^A-Za-z0-9]', '').ToUpper()
                                $oemOk = ($rowPart -and ($rowPart -match $oemPattern))
                            }
                            $descOk = $true
                            if ($descActive) {
                                $rowDesc = (([string]$_.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                                foreach ($token in $descTokens) {
                                    if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                                }
                            }
                            $nsnOk -and $oemOk -and $descOk
                        })
                    }

                    foreach ($item in $filteredSectionData) {
                        $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $bookName    -Force
                        $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $sectionName -Force
                        $crossRefResults += $item
                    }
                }
            }
        } else {
            Write-Log "No books defined in configuration."
        }

        Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference."

        $script:listViewCrossRef.Items.Clear()
        if ($crossRefResults.Count -gt 0) {
            $columns = @('Handbook', 'Section Name', 'NO.', 'PART DESCRIPTION', 'REF.', 'STOCK NO.', 'PART NO.', 'CAGE', 'Location', 'QTY')
            foreach ($row in $crossRefResults) {
                $item = New-Object System.Windows.Forms.ListViewItem($row.Handbook)
                foreach ($column in $columns[1..($columns.Count - 1)]) {
                    $value = if ($row.PSObject.Properties.Name -contains $column) { $row.$column } else { "" }
                    if ($column -eq 'REF.' -and $value -is [string]) {
                        $value = $value -replace '\.csv$', ''
                    }
                    $item.SubItems.Add($value)
                }
                $script:listViewCrossRef.Items.Add($item) | Out-Null
            }
            $script:listViewCrossRef.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
            Write-Log "Added $($crossRefResults.Count) items to Cross Reference ListView."
        } else {
            [System.Windows.Forms.MessageBox]::Show("No matching records found in Cross Reference.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Log "No matching records found in Cross Reference."
        }

        Write-Log "Search complete."
    })
    Write-Log "Search tab setup completed."
}

# Function to get or set the supervisor email
function Get-SupervisorEmail {
    if (-not $config.SupervisorEmail) {
        $emailForm = New-Object System.Windows.Forms.Form
        $emailForm.Text = "Supervisor Email"
        $emailForm.Size = New-Object System.Drawing.Size(300, 150)
        $emailForm.StartPosition = 'CenterScreen'

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = "Please enter the supervisor's email address:"
        $emailForm.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,50)
        $textBox.Size = New-Object System.Drawing.Size(260,20)
        $emailForm.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(100,80)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $emailForm.Controls.Add($okButton)

        $emailForm.AcceptButton = $okButton

        $result = $emailForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $email = $textBox.Text.Trim()
            if ($email -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
                $config | Add-Member -NotePropertyName SupervisorEmail -NotePropertyValue $email -Force
                $config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
                Write-Log "Supervisor email updated to: $email"
                return $email
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid email format. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return Get-SupervisorEmail
            }
        } else {
            return $null
        }
    } else {
        return $config.SupervisorEmail
    }
}

################################################################################
#                     Parts Room and Books Management                          #
################################################################################

# Function to update Parts Books with the latest inventory data
function Update-PartsBooks {
    param(
        [string]$sourceCSVPath
    )

    Write-Log "Starting Parts Books update process..."

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Updating Parts Books"
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = 'CenterScreen'

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(370, 20)
    $progressLabel.Text = "Loading source data..."
    $progressForm.Controls.Add($progressLabel)

	$progressBar = New-Object ModernProgressBar
	$progressBar.Location = New-Object System.Drawing.Point(10, 50)
	$progressBar.Size = New-Object System.Drawing.Size(370, 26)
	$progressForm.Controls.Add($progressBar)

    $bookLabel = New-Object System.Windows.Forms.Label
    $bookLabel.Location = New-Object System.Drawing.Point(10, 80)
    $bookLabel.Size = New-Object System.Drawing.Size(370, 20)
    $bookLabel.Text = ""
    $progressForm.Controls.Add($bookLabel)

    $progressForm.Show()
    $progressForm.Refresh()

    if (-not (Test-Path $sourceCSVPath)) {
        Write-Log "Error: Source CSV file not found at $sourceCSVPath"
        [System.Windows.Forms.MessageBox]::Show("Source CSV file not found at $sourceCSVPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }

    $sourceData = Import-Csv -Path $sourceCSVPath

    $nsnDict = @{}
    $oemDict = @{}

    foreach ($part in $sourceData) {
        if (-not [string]::IsNullOrEmpty($part.'Part (NSN)')) {
            $nsnDict[$part.'Part (NSN)'] = $part
        }
        if (-not [string]::IsNullOrEmpty($part.'OEM 1')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 1'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) { $oemDict[$normalizedOEM] = $part }
        }
        if (-not [string]::IsNullOrEmpty($part.'OEM 2')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 2'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) { $oemDict[$normalizedOEM] = $part }
        }
        if (-not [string]::IsNullOrEmpty($part.'OEM 3')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 3'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) { $oemDict[$normalizedOEM] = $part }
        }
    }

    $progressLabel.Text = "Loaded $($sourceData.Count) parts from source CSV"
    $progressForm.Refresh()
    Write-Log "Loaded $($sourceData.Count) parts, $($nsnDict.Count) unique NSNs, $($oemDict.Count) unique OEMs"

    if (-not $config.Books -or $config.Books.PSObject.Properties.Count -eq 0) {
        Write-Log "No parts books found in configuration"
        [System.Windows.Forms.MessageBox]::Show("No parts books found in configuration", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }

    # Collect books to process. Derive the folder from the config's stored
    # VolumesToUrlCsvPath so folder names containing ( ) . work.
    $booksToProcess = @()
    foreach ($bookProp in $config.Books.PSObject.Properties) {
        $bookName = $bookProp.Name
        $volumesCsvPath = $bookProp.Value.VolumesToUrlCsvPath
        if ($volumesCsvPath) {
            $bookDir = Split-Path -Path $volumesCsvPath -Parent
        } else {
            $bookDir = Join-Path $config.PartsBooksDirectory $bookName
        }

        $combinedSectionsDir = Join-Path $bookDir "CombinedSections"
        $excelFile = Get-ChildItem -Path $bookDir -Filter "*.xlsx" -ErrorAction SilentlyContinue |
                     Select-Object -First 1 -ExpandProperty FullName

        if (Test-Path $combinedSectionsDir) {
            $booksToProcess += @{
                Name = $bookName
                Directory = $bookDir
                CombinedSectionsDir = $combinedSectionsDir
                ExcelPath = $excelFile
            }
        } else {
            Write-Log "Warning: CombinedSections directory not found for book: $bookName (looked in $bookDir)"
        }
    }

    $totalBooks = $booksToProcess.Count
    if ($totalBooks -eq 0) {
        Write-Log "No parts books with combined sections found to update"
        [System.Windows.Forms.MessageBox]::Show("No parts books with combined sections found to update", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }

    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $totalUpdatedCSVs = 0
    $totalUpdatedParts = 0

    for ($bookIndex = 0; $bookIndex -lt $totalBooks; $bookIndex++) {
        $book = $booksToProcess[$bookIndex]
        $progressBar.Value = [Math]::Min([int](($bookIndex / [double]$totalBooks) * 100), 100)
        $bookLabel.Text = "Processing book: $($book.Name)"
        $progressLabel.Text = "Scanning sections..."
        $progressForm.Refresh()

        Write-Log "Processing book: $($book.Name)"

        $sectionFiles = Get-ChildItem -Path $book.CombinedSectionsDir -Filter "Section *.csv"
        if ($sectionFiles.Count -eq 0) {
            Write-Log "No section CSV files found for book: $($book.Name)"
            continue
        }

        $sectionsWithChanges = @()
        $bookUpdatedParts = 0

        foreach ($sectionFile in $sectionFiles) {
            $sectionName = $sectionFile.BaseName
            $progressLabel.Text = "Processing section: $sectionName"
            $progressForm.Refresh()

            try {
                $sectionData = Import-Csv -Path $sectionFile.FullName
                $sectionUpdated = $false
                $sectionUpdatedParts = 0

                foreach ($part in $sectionData) {
                    $stockNo = $part.'STOCK NO.'
                    $partNo = $part.'PART NO.'

                    if (-not [string]::IsNullOrEmpty($stockNo) -and $stockNo -ne "NSL") {
                        if ($nsnDict.ContainsKey($stockNo)) {
                            $sourcePart = $nsnDict[$stockNo]
                            if ($part.QTY -ne $sourcePart.QTY -or $part.Location -ne $sourcePart.Location) {
                                $part.QTY = $sourcePart.QTY
                                $part.Location = $sourcePart.Location
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        } else {
                            if ($part.QTY -ne "0" -or $part.Location -ne "Not in current inventory") {
                                $part.QTY = "0"
                                $part.Location = "Not in current inventory"
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        }
                    }
                    elseif (-not [string]::IsNullOrEmpty($partNo)) {
                        $normalizedPartNo = Normalize-OEM -oem $partNo
                        if (-not [string]::IsNullOrEmpty($normalizedPartNo) -and $oemDict.ContainsKey($normalizedPartNo)) {
                            $sourcePart = $oemDict[$normalizedPartNo]
                            if ($part.QTY -ne $sourcePart.QTY -or $part.Location -ne $sourcePart.Location) {
                                $part.QTY = $sourcePart.QTY
                                $part.Location = $sourcePart.Location
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        } else {
                            if ($part.QTY -ne "0" -or $part.Location -ne "Not in current inventory") {
                                $part.QTY = "0"
                                $part.Location = "Not in current inventory"
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        }
                    }
                }

                if ($sectionUpdated) {
                    $sectionData | Export-Csv -Path $sectionFile.FullName -NoTypeInformation
                    $sectionsWithChanges += $sectionName
                    $bookUpdatedParts += $sectionUpdatedParts
                    $totalUpdatedCSVs++
                    $totalUpdatedParts += $sectionUpdatedParts
                    Write-Log "Updated section $sectionName with $sectionUpdatedParts changes"
                }
            } catch {
                Write-Log "Error processing section $sectionName : $($_.Exception.Message)"
            }
        }

        if ($sectionsWithChanges.Count -gt 0 -and $book.ExcelPath -and (Test-Path $book.ExcelPath)) {
            $progressLabel.Text = "Updating Excel workbook..."
            $bookLabel.Text = "Processing book: $($book.Name) - Excel update"
            $progressForm.Refresh()

            $excel = $null
            $workbook = $null

            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false

                $workbook = $excel.Workbooks.Open($book.ExcelPath)

                $baseProgress = $bookIndex / $totalBooks * 100
                $progressWeight = 100 / $totalBooks / 2
                $totalOperations = $sectionsWithChanges.Count * 10
                $currentOperation = 0

                foreach ($sectionName in $sectionsWithChanges) {
                    try {
                        $currentOperation += 5
                        $sectionProgress = $currentOperation / $totalOperations * $progressWeight
                        $totalProgress = $baseProgress + $sectionProgress
                        $progressBar.Value = [Math]::Min([int]$totalProgress, 100)
                        $progressLabel.Text = "Processing section: $sectionName"
                        $progressForm.Refresh()

                        $possibleNames = @()
                        $possibleNames += $sectionName
                        $truncatedName = $sectionName.Substring(0, [Math]::Min(31, $sectionName.Length)) -replace '[:\\/?*\[\]]', ''
                        $possibleNames += $truncatedName
                        if ($sectionName -match '^(Section \d+)') {
                            $possibleNames += $matches[1]
                        }

                        $worksheet = $null
                        foreach ($nameVariant in $possibleNames) {
                            try {
                                $worksheet = $workbook.Worksheets.Item($nameVariant)
                                Write-Log "Found worksheet using name variant: $nameVariant"
                                break
                            } catch {
                                continue
                            }
                        }

                        if ($worksheet -eq $null) {
                            Write-Log "Could not find exact worksheet match for $sectionName, trying fuzzy matching..."
                            foreach ($ws in $workbook.Worksheets) {
                                if ($sectionName -match '^(Section \d+)' -and
                                    $ws.Name -match "^$($matches[1])") {
                                    $worksheet = $ws
                                    Write-Log "Found worksheet using fuzzy match: $($ws.Name)"
                                    break
                                }
                            }
                        }

                        $currentOperation += 5
                        $progressBar.Value = [Math]::Min([int]($baseProgress + $currentOperation / $totalOperations * $progressWeight), 100)
                        $progressForm.Refresh()

                        if ($worksheet -ne $null) {
                            $qtyCol = $null
                            $locationCol = $null
                            $lastCol = 20
                            for ($col = 1; $col -le $lastCol; $col++) {
                                $colName = $worksheet.Cells.Item(1, $col).Text
                                if ($colName -eq "QTY") { $qtyCol = $col }
                                elseif ($colName -eq "LOCATION") { $locationCol = $col }
                                if ($qtyCol -and $locationCol) { break }
                            }

                            $stockNoCol = $null
                            $partNoCol = $null
                            for ($col = 1; $col -le $lastCol; $col++) {
                                $colName = $worksheet.Cells.Item(1, $col).Text
                                if ($colName -eq "STOCK NO.") { $stockNoCol = $col }
                                elseif ($colName -eq "PART NO.") { $partNoCol = $col }
                                if ($stockNoCol -and $partNoCol) { break }
                            }

                            if ($stockNoCol -or $partNoCol) {
                                if ($qtyCol -and $locationCol) {
                                    $sectionFile = Get-ChildItem -Path $book.CombinedSectionsDir -Filter "$sectionName.csv" | Select-Object -First 1
                                    if ($sectionFile) {
                                        $sectionData = Import-Csv -Path $sectionFile.FullName

                                        $sectionDict = @{}
                                        foreach ($part in $sectionData) {
                                            $stockNo = $part.'STOCK NO.'
                                            if (-not [string]::IsNullOrEmpty($stockNo)) { $sectionDict[$stockNo] = $part }
                                            $partNo = $part.'PART NO.'
                                            if (-not [string]::IsNullOrEmpty($partNo)) { $sectionDict[$partNo] = $part }
                                        }

                                        $lastRow = $worksheet.UsedRange.Rows.Count
                                        for ($row = 2; $row -le $lastRow; $row++) {
                                            if (($row % 10) -eq 0 -or $row -eq $lastRow) {
                                                $rowProgress = ($row - 2) / ($lastRow - 2) * 10
                                                $currentOperation += $rowProgress
                                                $cellProgress = $currentOperation / $totalOperations * $progressWeight
                                                $totalProgress = $baseProgress + $cellProgress
                                                $progressBar.Value = [Math]::Min([int]$totalProgress, 100)
                                                $progressLabel.Text = "Processing section: $sectionName - Row $row of $lastRow"
                                                $progressForm.Refresh()
                                                [System.Windows.Forms.Application]::DoEvents()
                                            }

                                            $key = $null
                                            if ($stockNoCol) {
                                                $stockNo = $worksheet.Cells.Item($row, $stockNoCol).Text
                                                if (-not [string]::IsNullOrEmpty($stockNo)) { $key = $stockNo }
                                            }
                                            if (-not $key -and $partNoCol) {
                                                $partNo = $worksheet.Cells.Item($row, $partNoCol).Text
                                                if (-not [string]::IsNullOrEmpty($partNo)) { $key = $partNo }
                                            }
                                            if ($key -and $sectionDict.ContainsKey($key)) {
                                                $part = $sectionDict[$key]
                                                $worksheet.Cells.Item($row, $qtyCol).Value2 = $part.QTY
                                                $worksheet.Cells.Item($row, $locationCol).Value2 = $part.Location
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Log "Error updating worksheet ${sectionName}: $($_.Exception.Message)"
                    }
                }

                $progressBar.Value = [Math]::Min([int]($baseProgress + $progressWeight), 100)
                $progressLabel.Text = "Saving Excel workbook..."
                $progressForm.Refresh()
                $workbook.Save()
                $workbook.Close($false)
                $workbook = $null
                $excel.Quit()
                $excel = $null
            } catch {
                Write-Log "Error updating Excel workbook: $($_.Exception.Message)"
            } finally {
                if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
                if ($null -ne $excel) {
                    try { $excel.Quit() } catch { }
                    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }

        Write-Log "Completed updating book $($book.Name) - Updated $bookUpdatedParts parts"
    }

    $progressForm.Close()

    $message = "Parts Books update completed:`n"
    $message += "- Updated $totalUpdatedParts parts`n"
    $message += "- Updated $totalUpdatedCSVs section CSV files`n"
    $message += "- Processed $totalBooks books"

    Write-Log $message
    [System.Windows.Forms.MessageBox]::Show($message, "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    return $true
}

function Update-PartsRoom {
    Write-Log "Starting Parts Room update process..."
    
    # Get the configuration
    if (-not $config) {
        Write-Log "Error: Configuration not available"
        [System.Windows.Forms.MessageBox]::Show("Configuration not available. Cannot update Parts Room.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Check if a previously selected site exists in the config
    $selectedSiteName = $null
    $selectedSiteID = $null
    
    # Check parts room directory for existing CSV files
    $existingCsvFile = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.csv" | Select-Object -First 1
    if ($existingCsvFile) {
        $selectedSiteName = [System.IO.Path]::GetFileNameWithoutExtension($existingCsvFile.Name)
        Write-Log "Found existing Parts Room data for site: $selectedSiteName"
        
        # Get the site ID from Sites.csv
        $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
        if (Test-Path $sitesPath) {
            $sites = Import-Csv -Path $sitesPath
            $SiteIDColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
            $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
            
            $matchingSite = $sites | Where-Object { $_.$fullNameColumn -eq $selectedSiteName }
            if ($matchingSite) {
                $selectedSiteID = $matchingSite.$SiteIDColumn
                Write-Log "Found Site ID for ${selectedSiteName}: ${selectedSiteID}"
            }
        }
    }
    
    # Ask user if they want to update the existing site or select a new one
    $dialogResult = [System.Windows.Forms.DialogResult]::Yes
    if ($selectedSiteName) {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to update the existing Parts Room data for $selectedSiteName?`n`nClick 'Yes' to update the existing site.`nClick 'No' to select a different site.",
            "Update Parts Room",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)
    }
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Log "Parts Room update cancelled by user"
        return
    }
    
    # If user wants to select a new site, or no existing site was found
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::No -or -not $selectedSiteName) {
        # Load the sites from the CSV
        $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
        Write-Log "Loading Sites from CSV at $sitesPath..."
        
        if (-not (Test-Path $sitesPath)) {
            Write-Log "Error: Sites.csv not found at $sitesPath"
            [System.Windows.Forms.MessageBox]::Show("Sites CSV file not found. Cannot update Parts Room.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $sites = Import-Csv -Path $sitesPath
        
        # Check if the CSV was loaded successfully
        if ($sites.Count -eq 0) {
            Write-Log "Error: No data found in the Sites.csv file"
            [System.Windows.Forms.MessageBox]::Show("No data found in the Sites.csv file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Determine the correct column names
        $SiteIDColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
        $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
        Write-Log "Site ID Column: $SiteIDColumn, Full Name Column: $fullNameColumn"
        
        # Create the form for site selection
        Write-Log "Creating form for site selection..."
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select Facility"
        $form.Size = New-Object System.Drawing.Size(400, 200)
        $form.StartPosition = "CenterScreen"
        
        # Create the dropdown menu
        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Location = New-Object System.Drawing.Point(10, 20)
        $dropdown.Size = New-Object System.Drawing.Size(360, 20)
        $dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        
        # Populate the dropdown with the full names
        foreach ($site in $sites) {
            $fullName = $site.$fullNameColumn
            if (![string]::IsNullOrWhiteSpace($fullName)) {
                $dropdown.Items.Add($fullName)
            }
        }
        
        Write-Log "Dropdown populated with site names."
        $form.Controls.Add($dropdown)
        
        # Create the "Select" button
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(150, 60)
        $button.Size = New-Object System.Drawing.Size(75, 23)
        $button.Text = "Select"
        $button.Add_Click({
            $form.Tag = $dropdown.SelectedItem
            $form.Close()
        })
        $form.Controls.Add($button)
        
        # Show the form and get the selected site
        Write-Log "Showing form for site selection..."
        $form.ShowDialog()
        $selectedSiteName = $form.Tag
        
        if (-not $selectedSiteName) {
            Write-Log "No site selected, cancelling operation."
            return
        }
        
        Write-Log "Selected Site: $selectedSiteName"
        $selectedRow = $sites | Where-Object { $_.$fullNameColumn -eq $selectedSiteName }
        $selectedSiteID = $selectedRow.$SiteIDColumn
    }
    
    # Now we have a site name and ID, proceed with downloading and processing the data
    if (-not $selectedSiteID) {
        Write-Log "Error: Cannot determine site ID for $selectedSiteName"
        [System.Windows.Forms.MessageBox]::Show("Cannot determine site ID for $selectedSiteName. Update cancelled.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Construct the URL for the selected site
    $url = "http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$selectedSiteID&p_search_type=DESC&p_search_string=&p_boh_radio=-1"
    Write-Log "URL for site ${selectedSiteName}: $url"
    
    # Download HTML content
    Write-Log "Downloading HTML content for $selectedSiteName..."
    try {
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Updating Parts Room"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = 'CenterScreen'
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
        $progressLabel.Size = New-Object System.Drawing.Size(370, 20)
        $progressLabel.Text = "Downloading data for $selectedSiteName..."
        $progressForm.Controls.Add($progressLabel)
        
		$progressBar = New-Object ModernProgressBar
		$progressBar.Location = New-Object System.Drawing.Point(10, 50)
		$progressBar.Size = New-Object System.Drawing.Size(370, 26)
		$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
		$progressForm.Controls.Add($progressBar)
        
        # Show the progress form
        $progressForm.Show()
        $progressForm.Refresh()
        
        # Download the HTML content
        $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
        $htmlFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.html"
        Write-Log "Saving HTML content to $htmlFilePath..."
        Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8
        
        # Update the progress label
        $progressLabel.Text = "Processing HTML data..."
        $progressForm.Refresh()
        
        # Process the HTML content
        Write-Log "Processing the downloaded HTML file for $selectedSiteName using DOM..."
        
        $logPath = Join-Path $config.PartsRoomDirectory "error_log.txt"
        $htmlDoc = $null
        
        try {
            # Create a COM object for HTML Document
            $htmlDoc = New-Object -ComObject "HTMLFile"
            
            # Read the HTML file content
            $htmlFileContent = Get-Content -Path $htmlFilePath -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($htmlFileContent)) {
                throw "HTML file content is empty or null"
            }
            
            # Attempt to load the HTML
            try {
                # For newer PowerShell versions
                $htmlDoc.IHTMLDocument2_write($htmlFileContent)
            } catch {
                # Fallback method for older PowerShell versions
                $src = [System.Text.Encoding]::Unicode.GetBytes($htmlFileContent)
                $htmlDoc.write($src)
            }
            
            Write-Log "HTML document loaded into DOM successfully."
            
            # Find the main table with the parts data
            $tables = $htmlDoc.getElementsByTagName("table")
            $mainTable = $null
            
            Write-Log "Found $($tables.length) tables in the document."
            
            # Try multiple methods to find the main table
            foreach ($table in $tables) {
                # Try with className
                if ($table.className -eq "MAIN") {
                    $mainTable = $table
                    Write-Log "Found main table using className property."
                    break
                }
                
                # Try with getAttribute
                try {
                    if ($table.getAttribute("class") -eq "MAIN") {
                        $mainTable = $table
                        Write-Log "Found main table using getAttribute method."
                        break
                    }
                } catch {
                    # Ignore errors with getAttribute
                }
                
                # Check if it's a wide table with borders and multiple columns
                try {
                    if ($table.border -eq "1" -and $table.summary -match "stock") {
                        $mainTable = $table
                        Write-Log "Found main table by border and summary attributes."
                        break
                    }
                } catch {
                    # Ignore errors with border/summary checks
                }
            }
            
            if ($null -eq $mainTable) {
                # Last resort: find a table with at least 6 columns
                foreach ($table in $tables) {
                    try {
                        $headerRow = $table.rows.item(0)
                        if ($headerRow -and $headerRow.cells.length -ge 6) {
                            $mainTable = $table
                            Write-Log "Found table with $($headerRow.cells.length) columns, using as main table."
                            break
                        }
                    } catch {
                        # Ignore errors with checking rows/cells
                    }
                }
            }
            
            if ($null -eq $mainTable) {
                throw "Could not find the main parts table in the HTML content"
            }
            
            # Get all rows from the table
            $rows = $mainTable.getElementsByTagName("tr")
            Write-Log "Number of rows found: $($rows.length)"
            
            # Initialize array to hold parsed data
            $parsedData = @()
            
            # Update the progress bar
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            $progressBar.Maximum = $rows.length
            $progressBar.Value = 0
            
            # Process each row (skip first row which is the header)
            for ($i = 1; $i -lt $rows.length; $i++) {
                $progressBar.Value = $i
                $progressLabel.Text = "Processing row $i of $($rows.length)..."
                $progressForm.Refresh()
                
                try {
                    $row = $rows.item($i)
                    if ($null -eq $row) {
                        Write-Log "Warning: Row $i is null, skipping"
                        continue
                    }
                    
                    # Skip rows that don't have the MAIN class or enough cells
                    $rowClass = try { $row.className } catch { "" }
                    if ($rowClass -ne "MAIN" -and $rowClass -ne "HILITE") {
                        Write-Log "Skipping row $i - not a main data row (class: $rowClass)"
                        continue
                    }
                    
                    # Get all cells in the row
                    $cells = $row.getElementsByTagName("td")
                    
                    if ($null -eq $cells -or $cells.length -lt 6) {
                        Write-Log "Skipping row $i - insufficient cells (found: $(if ($null -eq $cells) { "null" } else { $cells.length }))"
                        continue
                    }
                    
                    # Extract part information from cells safely
                    $partNSN = try { $cells.item(0).innerText.Trim() } catch { "" }
                    $description = try { $cells.item(1).innerText.Trim() } catch { "" }
                    $qtyText = try { $cells.item(2).innerText } catch { "0" }
                    $usageText = try { $cells.item(3).innerText } catch { "0" }
                    $location = try { $cells.item(5).innerText.Trim() } catch { "" }
                    
                    # Use regex to extract digits only
                    $qty = [int]($qtyText -replace '[^\d]', '')
                    $usage = [int]($usageText -replace '[^\d]', '')
                    
                    # Extract OEM information from cell 4 safely
                    $oem1 = ""
                    $oem2 = ""
                    $oem3 = ""
                    
                    try {
                        $oemCell = $cells.item(4)
                        if ($oemCell) {
                            $oemDivs = $oemCell.getElementsByTagName("div")
                            
                            if ($oemDivs -and $oemDivs.length -gt 0) {
                                # Process each div to extract OEM information
                                for ($j = 0; $j -lt $oemDivs.length; $j++) {
                                    try {
                                        $oemDiv = $oemDivs.item($j)
                                        if ($oemDiv) {
                                            $oemText = $oemDiv.innerText.Trim()
                                            
                                            # Extract OEM number from text like "OEM:1 12345"
                                            if ($oemText -match "OEM:1\s+(.+)") {
                                                $oem1 = $matches[1]
                                            } elseif ($oemText -match "OEM:2\s+(.+)") {
                                                $oem2 = $matches[1]
                                            } elseif ($oemText -match "OEM:3\s+(.+)") {
                                                $oem3 = $matches[1]
                                            }
                                        }
                                    } catch {
                                        Write-Log "Warning: Error processing OEM div $j in row $i : $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Log "Warning: Error processing OEM cell in row $i : $($_.Exception.Message)"
                    }
                    
                    # Create object with parsed data
                    $parsedData += [PSCustomObject]@{
                        "Part (NSN)" = $partNSN
                        "Description" = $description
                        "QTY" = $qty
                        "13 Period Usage" = $usage
                        "Location" = $location
                        "OEM 1" = $oem1
                        "OEM 2" = $oem2
                        "OEM 3" = $oem3
                    }
                    
                    Write-Log "Added row ${i}: Part(NSN)=$partNSN, Description=$description, QTY=$qty, Location=$location"
                } catch {
                    Write-Log "Error processing row $i : $($_.Exception.Message)"
                    # Continue with next row instead of stopping
                }
            }
            
            Write-Log "Number of parsed data entries: $($parsedData.Count)"
            
            if ($parsedData.Count -eq 0) {
                throw "No data parsed from HTML content"
            }
            
            # Check if there's an existing CSV file with parts book data that we need to preserve
            $tempCsvPath = Join-Path $config.PartsRoomDirectory "temp_$selectedSiteName.csv"
            $csvFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.csv"
            $existingData = $null
            
            if (Test-Path $csvFilePath) {
                Write-Log "Found existing CSV file. Will merge with new data..."
                $progressLabel.Text = "Merging with existing data..."
                $progressForm.Refresh()
                
                # Load the existing data
                $existingData = Import-Csv -Path $csvFilePath
                
                # Identify column names from book references (any column not in the base set)
                $baseColumns = @("Part (NSN)", "Description", "QTY", "13 Period Usage", "Location", "OEM 1", "OEM 2", "OEM 3", "Changed Part (NSN)")
                $bookColumns = $existingData[0].PSObject.Properties.Name | Where-Object { $baseColumns -notcontains $_ }
                Write-Log "Found book reference columns: $($bookColumns -join ', ')"
                
                # First, export the new data to a temporary file
                $parsedData | Export-Csv -Path $tempCsvPath -NoTypeInformation
                
                # Now read it back to ensure formatting is consistent
                $newData = Import-Csv -Path $tempCsvPath
                
                # Add any missing columns from the existing data to the new data
                foreach ($column in $bookColumns) {
                    if ($newData[0].PSObject.Properties.Name -notcontains $column) {
                        $newData | ForEach-Object { $_ | Add-Member -NotePropertyName $column -NotePropertyValue "" }
                    }
                }
                
                # Add 'Changed Part (NSN)' column if it doesn't exist
                if ($newData[0].PSObject.Properties.Name -notcontains 'Changed Part (NSN)') {
                    $newData | ForEach-Object { $_ | Add-Member -NotePropertyName 'Changed Part (NSN)' -NotePropertyValue "" }
                }
                
                # Update progress for the merge operation
                $progressBar.Value = 0
                $progressBar.Maximum = $newData.Count
                
                # Create a dictionary for faster lookups of existing data
                $existingDict = @{}
                foreach ($item in $existingData) {
                    if (-not [string]::IsNullOrEmpty($item.'Part (NSN)')) {
                        $existingDict[$item.'Part (NSN)'] = $item
                    }
                }
                
                # Counters for stats
                $updatedCount = 0
                $newCount = 0
                
                # Process each item in the new data
                for ($i = 0; $i -lt $newData.Count; $i++) {
                    $progressBar.Value = $i
                    if ($i % 10 -eq 0) {  # Update the label less frequently for performance
                        $progressLabel.Text = "Merging item $i of $($newData.Count)..."
                        $progressForm.Refresh()
                    }
                    
                    $item = $newData[$i]
                    $partNSN = $item.'Part (NSN)'
                    
                    # Skip items with no Part (NSN)
                    if ([string]::IsNullOrEmpty($partNSN)) {
                        continue
                    }
                    
                    # Check if this part exists in the existing data
                    if ($existingDict.ContainsKey($partNSN)) {
                        # Update QTY and Location from new data
                        $existingItem = $existingDict[$partNSN]
                        
                        # Check if QTY or Location has changed
                        if ($existingItem.QTY -ne $item.QTY -or $existingItem.Location -ne $item.Location) {
                            $existingItem.QTY = $item.QTY
                            $existingItem.Location = $item.Location
                            $existingItem.'13 Period Usage' = $item.'13 Period Usage'
                            $updatedCount++
                            Write-Log "Updated part ${partNSN}: QTY=$($item.QTY), Location=$($item.Location)"
                        }
                        
                        # Update OEM information if it's more complete in the new data
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 1') -and -not [string]::IsNullOrEmpty($item.'OEM 1')) {
                            $existingItem.'OEM 1' = $item.'OEM 1'
                        }
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 2') -and -not [string]::IsNullOrEmpty($item.'OEM 2')) {
                            $existingItem.'OEM 2' = $item.'OEM 2'
                        }
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 3') -and -not [string]::IsNullOrEmpty($item.'OEM 3')) {
                            $existingItem.'OEM 3' = $item.'OEM 3'
                        }
                    } else {
                        # This is a new part, add it to the existing data with empty book references
                        foreach ($column in $bookColumns) {
                            if ($item.PSObject.Properties.Name -notcontains $column) {
                                $item | Add-Member -NotePropertyName $column -NotePropertyValue ""
                            }
                        }
                        
                        # Add to the dictionary for future reference
                        $existingDict[$partNSN] = $item
                        
                        # Add to the existing data list
                        $existingData += $item
                        $newCount++
                        Write-Log "Added new part: $partNSN"
                    }
                }
                
                # Now check for parts that exist in the existing data but not in the new data
                # They might have been removed from inventory
                $removedParts = @()
                foreach ($existingItem in $existingData) {
                    $partNSN = $existingItem.'Part (NSN)'
                    if (-not [string]::IsNullOrEmpty($partNSN)) {
                        $found = $false
                        foreach ($newItem in $newData) {
                            if ($newItem.'Part (NSN)' -eq $partNSN) {
                                $found = $true
                                break
                            }
                        }
                        
                        if (-not $found) {
                            # Mark this part as potentially removed from inventory
                            $existingItem.QTY = "0"
                            $existingItem.Location = "Not in current inventory"
                            $removedParts += $partNSN
                        }
                    }
                }
                
                if ($removedParts.Count -gt 0) {
                    Write-Log "Marked $($removedParts.Count) parts as not in current inventory"
                }
                
                # Export the updated data back to the main CSV file
                $progressLabel.Text = "Saving updated data..."
                $progressForm.Refresh()
                
                $existingData | Export-Csv -Path $csvFilePath -NoTypeInformation
                
                # Clean up the temporary file
                if (Test-Path $tempCsvPath) {
                    Remove-Item -Path $tempCsvPath -Force
                }
                
                Write-Log "Merged data successfully: Updated $updatedCount parts, added $newCount new parts, marked $($removedParts.Count) parts as not in current inventory"
                [System.Windows.Forms.MessageBox]::Show("Parts Room data updated successfully:`n- Updated $updatedCount existing parts`n- Added $newCount new parts`n- Marked $($removedParts.Count) parts as not in current inventory", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # No existing file, just save the new data
                $progressLabel.Text = "Saving data to CSV..."
                $progressForm.Refresh()
                
                # Export parsed data to CSV
                $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-Log "Created new CSV file at $csvFilePath with $($parsedData.Count) parts"
                [System.Windows.Forms.MessageBox]::Show("Parts Room data has been created successfully with $($parsedData.Count) parts.", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            
            # Now check if we need to update the Excel file
            $excelFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.xlsx"
            if (Test-Path $excelFilePath) {
                $updateExcel = [System.Windows.Forms.MessageBox]::Show(
                    "Do you want to update the Excel file with the new data?`n`nNote: This will create a new Excel file. Your existing links to figures in parts books will be preserved in the CSV file.",
                    "Update Excel",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($updateExcel -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $progressLabel.Text = "Updating Excel file..."
                    $progressForm.Refresh()
                    Create-ExcelFromCsv -siteName $selectedSiteName `
                      -csvDirectory (Split-Path $csvFilePath -Parent) `
                      -excelDirectory (Split-Path $excelFilePath -Parent)
                }
            } else {
                $createExcel = [System.Windows.Forms.MessageBox]::Show(
                    "Do you want to create an Excel file with the new data?",
                    "Create Excel",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($createExcel -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $progressLabel.Text = "Creating Excel file..."
                    $progressForm.Refresh()
                    Create-ExcelFromCsv -siteName $selectedSiteName `
                      -csvDirectory (Split-Path $csvFilePath -Parent) `
                      -excelDirectory (Split-Path $excelFilePath -Parent)
                }
            }
            
            # Ask if the user wants to update parts books with the latest QTY and Location data
            $updatePartsBooks = [System.Windows.Forms.MessageBox]::Show(
                "Do you want to update Parts Books with the latest quantity and location information?",
                "Update Parts Books",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
                
            if ($updatePartsBooks -eq [System.Windows.Forms.DialogResult]::Yes) {
                $progressLabel.Text = "Updating Parts Books..."
                $progressForm.Refresh()
                Update-PartsBooks -sourceCSVPath $csvFilePath
            }
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            Write-Log "Stack Trace: $($_.ScriptStackTrace)"
            $errorMessage = "Error: $($_.Exception.Message)`r`nStack Trace: $($_.ScriptStackTrace)"
            $errorMessage | Out-File -FilePath $logPath -Append
            [System.Windows.Forms.MessageBox]::Show("An error occurred. Please check the error log at $logPath for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            # Clean up COM objects
            if ($null -ne $htmlDoc) {
                try {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
                } catch {
                    Write-Log "Warning: Failed to release COM object: $($_.Exception.Message)"
                }
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            # Close the progress form
            $progressForm.Close()
        }
    } catch {
        Write-Log "Error downloading HTML content: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to download HTML content: $($_.Exception.Message)", "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to get Excel files (Parts Books)
function Get-ExcelFiles {
    $excelFiles = @()
    if ($config.Books) {
        Write-Log "Processing Books from config:"
        foreach ($book in $config.Books.PSObject.Properties) {
		  $volumesCsvPath = $book.Value.VolumesToUrlCsvPath
		  if ($volumesCsvPath) {
			  $bookDir = Split-Path -Path $volumesCsvPath -Parent
		  } else {
			  $bookDir = Join-Path $config.PartsBooksDirectory $book.Name
		  }
		  if (-not (Test-Path $bookDir)) { Write-Log "Directory not found for: $($book.Name)"; continue }
		  $excelFilePath = Get-ChildItem -Path $bookDir -Filter "*.xlsx" -ErrorAction SilentlyContinue |
						   Select-Object -First 1 -ExpandProperty FullName
            Write-Log "Checking file: $excelFilePath"
            if ($excelFilePath -and (Test-Path $excelFilePath)) {
                $excelFiles += @{
                    Name = $book.Name
                    Path = $excelFilePath
                }
                Write-Log "Added file: $($book.Name)"
            } else {
                Write-Log "File not found for: $($book.Name)"
            }
        }
    }
    return $excelFiles
}

function Add-SameDayPartsRoom {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Add Same Day Parts Room'
    $form.Size = New-Object System.Drawing.Size(500,600)
    $form.StartPosition = 'CenterScreen'

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10,10)
    $listView.Size = New-Object System.Drawing.Size(460,500)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.CheckBoxes = $true
    $listView.Columns.Add("Site ID", 100) | Out-Null
    $listView.Columns.Add("Full Name", 340) | Out-Null
    $form.Controls.Add($listView)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(200,520)
    $addButton.Size = New-Object System.Drawing.Size(100,30)
    $addButton.Text = 'Add Selected'
    $form.Controls.Add($addButton)

    # Load sites from CSV
    $sitesPath = Join-Path -Path $config.DropdownCsvsDirectory -ChildPath "Sites.csv"
    if (Test-Path $sitesPath) {
        $sites = Import-Csv -Path $sitesPath

        # Exclude sites already in SameDayPartsRooms
		if (-not $config.SameDayPartsRooms) {
			$config | Add-Member -NotePropertyName SameDayPartsRooms -NotePropertyValue @() -Force
		}
        $existingSites = $config.SameDayPartsRooms | ForEach-Object { $_.SiteID }
        $sitesToAdd = $sites | Where-Object { $existingSites -notcontains $_.'Site ID' }

        foreach ($site in $sitesToAdd) {
            $item = New-Object System.Windows.Forms.ListViewItem($site.'Site ID')
            $item.SubItems.Add($site.'Full Name')
            $listView.Items.Add($item)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Sites.csv file not found at $sitesPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $addButton.Add_Click({
        $selectedSites = $listView.CheckedItems | ForEach-Object {
            @{
                SiteID = $_.Text
                FullName = $_.SubItems[1].Text
                Email = ""  # Placeholder for future use
            }
        }

        if ($selectedSites.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No sites selected.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        # Update the configuration
        if (-not $config.SameDayPartsRooms) {
            $config | Add-Member -NotePropertyName SameDayPartsRooms -NotePropertyValue @()
        }
        $config.SameDayPartsRooms += $selectedSites

        # Save the updated configuration
        $config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8

        # Create subdirectory for Same Day Parts Room
        $sameDayPartsRoomDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
        if (-not (Test-Path $sameDayPartsRoomDir)) {
            New-Item -Path $sameDayPartsRoomDir -ItemType Directory | Out-Null
            Write-Log "Created directory: $sameDayPartsRoomDir"
        }

        # Process each selected site
        foreach ($site in $selectedSites) {
            $SiteID = $site.SiteID
            $siteName = $site.FullName

            # Get site URL from Sites.csv
            $siteInfo = $sites | Where-Object { $_.'Site ID' -eq $SiteID }
            if ($siteInfo) {
                $siteUrl = "http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$($SiteID)&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

                if ($siteUrl) {
                    # Download HTML
                    Write-Log "Downloading HTML content for site $siteName..."
                    $htmlContent = Invoke-WebRequest -Uri $siteUrl -UseBasicParsing

                    # Save HTML to a file
                    $htmlFilePath = Join-Path $sameDayPartsRoomDir "$siteName.html"
                    $htmlContent.Content | Out-File -FilePath $htmlFilePath -Encoding UTF8
                    Write-Log "Downloaded HTML for site $siteName to $htmlFilePath"

                    # Parse HTML into CSV
                    $parsedData = Parse-HTMLToCSV -htmlFilePath $htmlFilePath -siteName $siteName

                    # Save CSV file named after the site
                    $csvFilePath = Join-Path $sameDayPartsRoomDir "$siteName.csv"
                    $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                    Write-Log "Parsed HTML and saved CSV for site $siteName to $csvFilePath"
                } else {
                    Write-Log "URL not found for site $SiteID"
                }
            } else {
                Write-Log "Site information not found for site $SiteID"
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Selected sites have been added and processed.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })

    $form.ShowDialog()
}

function Add-PartsBookFromCatalog {
    param($config)

    $parsedCsvPath = Join-Path $config.DropdownCsvsDirectory "Parsed-Parts-Volumes.csv"
    if (-not (Test-Path $parsedCsvPath) -or (Get-Item $parsedCsvPath).Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Parsed-Parts-Volumes.csv is missing or empty. Run Handbook-Dropdowns-CSV-Creator first.",
            "Cannot Add Books",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $csvData = Import-Csv $parsedCsvPath
    $alreadyHave = if ($config.Books) { $config.Books.PSObject.Properties.Name } else { @() }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Add Parts Book"
    $form.Size = New-Object System.Drawing.Size(720, 520)
    $form.StartPosition = 'CenterScreen'

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(10, 10)
    $checkedList.Size = New-Object System.Drawing.Size(690, 430)
    $checkedList.CheckOnClick = $true

    foreach ($row in $csvData) {
        $bookKey = ($row.'Full Name' -replace '[^\w\s-]', '' -replace '\s+', ' ').Trim()
        $display = "$($row.'Full Name')   MS$($row.'MS Book No') Vol $($row.Volume)"
        if ($alreadyHave -contains $bookKey) { $display += "   [already installed]" }
        $checkedList.Items.Add($display) | Out-Null
    }
    $form.Controls.Add($checkedList)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Add Selected"
    $ok.Location = New-Object System.Drawing.Point(520, 450)
    $ok.Size = New-Object System.Drawing.Size(180, 30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    if (-not $config.Books) {
        $config | Add-Member -NotePropertyName Books -NotePropertyValue @{} -Force
    }

    # Helper: build a folder name the same way Parts-Books-Creator's Sanitize-Name does,
    # so the config path and the on-disk folder agree.
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars()

    $added = 0
    foreach ($idx in $checkedList.CheckedIndices) {
        $row = $csvData[$idx]
        $bookKey = ($row.'Full Name' -replace '[^\w\s-]', '' -replace '\s+', ' ').Trim()
        if ($alreadyHave -contains $bookKey) { continue }

        $folderName = $row.'Full Name'
        foreach ($c in $invalidChars) {
            $folderName = $folderName -replace [regex]::Escape($c), '-'
        }
        $bookDir = Join-Path $config.PartsBooksDirectory $folderName
        New-Item -ItemType Directory -Force -Path $bookDir | Out-Null

        $config.Books | Add-Member -NotePropertyName $bookKey -NotePropertyValue @{
            VolumesToUrlCsvPath = Join-Path $bookDir "Volumes-to-URL.csv"
            SectionNamesCsvPath = Join-Path $bookDir "SectionNames.txt"
        } -Force
        $added++
    }

    $config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
    Write-Log "Added $added book(s) to config."

    [System.Windows.Forms.MessageBox]::Show(
        "Added $added book(s). Use 'Create Parts Book' to download and index their contents.",
        "Books Added",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Remove-PartsBook {
    param($config)

    $books = if ($config.Books) { @($config.Books.PSObject.Properties.Name) } else { @() }
    if ($books.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No books configured.", "Nothing to Remove")
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Remove Parts Book"
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = 'CenterScreen'

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Location = New-Object System.Drawing.Point(10, 10)
    $list.Size = New-Object System.Drawing.Size(470, 240)
    foreach ($b in $books) { $list.Items.Add($b) | Out-Null }
    $form.Controls.Add($list)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Remove Selected from Config"
    $ok.Location = New-Object System.Drawing.Point(270, 260)
    $ok.Size = New-Object System.Drawing.Size(210, 30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $removed = 0
    foreach ($idx in $list.CheckedIndices) {
        $name = $list.Items[$idx]
        $config.Books.PSObject.Properties.Remove($name)
        $removed++
    }

    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $script:configPath
    Write-Log "Removed $removed book(s) from config. Files on disk preserved."

    [System.Windows.Forms.MessageBox]::Show(
        "Removed $removed book(s) from config. The downloaded files were left on disk.",
        "Books Removed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Remove-SameDayPartsRoom {
    param($config)

    if (-not $config.SameDayPartsRooms -or $config.SameDayPartsRooms.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No Same Day sites configured.", "Nothing to Remove")
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Remove Same Day Parts Room Site"
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = 'CenterScreen'

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Location = New-Object System.Drawing.Point(10, 10)
    $list.Size = New-Object System.Drawing.Size(470, 240)
    foreach ($s in $config.SameDayPartsRooms) {
        $list.Items.Add("$($s.SiteID)   $($s.FullName)") | Out-Null
    }
    $form.Controls.Add($list)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Remove Selected"
    $ok.Location = New-Object System.Drawing.Point(270, 260)
    $ok.Size = New-Object System.Drawing.Size(210, 30)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $keep = @()
    $remove = @{}
    foreach ($idx in $list.CheckedIndices) {
        $remove[$config.SameDayPartsRooms[$idx].SiteID] = $true
    }
    foreach ($s in $config.SameDayPartsRooms) {
        if (-not $remove.ContainsKey($s.SiteID)) { $keep += $s }
    }
    $config.SameDayPartsRooms = $keep

    $config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
    Write-Log "Removed $($remove.Count) Same Day site(s) from config. Files preserved."

    [System.Windows.Forms.MessageBox]::Show(
        "Removed $($remove.Count) site(s) from config. The cached CSVs were left on disk.",
        "Sites Removed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

################################################################################
#                           Needs Category                                     #
################################################################################

# Ensure RootDirectory is set
if (-not $config.RootDirectory) {
    $config.RootDirectory = $PSScriptRoot
}

# Ensure all required paths are set
$requiredPaths = @('RootDirectory', 'PartsRoomDirectory', 'DropdownCsvsDirectory', 'PartsBooksDirectory')
foreach ($path in $requiredPaths) {
    if (-not $config.$path) {
        Write-Log "Error: $path is not set in the configuration"
        throw "$path is missing from the configuration"
    }
}

$script:unacknowledgedEntries = @{}



# Define and set default paths if not specified in config
$defaultPaths = @{
    PartsRoomDirectory = "PartsRoom"
    PartsBooksDirectory = "PartsBooks"
    DropdownCsvsDirectory = "DropdownCsvs"
}

$configUpdated = $false

foreach ($key in $defaultPaths.Keys) {
    if (-not $config.$key) {
        $config | Add-Member -NotePropertyName $key -NotePropertyValue (Join-Path $config.RootDirectory $defaultPaths[$key])
        $configUpdated = $true
    }
}

# Check and create required directories
$requiredDirs = @($config.PartsRoomDirectory, $config.PartsBooksDirectory, $config.DropdownCsvsDirectory)

# Save updated config if changes were made
if ($configUpdated) {
    $config | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
    Write-Log "Config file updated with default paths"
}

################################################################################
#                           Main Execution                                     #
################################################################################

# Function to create and show the main form
function Show-MainForm {
    if (-not (Test-Path $script:configPath)) {
        $suggested = Join-Path $PSScriptRoot 'PartsMgmt'
        $ok = Show-InstallationWizard -InitialRoot $suggested
        if (-not $ok) { return }
        $script:config = Get-Content -Path $script:configPath -Raw | ConvertFrom-Json
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Parts Management System'
    $form.Size = New-Object System.Drawing.Size(1280, 900)
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::FromArgb(236,240,241)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = 'Fill'
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $tabControl.Padding = New-Object System.Drawing.Point(16, 6)
    $form.Controls.Add($tabControl)

    # ---- Parts Books Tab ----
    $partsBookTab = New-Object System.Windows.Forms.TabPage
    $partsBookTab.Text = "Parts Books"
    $partsBookTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $partsBookTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($partsBookTab)

    $partsBookPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $partsBookPanel.Dock = 'Fill'
    $partsBookPanel.FlowDirection = 'TopDown'
    $partsBookPanel.WrapContents = $false
    $partsBookPanel.AutoScroll = $true
    $partsBookPanel.BackColor = [System.Drawing.Color]::Transparent
    $partsBookTab.Controls.Add($partsBookPanel)

    # ---- Work Tracking Tab ----
    $workTrackingTab = New-Object System.Windows.Forms.TabPage
    $workTrackingTab.Text = "Work Tracking"
    $tabControl.TabPages.Add($workTrackingTab)
    Setup-WorkTrackingTab -parentTab $workTrackingTab

    # ---- Open Parts Room ----
	$openPartsRoomButton = New-Button "Open Parts Room" {
		$partsRoomFilePath = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
							 Select-Object -First 1 -ExpandProperty FullName

		if ($partsRoomFilePath -and (Test-Path $partsRoomFilePath)) {
			Start-Process $partsRoomFilePath
			return
		}

		$csvPath = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.csv" -File -ErrorAction SilentlyContinue |
				   Select-Object -First 1 -ExpandProperty FullName

		if ($csvPath) {
			$ans = [System.Windows.Forms.MessageBox]::Show(
				"No Parts Room Excel file found.`r`n`r`nOpen the CSV instead?`r`n$csvPath",
				"Parts Room",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question)
			if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
				Start-Process $csvPath
			}
		} else {
			[System.Windows.Forms.MessageBox]::Show(
				"No Parts Room data found in:`r`n$($config.PartsRoomDirectory)`r`n`r`nRun 'Update Parts Room' from the Actions tab first.",
				"Parts Room",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information)
		}
	} -Style 'Primary'
    $partsBookPanel.Controls.Add($openPartsRoomButton)

    # ---- Create Parts Book ----
    $createPartsBookButton = New-Button "Create Parts Book" {
        Show-PartsBookCreatorDialog
        $form.Refresh()
    } -Style 'Primary'
    $partsBookPanel.Controls.Add($createPartsBookButton)

    # ---- Excel parts books ----
    $excelFiles = Get-ExcelFiles
    foreach ($file in $excelFiles) {
        $button = New-Button $file.Name -Action {
            $filePath = $this.Tag
            Write-Log "Button clicked for file: $filePath"
            if (Test-Path $filePath) {
                Start-Process $filePath
            } else {
                [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } -Tag $file.Path -Style 'Ghost'
        $partsBookPanel.Controls.Add($button)
    }

    # ---- Actions Tab ----
    $actionsTab = New-Object System.Windows.Forms.TabPage
    $actionsTab.Text = "Actions"
    $actionsTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $actionsTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($actionsTab)

    $actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionsPanel.Dock = 'Fill'
    $actionsPanel.FlowDirection = 'TopDown'
    $actionsPanel.WrapContents = $false
    $actionsPanel.AutoScroll = $true
    $actionsPanel.BackColor = [System.Drawing.Color]::Transparent
    $actionsTab.Controls.Add($actionsPanel)

    # Search tab needs to exist before "Search for a Part" action can switch to it
    $searchTab = New-Object System.Windows.Forms.TabPage
    $searchTab.Text = "Search"
    $searchTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $searchTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($searchTab)
    Setup-SearchTab -parentTab $searchTab -config $config

    # ---- Settings Tab ----
    $settingsTab = New-Object System.Windows.Forms.TabPage
    $settingsTab.Text = "Settings"
    $settingsTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $settingsTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($settingsTab)
    Setup-SettingsTab -parentTab $settingsTab

    # Style per action
    $actionButtons = @(
        @{Text="Update Files";                   Action={ Update-AllFiles };                                      Style='Success' }
        @{Text="Update Parts Books";             Action={ Update-PartsBooks };                                    Style='Primary' }
        @{Text="Update Parts Room";              Action={ Update-PartsRoom };                                     Style='Primary' }
        @{Text="Take a Part Out";                Action={ Take-PartOut };                                         Style='Primary' }
        @{Text="Search for a Part";              Action={ $tabControl.SelectedTab = $searchTab };                 Style='Primary' }
        @{Text="Request a Part to be Ordered";   Action={ Request-PartOrder };                                    Style='Secondary' }
        @{Text="Request a Work Order";           Action={ Request-WorkOrder };                                    Style='Secondary' }
        @{Text="Make an MTSC Ticket";            Action={ Make-MTSCTicket };                                      Style='Secondary' }
        @{Text="Search Knowledge Base";          Action={ Search-KnowledgeBase };                                 Style='Secondary' }
        @{Text="Add Parts Book";                 Action={ Add-PartsBookFromCatalog -config $config };             Style='Success' }
        @{Text="Remove Parts Book";              Action={ Remove-PartsBook -config $config };                     Style='Danger' }
        @{Text="Add Same Day Parts Room";        Action={ Add-SameDayPartsRoom };                                 Style='Success' }
        @{Text="Remove Same Day Parts Room";     Action={ Remove-SameDayPartsRoom -config $config };              Style='Danger' }
        @{Text="Add 1-Day Parts Room";           Action={ [System.Windows.Forms.MessageBox]::Show("Not yet implemented.") }; Style='Ghost' }
        @{Text="Add 2-Day Parts Room";           Action={ [System.Windows.Forms.MessageBox]::Show("Not yet implemented.") }; Style='Ghost' }
    )

    foreach ($actionButton in $actionButtons) {
        $button = New-Button $actionButton.Text $actionButton.Action -Style $actionButton.Style
        $actionsPanel.Controls.Add($button)
    }

    Write-Log "UI setup completed"
    $form.ShowDialog()
}


# Main execution
Show-MainForm
Write-Log "Application closed"
