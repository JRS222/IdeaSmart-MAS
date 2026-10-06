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

<# 
# Configuration variables
$global:config = $null                    # Configuration object
$global:configPath = $null                # Path to configuration file

# File paths
$script:laborLogsFilePath = $null         # Path to labor logs file
$script:callLogsFilePath = $null          # Path to call logs file
$script:logPath = "UI.log"                # Path to UI log file

# State tracking
$script:unacknowledgedEntries = @{}       # Unacknowledged entry tracking
$script:processedCallLogs = @{}           # Processed call logs tracking
$script:workOrderParts = @{}              # Parts for work orders
$script:machinesUpdated = $false          # Flag for machines list update

# Labor and Call Logs UI elements
$script:listViewLaborLog = New-Object System.Windows.Forms.ListView  # Labor logs
$script:listViewCallLogs = $null          # Call logs list view
$script:notificationIcon = $null          # Notification icon

# Search UI elements
$script:textBoxNSN = $null                # NSN search text box
$script:textBoxOEM = $null                # OEM search text box
$script:textBoxDescription = $null        # Description search text box
$script:listViewAvailability = $null      # Availability results
$script:listViewSameDayAvailability = $null  # Same-day availability
$script:listViewCrossRef = $null          # Cross-reference results
$script:openFiguresButton = $null         # Open figures button
$script:takePartOutButton = $null         # Take part out button

# Tab controls
$script:tabControl = $null                # Main tab control
$script:partsBookTab = $null              # Parts book tab
$script:callLogsTab = $null               # Call logs tab
$script:laborLogTab = $null               # Labor log tab
$script:searchTab = $null                 # Search tab
$script:actionsTab = $null                # Actions tab 
$script:configPath = $null
#>

# Make sure this dictionary always exists before anything tries to read it
if ($null -eq $script:workOrderParts) { $script:workOrderParts = @{} }
if ($null -eq $script:unacknowledgedEntries) { $script:unacknowledgedEntries = @{} }
if ($null -eq $script:processedCallLogs) { $script:processedCallLogs = @{} }
if ($null -eq $script:pmSuppressEvents)      { $script:pmSuppressEvents = $false }
if ($null -eq $script:pmTotalSelectedHours)  { $script:pmTotalSelectedHours = 0.0 }

################################################################################
#                            Core Utilities                                    #
################################################################################

$script:configPath = Join-Path $PSScriptRoot "Config.json"

#UI Log
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp - $message"
    # Optionally, you can also write to a log file:
    "$timestamp - $message" | Out-File -Append -FilePath "UI.log"
}

#Initialize the config
function Initialize-Config {
    if (Test-Path $configPath) {
        $config = Get-Content -Path $configPath | ConvertFrom-Json
        if (-not ($config.PSObject.Properties.Name -contains 'SameDayPartsRooms')) {
            $config | Add-Member -NotePropertyName 'SameDayPartsRooms' -NotePropertyValue @()
            Write-Log "Backfilled missing SameDayPartsRooms key."
        }
        if (-not ($config.PSObject.Properties.Name -contains 'Books')) {
            $config | Add-Member -NotePropertyName 'Books' -NotePropertyValue @{}
            Write-Log "Backfilled missing Books key."
        }
        Write-Log "Config loaded successfully"
    } else {
        Write-Log "Config file not found. Using default configuration."
        $config = @{
            RootDirectory = $PSScriptRoot
            LaborDirectory = Join-Path $PSScriptRoot "Labor"
            CallLogsDirectory = Join-Path $PSScriptRoot "Call Logs"
            PartsRoomDirectory = Join-Path $PSScriptRoot "Parts Room"
            DropdownCsvsDirectory = Join-Path $PSScriptRoot "Dropdown CSVs"
            PartsBooksDirectory = Join-Path $PSScriptRoot "Parts Books"
        }
    }
    return $config
}

$config = Initialize-Config
$laborLogsFilePath = $config.PrerequisiteFiles.LaborLogs
if (-not $laborLogsFilePath) {
    $laborLogsFilePath = Join-Path $config.LaborDirectory "LaborLogs.csv"
    Write-Log "Labor logs file path was not in config, set to: $laborLogsFilePath"
}

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
	
    try {
        Write-Log "Starting to create Excel file from CSV for $siteName..."
        
        # Add timing diagnostics
        $startTime = Get-Date
        
        # Show progress form if not already visible
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Creating Parts Room Excel"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = 'CenterScreen'
        
		$progressBar = New-Object ModernProgressBar
		$progressBar.Size = New-Object System.Drawing.Size(360,24)
		$progressBar.Location = New-Object System.Drawing.Point(10,12)
		$progressForm.Controls.Add($progressBar)
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Size = New-Object System.Drawing.Size(360,40)
        $progressLabel.Location = New-Object System.Drawing.Point(10,40)
        $progressForm.Controls.Add($progressLabel)
        
        $timeLabel = New-Object System.Windows.Forms.Label
        $timeLabel.Size = New-Object System.Drawing.Size(360,20)
        $timeLabel.Location = New-Object System.Drawing.Point(10,90)
        $progressForm.Controls.Add($timeLabel)
        
        $progressForm.Show()
        $progressForm.Refresh()
        
        $csvFilePath = Join-Path $csvDirectory "$siteName.csv"
        $excelFilePath = Join-Path $excelDirectory "$siteName.xlsx"
        
        $progressLabel.Text = "Loading CSV file..."
        $progressBar.Value = 5
        $progressForm.Refresh()
        
        if (-not (Test-Path $csvFilePath)) {
            throw "CSV file not found at $csvFilePath"
        }

        # Read CSV data
        $loadStart = Get-Date
        $csvData = Import-Csv -Path $csvFilePath
        $loadEnd = Get-Date
        $loadDuration = ($loadEnd - $loadStart).TotalSeconds
        Write-Log "CSV loading completed in $loadDuration seconds"
        $timeLabel.Text = "CSV loaded in $loadDuration seconds"
        $progressForm.Refresh()
        
        $progressLabel.Text = "Creating Excel application..."
        $progressBar.Value = 10
        $progressForm.Refresh()
        
        # Create Excel
        $excelStart = Get-Date
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        if (Test-Path $excelFilePath) {
            $workbook = $excel.Workbooks.Open($excelFilePath)
        } else {
            $workbook = $excel.Workbooks.Add()
        }
        
        $excelEnd = Get-Date
        $excelDuration = ($excelEnd - $excelStart).TotalSeconds
        Write-Log "Excel application created in $excelDuration seconds"
        $timeLabel.Text = "Excel app created in $excelDuration seconds"
        
        $progressLabel.Text = "Setting up worksheet..."
        $progressBar.Value = 15
        $progressForm.Refresh()

        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Parts Data"

        # Clear existing content
        $worksheet.Cells.Clear()
        
        $progressLabel.Text = "Adding headers..."
        $progressBar.Value = 20
        $progressForm.Refresh()

        # Add headers first
        $headers = $csvData[0].PSObject.Properties.Name
        for ($col = 1; $col -le $headers.Count; $col++) {
            $worksheet.Cells.Item(1, $col).Value2 = $headers[$col-1]
        }
        
        $progressLabel.Text = "Preparing to add data rows..."
        $progressBar.Value = 25
        $progressForm.Refresh()
        
        # Optimize by using array assignment for data
        $rowCount = $csvData.Count
        $colCount = $headers.Count
        
        # Create a 2D array to hold all data
        $dataArray = New-Object 'object[,]' $rowCount, $colCount
        
        $progressLabel.Text = "Filling data array..."
        $progressBar.Value = 30
        $progressForm.Refresh()
        
        # Fill the array with data
        $arrayStart = Get-Date
        for ($rowIdx = 0; $rowIdx -lt $rowCount; $rowIdx++) {
            $dataRow = $csvData[$rowIdx]
            for ($colIdx = 0; $colIdx -lt $colCount; $colIdx++) {
                $header = $headers[$colIdx]
                $dataArray[$rowIdx, $colIdx] = $dataRow.$header
            }
            
            # Update progress every 100 rows
            if ($rowIdx % 100 -eq 0 -or $rowIdx -eq $rowCount - 1) {
                $percent = 30 + ($rowIdx / $rowCount * 20)  # Scale from 30% to 50%
                $progressBar.Value = [int]$percent
                $progressLabel.Text = "Filling data array: row $($rowIdx+1) of $rowCount"
                $progressForm.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        $arrayEnd = Get-Date
        $arrayDuration = ($arrayEnd - $arrayStart).TotalSeconds
        Write-Log "Data array filled in $arrayDuration seconds"
        $timeLabel.Text = "Array filled in $arrayDuration seconds"
        
        $progressLabel.Text = "Writing data to Excel..."
        $progressBar.Value = 50
        $progressForm.Refresh()
        
        # Get the range to fill (offset by 1 for header row)
        $startRange = $worksheet.Cells.Item(2, 1)
        $endRange = $worksheet.Cells.Item($rowCount + 1, $colCount)
        $dataRange = $worksheet.Range($startRange, $endRange)
        
        # Fill the range in one operation
        $rangeStart = Get-Date
        $dataRange.Value2 = $dataArray
        $rangeEnd = Get-Date
        $rangeDuration = ($rangeEnd - $rangeStart).TotalSeconds
        Write-Log "Excel range filled in $rangeDuration seconds"
        $timeLabel.Text = "Excel range filled in $rangeDuration seconds"
        
        $progressLabel.Text = "Formatting table..."
        $progressBar.Value = 70
        $progressForm.Refresh()

        # Format as table
        $formatStart = Get-Date
        $usedRange = $worksheet.UsedRange
        if ($worksheet.ListObjects.Count -gt 0) {
            $worksheet.ListObjects.Item(1).Unlist()
        }
        $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $usedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
        $listObject.Name = $tableName
        $listObject.TableStyle = "TableStyleMedium2"
        $formatEnd = Get-Date
        $formatDuration = ($formatEnd - $formatStart).TotalSeconds
        Write-Log "Table formatting completed in $formatDuration seconds"
        $timeLabel.Text = "Table formatted in $formatDuration seconds"
        
        $progressLabel.Text = "Applying cell formatting..."
        $progressBar.Value = 80
        $progressForm.Refresh()

        # Apply formatting
        $cellFormatStart = Get-Date
        $usedRange.Cells.VerticalAlignment = -4108 # xlCenter
        $usedRange.Cells.HorizontalAlignment = -4108 # xlCenter
        $usedRange.Cells.WrapText = $false
        $usedRange.Cells.Font.Name = "Courier New"
        $usedRange.Cells.Font.Size = 12
        $cellFormatEnd = Get-Date
        $cellFormatDuration = ($cellFormatEnd - $cellFormatStart).TotalSeconds
        Write-Log "Cell formatting completed in $cellFormatDuration seconds"
        $timeLabel.Text = "Cell formatting in $cellFormatDuration seconds"
        
        $progressLabel.Text = "Auto-fitting columns..."
        $progressBar.Value = 90
        $progressForm.Refresh()

        # AutoFit columns
        $autoFitStart = Get-Date
        $usedRange.Columns.AutoFit() | Out-Null
        $autoFitEnd = Get-Date
        $autoFitDuration = ($autoFitEnd - $autoFitStart).TotalSeconds
        Write-Log "Column auto-fit completed in $autoFitDuration seconds"
        $timeLabel.Text = "Columns auto-fit in $autoFitDuration seconds"

        # Left-align the Description column if it exists
        $descriptionColumn = $listObject.ListColumns | Where-Object { $_.Name -eq "Description" }
        if ($descriptionColumn) {
            $descriptionColumn.Range.Offset(1, 0).HorizontalAlignment = -4131 # xlLeft
        }

        # Remove any columns named "Importing data..."
        for ($col = $headers.Count; $col -ge 1; $col--) {
            $columnHeader = $worksheet.Cells.Item(1, $col).Value2
            if ($columnHeader -eq "Importing data...") {
                $column = $worksheet.Columns.Item($col)
                $column.Delete()
                Write-Log "Removed 'Importing data...' column"
            }
        }
        
        $progressLabel.Text = "Saving Excel file..."
        $progressBar.Value = 95
        $progressForm.Refresh()

        # Save and close
        $saveStart = Get-Date
        $workbook.SaveAs($excelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
        $workbook.Close($false)
        $excel.Quit()
        $saveEnd = Get-Date
        $saveDuration = ($saveEnd - $saveStart).TotalSeconds
        Write-Log "Excel file saved in $saveDuration seconds"
        
        $endTime = Get-Date
        $totalDuration = ($endTime - $startTime).TotalSeconds
        Write-Log "Excel file created successfully at $excelFilePath in total time: $totalDuration seconds"
        
        $progressBar.Value = 100
        $progressLabel.Text = "Excel file created successfully!"
        $timeLabel.Text = "Total time: $totalDuration seconds"
        $progressForm.Refresh()
        Start-Sleep -Seconds 2  # Show completion for 2 seconds
        $progressForm.Close()
    }
    catch {
        Write-Log "Error during Excel file creation: $($_.Exception.Message)"
        if ($progressForm -and $progressForm.Visible) {
            $progressLabel.Text = "Error: $($_.Exception.Message)"
            $progressForm.Refresh()
            Start-Sleep -Seconds 3  # Show error for 3 seconds
            $progressForm.Close()
        }
    }
    finally {
        if ($null -ne $excel) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
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
#                     Labor and Call Logs Management                           #
################################################################################

# Save Log Entries
function Save-Logs {
    param (
        [System.Windows.Forms.ListView]$listView,
        [string]$filePath
    )
    
    try {
        $logs = @()
        foreach ($item in $listView.Items) {
            $log = [PSCustomObject]@{
                Date = $item.SubItems[0].Text
                Machine = $item.SubItems[1].Text
                Cause = $item.SubItems[2].Text
                Action = $item.SubItems[3].Text
                Noun = $item.SubItems[4].Text
                'Time Down' = $item.SubItems[5].Text
                'Time Up' = $item.SubItems[6].Text
                Notes = $item.SubItems[7].Text
            }
            $logs += $log
            Write-Log "Saving log entry: Date=$($log.Date), Machine=$($log.Machine), Time Down=$($log.'Time Down'), Time Up=$($log.'Time Up')"
        }
        
        $logs | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log "Call logs saved to $filePath"
        
        # Verify the saved content
        $savedContent = Get-Content -Path $filePath -Raw
        Write-Log "Saved CSV content: $savedContent"
    }
    catch {
        Write-Log "Error saving call logs: $_"
    }
}

function Save-LaborLogs {
    param($listView, $filePath)
    
    try {
        $logs = @()
        foreach ($item in $listView.Items) {
            $workOrderNumber = $item.SubItems[1].Text
            
            $log = [PSCustomObject]@{
                Date = $item.SubItems[0].Text
                'Work Order' = $workOrderNumber
                'Description' = $item.SubItems[2].Text
                Machine = $item.SubItems[3].Text
                Duration = $item.SubItems[4].Text
                Notes = $item.SubItems[5].Text
                Parts = if ($script:workOrderParts.ContainsKey($workOrderNumber)) {
                          $script:workOrderParts[$workOrderNumber] | ConvertTo-Json -Compress
                        } else {
                          ""
                        }
            }
            $logs += $log
        }
        
        $logs | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log "Labor logs saved to $filePath with $($logs.Count) entries"
    }
    catch {
        Write-Log "Error saving labor logs: $_"
    }
}

# Function to load logs from a CSV file
function Load-Logs {
    param (
        [System.Windows.Forms.ListView]$listView,
        [string]$filePath
    )
    if (Test-Path $filePath) {
        $logs = Import-Csv -Path $filePath
        foreach ($log in $logs) {
            $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
            $item.SubItems.Add($log.Machine)
            $item.SubItems.Add($log.Cause)
            $item.SubItems.Add($log.Action)
            $item.SubItems.Add($log.Noun)
            $item.SubItems.Add($log.'Time Down')
            $item.SubItems.Add($log.'Time Up')
            $item.SubItems.Add($log.Notes)
            $listView.Items.Add($item)
        }
        Write-Log "Logs loaded from $filePath"
    } else {
        Write-Log "Log file not found: $filePath"
    }
}

# Enhanced Load-LaborLogs function with debugging
function Load-LaborLogs {
    param(
        [System.Windows.Forms.ListView]$listView,
        [string]$filePath
    )

    Write-Log "=== START Load-LaborLogs ==="
    Write-Log "Loading labor logs from: $filePath"

    if (-not (Test-Path $filePath)) {
        Write-Log "ERROR: Labor logs file not found at: $filePath"
        return
    }

    try {
        $laborLogs = Import-Csv -Path $filePath
        Write-Log "Successfully loaded labor logs. Entry count: $($laborLogs.Count)"

        if ($null -eq $script:workOrderParts) {
            Write-Log "Initializing workOrderParts dictionary"
            $script:workOrderParts = @{}
        }

        $listView.Items.Clear()

        foreach ($log in $laborLogs) {
            $workOrderNumber = $log.'Work Order'
            Write-Log "Processing work order: $workOrderNumber"

            # Column order: 0 Date, 1 W/O, 2 Description, 3 Machine, 4 Duration, 5 Parts, 6 Notes
            $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
            $item.SubItems.Add($workOrderNumber)  | Out-Null   # 1
            $item.SubItems.Add($log.Description)  | Out-Null   # 2
            $item.SubItems.Add($log.Machine)      | Out-Null   # 3
            $item.SubItems.Add($log.Duration)     | Out-Null   # 4

            # --- Parts (index 5) ---
            if ($log.PSObject.Properties.Name -contains 'Parts' -and
                -not [string]::IsNullOrWhiteSpace($log.Parts)) {
                Write-Log "Work order has parts data: $($log.Parts)"
                try {
                    $parts = $log.Parts | ConvertFrom-Json
                    Write-Log "Parsed JSON parts data. Part count: $($parts.Count)"
                    $script:workOrderParts[$workOrderNumber] = $parts

                    $partsDisplay = ($parts | ForEach-Object {
                        "$($_.PartNumber) - $($_.PartNo) - Qty:$($_.Quantity)"
                    }) -join ", "
                    $item.SubItems.Add($partsDisplay) | Out-Null
                } catch {
                    Write-Log "ERROR parsing Parts JSON for work order ${workOrderNumber}: $($_.Exception.Message)"
                    $item.SubItems.Add("Invalid Parts Data") | Out-Null
                }
            } else {
                Write-Log "Work order has no parts data"
                $item.SubItems.Add("") | Out-Null
            }

            # --- Notes (index 6) ---
            $item.SubItems.Add($log.Notes) | Out-Null

            $listView.Items.Add($item) | Out-Null
            Write-Log "Added item to list view for work order: $workOrderNumber"
        }

        Write-Log "Finished loading labor logs. List view now has $($listView.Items.Count) items"
        Write-Log "=== END Load-LaborLogs ==="
    }
    catch {
        Write-Log "ERROR in Load-LaborLogs: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show(
            "Error loading labor logs: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# A hashtable to store processed call logs to avoid duplication
$script:processedCallLogs = @{}
function Process-HistoricalLogs {
    Write-Log "Processing historical Call Logs to create Labor Logs if necessary..."
   
    if ($null -eq $script:listViewLaborLog) {
        Write-Log "Error: Labor Log ListView is not initialized. Cannot process historical logs."
        return
    }

    $script:listViewLaborLog.Items.Clear()  # Clear the ListView
    $script:processedCallLogs = @{}  # Clear the dictionary
    $callLogs = Import-Csv -Path $callLogsFilePath

    foreach ($log in $callLogs) {
        $startTime = $log.'Time Down'
        $endTime = $log.'Time Up'
        $logKey = "$($log.Date)_$($log.Machine)_$startTime"
       
        if ([string]::IsNullOrWhiteSpace($log.Machine)) {
            Write-Log "Warning: Missing machine information for log entry on $($log.Date)"
            continue
        }
        Write-Log "Processing log: Date=$($log.Date), Machine=$($log.Machine), Time Down=$startTime, Time Up=$endTime"
       
        $timeDiff = Get-TimeDifference -startTime $startTime -endTime $endTime
        Write-Log "Calculated time difference: $timeDiff minutes"
       
        if ($timeDiff -gt 30) {
            Write-Log "Time difference exceeds 30 minutes. Adding to Labor Log."
            Add-LaborLogEntryFromCallLog -log $log
            $script:processedCallLogs[$logKey] = $true
        } else {
            Write-Log "Time difference does not exceed 30 minutes. Skipping."
        }
    }
    # Save the labor logs after processing all call logs
    Save-LaborLogs -listView $script:listViewLaborLog -filePath $global:laborLogsFilePath
}

# Moving Calls from Calls to Labor Log
function Add-LaborLogEntryFromCallLog {
    param ($log)
    try {
        $duration = Get-TimeDifference -startTime $log.'Time Down' -endTime $log.'Time Up'
        $durationHours = [Math]::Round($duration / 60, 2)

        $workOrderNumber = "Need W/O #-$(New-Guid)"
        $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
        $item.SubItems.Add($workOrderNumber)                                                  | Out-Null
        $item.SubItems.Add("$($log.Cause) / $($log.Action) / $($log.Noun)")                   | Out-Null
        $item.SubItems.Add($log.Machine)                                                      | Out-Null
        $item.SubItems.Add($durationHours.ToString("F2"))                                     | Out-Null
        $item.SubItems.Add("")                                                                | Out-Null
        $item.SubItems.Add($log.Notes)                                                        | Out-Null

        $script:listViewLaborLog.Items.Add($item)                                             | Out-Null
        Write-Log "Added Labor Log entry: Date=$($log.Date), Machine=$($log.Machine), Duration=$durationHours hours"

        Save-LaborLogs -listView $script:listViewLaborLog -filePath $global:laborLogsFilePath
    } catch {
        Write-Log "Error adding Labor Log entry: $_"
    }
}

# Function to calculate time difference
function Get-TimeDifference {
    param (
        [string]$startTime,
        [string]$endTime
    )

    try {
        $start = [DateTime]::ParseExact($startTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
        $end = [DateTime]::ParseExact($endTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)

        # Handle cases where end time is on the next day
        if ($end -lt $start) {
            $end = $end.AddDays(1)
        }

        $diff = $end - $start
        return [math]::Round($diff.TotalMinutes)
    }
    catch {
        Write-Log "Error parsing time: $_"
        return 0
    }
}

# New function to update the notification icon
function Update-NotificationIcon {
    if ($script:notificationIcon -eq $null) {
        Write-Log "Error: Notification icon not initialized"
        return
    }

    if ($script:unacknowledgedEntries.Count -gt 0) {
        $script:notificationIcon.Visible = $true
        $script:notificationIcon.Text = "•$($script:unacknowledgedEntries.Count)"
    } else {
        $script:notificationIcon.Visible = $false
    }
}

################################################################################
#                          Parts Management                                    #
################################################################################

# Function to take a part out
function Take-PartOut {
    Write-Log "Taking a part out..."
    [System.Windows.Forms.MessageBox]::Show("Take Part Out process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Add Parts to Work order
function Add-PartsToWorkOrder {
    param($workOrderNumber)
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Add Parts to Work Order #$workOrderNumber"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = 'CenterScreen'
    
    # Search panel (top)
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $searchPanel.Size = New-Object System.Drawing.Size(770, 80)
    $form.Controls.Add($searchPanel)
    
    # NSN search
    $labelNSN = New-Object System.Windows.Forms.Label
    $labelNSN.Text = "Part Number/NSN:"
    $labelNSN.Location = New-Object System.Drawing.Point(10, 15)
    $labelNSN.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelNSN)
    
    $textBoxNSN = New-Object System.Windows.Forms.TextBox
    $textBoxNSN.Location = New-Object System.Drawing.Point(110, 12)
    $textBoxNSN.Size = New-Object System.Drawing.Size(150, 20)
    $searchPanel.Controls.Add($textBoxNSN)
    
    # Description search
    $labelDesc = New-Object System.Windows.Forms.Label
    $labelDesc.Text = "Description:"
    $labelDesc.Location = New-Object System.Drawing.Point(280, 15)
    $labelDesc.Size = New-Object System.Drawing.Size(80, 20)
    $searchPanel.Controls.Add($labelDesc)
    
    $textBoxDesc = New-Object System.Windows.Forms.TextBox
    $textBoxDesc.Location = New-Object System.Drawing.Point(360, 12)
    $textBoxDesc.Size = New-Object System.Drawing.Size(250, 20)
    $searchPanel.Controls.Add($textBoxDesc)
    
    # Search button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(630, 10)
    $searchButton.Size = New-Object System.Drawing.Size(120, 25)
    $searchPanel.Controls.Add($searchButton)
    
    # Results panel (middle)
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Location = New-Object System.Drawing.Point(10, 100)
    $resultsPanel.Size = New-Object System.Drawing.Size(770, 250)
    $form.Controls.Add($resultsPanel)
    
    $resultsLabel = New-Object System.Windows.Forms.Label
    $resultsLabel.Text = "Search Results:"
    $resultsLabel.Location = New-Object System.Drawing.Point(10, 5)
    $resultsLabel.Size = New-Object System.Drawing.Size(100, 20)
    $resultsPanel.Controls.Add($resultsLabel)
    
    $resultsListView = New-Object System.Windows.Forms.ListView
    $resultsListView.Location = New-Object System.Drawing.Point(10, 25)
    $resultsListView.Size = New-Object System.Drawing.Size(750, 220)
    $resultsListView.View = [System.Windows.Forms.View]::Details
    $resultsListView.FullRowSelect = $true
    $resultsListView.CheckBoxes = $true
    $resultsListView.Columns.Add("Part Number", 100)
    $resultsListView.Columns.Add("Description", 250)
    $resultsListView.Columns.Add("QTY Available", 80)
    $resultsListView.Columns.Add("Location", 100)
    $resultsListView.Columns.Add("OEM Number", 100)
    $resultsListView.Columns.Add("Source", 100)
    $resultsPanel.Controls.Add($resultsListView)
    
    # Selected parts panel (bottom)
    $selectedPartsPanel = New-Object System.Windows.Forms.Panel
    $selectedPartsPanel.Location = New-Object System.Drawing.Point(10, 360)
    $selectedPartsPanel.Size = New-Object System.Drawing.Size(770, 150)
    $form.Controls.Add($selectedPartsPanel)
    
    $selectedLabel = New-Object System.Windows.Forms.Label
    $selectedLabel.Text = "Selected Parts:"
    $selectedLabel.Location = New-Object System.Drawing.Point(10, 5)
    $selectedLabel.Size = New-Object System.Drawing.Size(100, 20)
    $selectedPartsPanel.Controls.Add($selectedLabel)
    
    $selectedListView = New-Object System.Windows.Forms.ListView
    $selectedListView.Location = New-Object System.Drawing.Point(10, 25)
    $selectedListView.Size = New-Object System.Drawing.Size(750, 120)
    $selectedListView.View = [System.Windows.Forms.View]::Details
    $selectedListView.FullRowSelect = $true
    $selectedListView.Columns.Add("Part Number", 100)
    $selectedListView.Columns.Add("Description", 250)
    $selectedListView.Columns.Add("Quantity", 80)
    $selectedListView.Columns.Add("Source", 200)
    $selectedListView.Columns.Add("Location", 100)
    $selectedPartsPanel.Controls.Add($selectedListView)
    
    # Buttons panel
    $buttonsPanel = New-Object System.Windows.Forms.Panel
    $buttonsPanel.Location = New-Object System.Drawing.Point(10, 520)
    $buttonsPanel.Size = New-Object System.Drawing.Size(770, 40)
    $form.Controls.Add($buttonsPanel)
    
    $addSelectedButton = New-Object System.Windows.Forms.Button
    $addSelectedButton.Text = "Add Selected Part(s)"
    $addSelectedButton.Location = New-Object System.Drawing.Point(10, 10)
    $addSelectedButton.Size = New-Object System.Drawing.Size(150, 25)
    $buttonsPanel.Controls.Add($addSelectedButton)
    
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove Selected"
    $removeButton.Location = New-Object System.Drawing.Point(170, 10)
    $removeButton.Size = New-Object System.Drawing.Size(150, 25)
    $buttonsPanel.Controls.Add($removeButton)
    
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Parts to Work Order"
    $saveButton.Location = New-Object System.Drawing.Point(610, 10)
    $saveButton.Size = New-Object System.Drawing.Size(150, 25)
    $buttonsPanel.Controls.Add($saveButton)
    
    # Event handlers
    $searchButton.Add_Click({
        $nsn = $textBoxNSN.Text.Trim()
        $desc = $textBoxDesc.Text.Trim()
        
        # Clear previous results
        $resultsListView.Items.Clear()
        
        # Search logic - Simplified for clarity
        $results = Search-Parts -NSN $nsn -Description $desc
        
        # Populate results
        foreach ($part in $results) {
            $item = New-Object System.Windows.Forms.ListViewItem($part.PartNumber)
            $item.SubItems.Add($part.Description) | Out-Null
            $item.SubItems.Add($part.Quantity) | Out-Null
            $item.SubItems.Add($part.Location) | Out-Null
            $item.SubItems.Add($part.OEMNumber) | Out-Null
            $item.SubItems.Add($part.Source) | Out-Null
            $item.Tag = $part  # Store the full part object for later use
            $resultsListView.Items.Add($item) | Out-Null
        }
    })
    
    $addSelectedButton.Add_Click({
        foreach ($item in $resultsListView.CheckedItems) {
            $partObj = $item.Tag
            
            # Prompt for quantity
            $qty = Get-PartQuantity -PartNumber $partObj.PartNumber -MaxQty $partObj.Quantity
            
            if ($qty -gt 0) {
                # Add to selected parts list
                $newItem = New-Object System.Windows.Forms.ListViewItem($partObj.PartNumber)
                $newItem.SubItems.Add($partObj.Description) | Out-Null
                $newItem.SubItems.Add($qty) | Out-Null
                $newItem.SubItems.Add($partObj.Source) | Out-Null
                $newItem.SubItems.Add($partObj.Location) | Out-Null
                $newItem.Tag = [PSCustomObject]@{
                    PartNumber = $partObj.PartNumber
                    Description = $partObj.Description
                    Quantity = $qty
                    Source = $partObj.Source
                    Location = $partObj.Location
                    OEMNumber = $partObj.OEMNumber
                }
                $selectedListView.Items.Add($newItem) | Out-Null
            }
        }
    })
    
    $removeButton.Add_Click({
        foreach ($item in $selectedListView.SelectedItems) {
            $selectedListView.Items.Remove($item)
        }
    })
    
    $saveButton.Add_Click({
        $partsToAdd = @()
        
        foreach ($item in $selectedListView.Items) {
            $partsToAdd += $item.Tag
        }
        
        if ($partsToAdd.Count -gt 0) {
            # Simple save function that doesn't depend on complex state
            if (Save-PartsToWorkOrder -WorkOrderNumber $workOrderNumber -Parts $partsToAdd) {
                # Success is already shown in Save-PartsToWorkOrder
                $form.Close()
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No parts selected to add to the work order.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    
    # Show the form
    $form.ShowDialog()
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

# Simple helper function to save parts to work order
function Save-PartsToWorkOrder {
    param(
        [string]$WorkOrderNumber,
        [array]$Parts
    )
    
    Write-Log "=== START Save-PartsToWorkOrder ==="
    Write-Log "Saving parts to Work Order: $WorkOrderNumber"
    Write-Log "Number of parts to save: $($Parts.Count)"
    
    # 1. Initialize/update global work order parts dictionary if needed
    if ($null -eq $script:workOrderParts) {
        Write-Log "Initializing workOrderParts dictionary"
        $script:workOrderParts = @{}
    }
    
    # 2. Find the item in the ListView 
    $targetItem = $null
    $targetIndex = -1
    
    for ($i = 0; $i -lt $script:listViewLaborLog.Items.Count; $i++) {
        $item = $script:listViewLaborLog.Items[$i]
        if ($item.SubItems[1].Text -eq $WorkOrderNumber) {
            $targetItem = $item
            $targetIndex = $i
            break
        }
    }
    
    if ($targetItem -eq $null) {
        Write-Log "WARNING: Work order $WorkOrderNumber not found in ListView"
        [System.Windows.Forms.MessageBox]::Show("Work order $WorkOrderNumber not found in labor logs", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # 3. For items without a proper work order number, update with row-based identifier
    if ($WorkOrderNumber -eq "Need W/O #") {
        $newWorkOrderNumber = "Need W/O #-$(New-Guid)"
        $targetItem.SubItems[1].Text = $newWorkOrderNumber
        $WorkOrderNumber = $newWorkOrderNumber
    }
    
    # 4. Add/update parts for this work order in our dictionary
    Write-Log "Updating workOrderParts dictionary for work order: $WorkOrderNumber"
    $script:workOrderParts[$WorkOrderNumber] = $Parts
    
    # 5. Load existing labor logs
    $laborLogsPath = Join-Path $config.LaborDirectory "LaborLogs.csv"
    Write-Log "Loading labor logs from: $laborLogsPath"
    
    if (-not (Test-Path $laborLogsPath)) {
        Write-Log "ERROR: Labor logs file not found at: $laborLogsPath"
        [System.Windows.Forms.MessageBox]::Show("Labor logs file not found at: $laborLogsPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    try {
        # 6. Build a formatted parts display for the ListView
        $partsSummary = ($Parts | ForEach-Object {
            "$($_.PartNumber) - Qty:$($_.Quantity)"
        }) -join ", "
        
        # 7. Update the parts column in the ListView (index 5 is Parts column)
        $targetItem.SubItems[5].Text = $partsSummary
        
        Write-Log ("Parts display updated in ListView for {0}: {1}" -f $WorkOrderNumber, $partsSummary)
        
        # 8. Save all labor logs with updated work order IDs and parts
        Save-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsPath
        
        # 9. Only show one success dialog
        [System.Windows.Forms.MessageBox]::Show("Parts added successfully to $WorkOrderNumber", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        Write-Log "=== END Save-PartsToWorkOrder ==="
        return $true
    }
    catch {
        Write-Log "ERROR in Save-PartsToWorkOrder: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("Error saving parts to work order: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
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
    $split.Panel1MinSize = 220
    $split.Panel2MinSize = 240
    $split.BackColor = [System.Drawing.Color]::FromArgb(220,225,232)
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
        [System.Windows.Forms.MessageBox]::Show("Historian wiring coming later.", "Job Log", "OK", "Information")
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
        $script:__pm_openPath = $m.HtmlPath
        $openBtn.Add_Click({ Open-PmArchivedHtml -Path $script:__pm_openPath })
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
        $script:__pm_cur_totalsLabel = $totals

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
        $lv.Tag = $m

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



        $script:__pm_curMachine = $m
        $lv.Add_ItemChecked({
            param($sender, $e)
            if ($script:pmSuppressEvents) { return }

            $machine = $script:__pm_curMachine
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

            $d = 0; $mins = 0
            foreach ($t in @($machine.Parsed.Tasks)) {
                $ee = $machine.State[$t.ItemNo]
                if ($ee -and $ee.Selected) {
                    $d++
                    $mins += if ($null -ne $ee.CustomTimeMin) { [int]$ee.CustomTimeMin } else { [int]$t.EstTimeMin }
                }
            }
            $tl = $sender.Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Dock -eq 'Bottom' }
            if ($tl) { $tl.Text = "Done: $d of $($machine.TaskCount)  •  $([Math]::Round($mins/60.0, 2)) h" }

            Update-PmSelectedHoursTotal
        })

        # Wire double-click editor
        $lv.Add_DoubleClick({
            param($sender, $e)
            if ($script:pmSuppressEvents) { return }

            $machine = $script:__pm_curMachine
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
            $d = 0; $mins = 0
            foreach ($t in @($machine.Parsed.Tasks)) {
                $ee = $machine.State[$t.ItemNo]
                if ($ee -and $ee.Selected) {
                    $d++
                    $mins += if ($null -ne $ee.CustomTimeMin) { [int]$ee.CustomTimeMin } else { [int]$t.EstTimeMin }
                }
            }
            $tl = $sender.Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Dock -eq 'Bottom' }
            if ($tl) { $tl.Text = "Done: $d of $($machine.TaskCount)  •  $([Math]::Round($mins/60.0, 2)) h" }

            Update-PmSelectedHoursTotal
        })

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

    $subTabs.TabPages.Add((New-SubTab -Title "Weekly Worksheet" -PlaceholderText "Weekly Worksheet — fetch eDAC, edit actual times, sync")) | Out-Null
    $subTabs.TabPages.Add((New-SubTab -Title "Historian"        -PlaceholderText "Historian — browse sent events, generate ASCII reports")) | Out-Null
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

function Load-WtSettingsValues {
    $c = $script:wtSettings
    $wt = $script:config.WorkTracking

    $c.TechNameBox.Text  = if ($wt.TechnicianName) { $wt.TechnicianName } else { '' }
    $c.EmailBox.Text     = if ($script:config.SupervisorEmail) { $script:config.SupervisorEmail } else { '' }
    $c.EdacUrlBox.Text   = if ($wt.eDacUrl) { $wt.eDacUrl } else { '' }
    $c.ThresholdBox.Value = if ($wt.ReactiveToWorkOrderMinutes) { [int]$wt.ReactiveToWorkOrderMinutes } else { 15 }

    $wb = $wt.WorkBudget
    if ($wb) {
        $c.DayLenBox.Value  = if ($wb.DayLengthHours) { [decimal]$wb.DayLengthHours } else { 8.5 }
        $c.LunchBox.Value   = if ($wb.LunchMinutes)   { [decimal]$wb.LunchMinutes }   else { 30 }
        $c.WashupBox.Value  = if ($wb.WashupMinutes)  { [decimal]$wb.WashupMinutes }  else { 15 }
        $c.StartupBox.Value = if ($wb.StartupMinutes) { [decimal]$wb.StartupMinutes } else { 15 }
        $pbm = @($wb.PaidBreaksMinutes)
        $c.Break1Box.Value  = if ($pbm.Count -ge 1 -and $pbm[0]) { [decimal]$pbm[0] } else { 15 }
        $c.Break2Box.Value  = if ($pbm.Count -ge 2 -and $pbm[1]) { [decimal]$pbm[1] } else { 15 }
        $c.EndWashBox.Value = if ($wb.EndOfDayWashupMinutes) { [decimal]$wb.EndOfDayWashupMinutes } else { 15 }
        $c.PaperBox.Value   = if ($wb.PaperworkMinutes)      { [decimal]$wb.PaperworkMinutes }      else { 15 }
        $c.TargetBox.Value  = if ($wb.WorkTargetHours)       { [decimal]$wb.WorkTargetHours }       else { 6.5 }
    }
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

    # --- Body: FlowLayoutPanel of GroupBoxes ---
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

    New-SettingRow -Parent $gBudget -Top  28 -LabelText "Day Length (hours):"              -InputControl $dayLenBox
    New-SettingRow -Parent $gBudget -Top  62 -LabelText "Lunch (min, 0=skip):"             -InputControl $lunchBox
    New-SettingRow -Parent $gBudget -Top  96 -LabelText "Wash-up at Lunch (min):"          -InputControl $washupBox
    New-SettingRow -Parent $gBudget -Top 130 -LabelText "Startup (min):"                   -InputControl $startupBox
    New-SettingRow -Parent $gBudget -Top 164 -LabelText "Paid Break 1 (min):"              -InputControl $break1Box
    New-SettingRow -Parent $gBudget -Top 198 -LabelText "Paid Break 2 (min):"              -InputControl $break2Box
    New-SettingRow -Parent $gBudget -Top 232 -LabelText "Wash-up End-of-Day (min):"        -InputControl $endWashBox
    New-SettingRow -Parent $gBudget -Top 266 -LabelText "Paperwork (min):"                 -InputControl $paperBox
    New-SettingRow -Parent $gBudget -Top 300 -LabelText "Work Target (hours):"             -InputControl $targetBox

    # --- Store control references in script scope so handlers can find them ---
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

    # --- Load helpers ---
    Load-WtSettingsValues

    # --- Save handler ---
    $saveBtn.Add_Click({
        try {
            $c = $script:wtSettings
            $cfg = $script:config
            if (-not $cfg.WorkTracking) { throw "WorkTracking section missing from config." }

            $cfg.WorkTracking.TechnicianName             = $c.TechNameBox.Text.Trim()
            $cfg.SupervisorEmail                         = $c.EmailBox.Text.Trim()
            $cfg.WorkTracking.eDacUrl                    = $c.EdacUrlBox.Text.Trim()
            $cfg.WorkTracking.ReactiveToWorkOrderMinutes = [int]$c.ThresholdBox.Value

            $cfg.WorkTracking.WorkBudget.DayLengthHours        = [double]$c.DayLenBox.Value
            $cfg.WorkTracking.WorkBudget.LunchMinutes          = [int]$c.LunchBox.Value
            $cfg.WorkTracking.WorkBudget.WashupMinutes         = [int]$c.WashupBox.Value
            $cfg.WorkTracking.WorkBudget.StartupMinutes        = [int]$c.StartupBox.Value
            $cfg.WorkTracking.WorkBudget.PaidBreaksMinutes     = @([int]$c.Break1Box.Value, [int]$c.Break2Box.Value)
            $cfg.WorkTracking.WorkBudget.EndOfDayWashupMinutes = [int]$c.EndWashBox.Value
            $cfg.WorkTracking.WorkBudget.PaperworkMinutes      = [int]$c.PaperBox.Value
            $cfg.WorkTracking.WorkBudget.WorkTargetHours       = [double]$c.TargetBox.Value

            $cfg | ConvertTo-Json -Depth 12 | Set-Content -Path $script:configPath -Encoding UTF8
            $script:config = $cfg
            Write-Log "Settings saved."
            [System.Windows.Forms.MessageBox]::Show("Settings saved.", "Settings", "OK", "Information")
        } catch {
            Write-Log "Error saving settings: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error saving settings: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    $reloadBtn.Add_Click({
        try {
            $script:config = Get-Content -Path $script:configPath -Raw | ConvertFrom-Json
            Load-WtSettingsValues
            Write-Log "Settings reloaded from file."
            [System.Windows.Forms.MessageBox]::Show("Settings reloaded.", "Settings", "OK", "Information")
        } catch {
            Write-Log "Error reloading settings: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error reloading settings: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    Write-Log "Settings tab setup completed."
}

# Function to set up the Call Logs tab
function Setup-CallLogsTab {
    param($parentTab)

    Write-Log "Setting up Call Logs tab..."

    # Root container
    $callLogsPanel = New-Object System.Windows.Forms.Panel
    $callLogsPanel.Dock = 'Fill'
    $callLogsPanel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $callLogsPanel.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($callLogsPanel)

    # Bottom action bar (Dock=Bottom first, so ListView fill goes above it)
    $actionBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionBar.Dock = 'Bottom'
    $actionBar.Height = 52
    $actionBar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $actionBar.WrapContents = $false
    $actionBar.BackColor = [System.Drawing.Color]::Transparent
    $actionBar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $callLogsPanel.Controls.Add($actionBar)

    # Call Log ListView (Dock=Fill)
    $script:listViewCallLogs = New-Object System.Windows.Forms.ListView
    $script:listViewCallLogs.Dock = 'Fill'
    $script:listViewCallLogs.Columns.Clear()
    $script:listViewCallLogs.Columns.Add("Date", 100)      | Out-Null
    $script:listViewCallLogs.Columns.Add("Machine ID", 140) | Out-Null
    $script:listViewCallLogs.Columns.Add("Cause", 100)     | Out-Null
    $script:listViewCallLogs.Columns.Add("Action", 100)    | Out-Null
    $script:listViewCallLogs.Columns.Add("Noun", 100)      | Out-Null
    $script:listViewCallLogs.Columns.Add("Time Down", 80)  | Out-Null
    $script:listViewCallLogs.Columns.Add("Time Up", 80)    | Out-Null
    $script:listViewCallLogs.Columns.Add("Notes", 200)     | Out-Null
    Set-ListViewStyle -ListView $script:listViewCallLogs
    $callLogsPanel.Controls.Add($script:listViewCallLogs)
    $script:listViewCallLogs.BringToFront()

    # Load existing call logs
    Load-Logs -listView $script:listViewCallLogs -filePath $callLogsFilePath


    # ---------- Add New Call Log button (unchanged sub-form logic) ----------
    $addCallLogButton = New-Object System.Windows.Forms.Button
    $addCallLogButton.Size = New-Object System.Drawing.Size(160, 36)
    $addCallLogButton.Text = "Add New Call Log"
    $addCallLogButton.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $addCallLogButton.ForeColor = [System.Drawing.Color]::White
    $addCallLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addCallLogButton.FlatAppearance.BorderSize = 0
    $addCallLogButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $addCallLogButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $addCallLogButton.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)

    $addCallLogButton.Add_Click({
        $addCallLogForm = New-Object System.Windows.Forms.Form
        $addCallLogForm.Text = "Add New Call Log"
        $addCallLogForm.Size = New-Object System.Drawing.Size(400, 500)
        $addCallLogForm.StartPosition = 'CenterParent'
        $addCallLogForm.FormBorderStyle = 'FixedDialog'
        $addCallLogForm.MaximizeBox = $false
        $addCallLogForm.MinimizeBox = $false
        $addCallLogForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
        $addCallLogForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # ---- (unchanged internals) ----
        $labelDate = New-Object System.Windows.Forms.Label
        $labelDate.Location = New-Object System.Drawing.Point(10, 20)
        $labelDate.Size = New-Object System.Drawing.Size(100, 20)
        $labelDate.Text = "Date:"
        $addCallLogForm.Controls.Add($labelDate)

        $textBoxDate = New-Object System.Windows.Forms.TextBox
        $textBoxDate.Location = New-Object System.Drawing.Point(120, 20)
        $textBoxDate.Size = New-Object System.Drawing.Size(250, 20)
        $textBoxDate.Text = Get-Date -Format "yyyy-MM-dd"
        $addCallLogForm.Controls.Add($textBoxDate)

        $labelMachineId = New-Object System.Windows.Forms.Label
        $labelMachineId.Location = New-Object System.Drawing.Point(10, 50)
        $labelMachineId.Size = New-Object System.Drawing.Size(100, 20)
        $labelMachineId.Text = "Machine ID:"
        $addCallLogForm.Controls.Add($labelMachineId)

        $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox
        $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 50)
        $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20)
        $addCallLogForm.Controls.Add($comboBoxMachineId)
        Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"

        $labelCause = New-Object System.Windows.Forms.Label
        $labelCause.Location = New-Object System.Drawing.Point(10, 80)
        $labelCause.Size = New-Object System.Drawing.Size(100, 20)
        $labelCause.Text = "Cause:"
        $addCallLogForm.Controls.Add($labelCause)

        $comboBoxCause = New-Object System.Windows.Forms.ComboBox
        $comboBoxCause.Location = New-Object System.Drawing.Point(120, 80)
        $comboBoxCause.Size = New-Object System.Drawing.Size(250, 20)
        $addCallLogForm.Controls.Add($comboBoxCause)
        Load-ComboBoxData -comboBox $comboBoxCause -csvName "Causes"

        $labelAction = New-Object System.Windows.Forms.Label
        $labelAction.Location = New-Object System.Drawing.Point(10, 110)
        $labelAction.Size = New-Object System.Drawing.Size(100, 20)
        $labelAction.Text = "Action:"
        $addCallLogForm.Controls.Add($labelAction)

        $comboBoxAction = New-Object System.Windows.Forms.ComboBox
        $comboBoxAction.Location = New-Object System.Drawing.Point(120, 110)
        $comboBoxAction.Size = New-Object System.Drawing.Size(250, 20)
        $addCallLogForm.Controls.Add($comboBoxAction)
        Load-ComboBoxData -comboBox $comboBoxAction -csvName "Actions"

        $labelNoun = New-Object System.Windows.Forms.Label
        $labelNoun.Location = New-Object System.Drawing.Point(10, 140)
        $labelNoun.Size = New-Object System.Drawing.Size(100, 20)
        $labelNoun.Text = "Noun:"
        $addCallLogForm.Controls.Add($labelNoun)

        $comboBoxNoun = New-Object System.Windows.Forms.ComboBox
        $comboBoxNoun.Location = New-Object System.Drawing.Point(120, 140)
        $comboBoxNoun.Size = New-Object System.Drawing.Size(250, 20)
        $addCallLogForm.Controls.Add($comboBoxNoun)
        Load-ComboBoxData -comboBox $comboBoxNoun -csvName "Nouns"

        $labelTimeDown = New-Object System.Windows.Forms.Label
        $labelTimeDown.Location = New-Object System.Drawing.Point(10, 200)
        $labelTimeDown.Size = New-Object System.Drawing.Size(100, 20)
        $labelTimeDown.Text = "Time Down:"
        $addCallLogForm.Controls.Add($labelTimeDown)

        $comboBoxTimeDownHour = New-Object System.Windows.Forms.ComboBox
        $comboBoxTimeDownHour.Location = New-Object System.Drawing.Point(120, 200)
        $comboBoxTimeDownHour.Size = New-Object System.Drawing.Size(50, 20)
        $comboBoxTimeDownHour.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        0..23 | ForEach-Object { $comboBoxTimeDownHour.Items.Add($_.ToString("00")) }
        $comboBoxTimeDownHour.SelectedIndex = 0
        $addCallLogForm.Controls.Add($comboBoxTimeDownHour)

        $labelTimeDownSeparator = New-Object System.Windows.Forms.Label
        $labelTimeDownSeparator.Location = New-Object System.Drawing.Point(175, 203)
        $labelTimeDownSeparator.Size = New-Object System.Drawing.Size(10, 20)
        $labelTimeDownSeparator.Text = ":"
        $addCallLogForm.Controls.Add($labelTimeDownSeparator)

        $comboBoxTimeDownMinute = New-Object System.Windows.Forms.ComboBox
        $comboBoxTimeDownMinute.Location = New-Object System.Drawing.Point(190, 200)
        $comboBoxTimeDownMinute.Size = New-Object System.Drawing.Size(50, 20)
        $comboBoxTimeDownMinute.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        0..59 | ForEach-Object { $comboBoxTimeDownMinute.Items.Add($_.ToString("00")) }
        $comboBoxTimeDownMinute.SelectedIndex = 0
        $addCallLogForm.Controls.Add($comboBoxTimeDownMinute)

        $buttonTimeDownNow = New-Object System.Windows.Forms.Button
        $buttonTimeDownNow.Location = New-Object System.Drawing.Point(250, 200)
        $buttonTimeDownNow.Size = New-Object System.Drawing.Size(50, 20)
        $buttonTimeDownNow.Text = "Now"
        $buttonTimeDownNow.Add_Click({
            $now = Get-Date
            $comboBoxTimeDownHour.SelectedItem = $now.ToString("HH")
            $comboBoxTimeDownMinute.SelectedItem = $now.ToString("mm")
        })
        $addCallLogForm.Controls.Add($buttonTimeDownNow)

        $labelTimeUp = New-Object System.Windows.Forms.Label
        $labelTimeUp.Location = New-Object System.Drawing.Point(10, 230)
        $labelTimeUp.Size = New-Object System.Drawing.Size(100, 20)
        $labelTimeUp.Text = "Time Up:"
        $addCallLogForm.Controls.Add($labelTimeUp)

        $comboBoxTimeUpHour = New-Object System.Windows.Forms.ComboBox
        $comboBoxTimeUpHour.Location = New-Object System.Drawing.Point(120, 230)
        $comboBoxTimeUpHour.Size = New-Object System.Drawing.Size(50, 20)
        $comboBoxTimeUpHour.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        0..23 | ForEach-Object { $comboBoxTimeUpHour.Items.Add($_.ToString("00")) }
        $comboBoxTimeUpHour.SelectedIndex = 0
        $addCallLogForm.Controls.Add($comboBoxTimeUpHour)

        $labelTimeUpSeparator = New-Object System.Windows.Forms.Label
        $labelTimeUpSeparator.Location = New-Object System.Drawing.Point(175, 233)
        $labelTimeUpSeparator.Size = New-Object System.Drawing.Size(10, 20)
        $labelTimeUpSeparator.Text = ":"
        $addCallLogForm.Controls.Add($labelTimeUpSeparator)

        $comboBoxTimeUpMinute = New-Object System.Windows.Forms.ComboBox
        $comboBoxTimeUpMinute.Location = New-Object System.Drawing.Point(190, 230)
        $comboBoxTimeUpMinute.Size = New-Object System.Drawing.Size(50, 20)
        $comboBoxTimeUpMinute.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        0..59 | ForEach-Object { $comboBoxTimeUpMinute.Items.Add($_.ToString("00")) }
        $comboBoxTimeUpMinute.SelectedIndex = 0
        $addCallLogForm.Controls.Add($comboBoxTimeUpMinute)

        $buttonTimeUpNow = New-Object System.Windows.Forms.Button
        $buttonTimeUpNow.Location = New-Object System.Drawing.Point(250, 230)
        $buttonTimeUpNow.Size = New-Object System.Drawing.Size(50, 20)
        $buttonTimeUpNow.Text = "Now"
        $buttonTimeUpNow.Add_Click({
            $now = Get-Date
            $comboBoxTimeUpHour.SelectedItem = $now.ToString("HH")
            $comboBoxTimeUpMinute.SelectedItem = $now.ToString("mm")
        })
        $addCallLogForm.Controls.Add($buttonTimeUpNow)

        $labelNotes = New-Object System.Windows.Forms.Label
        $labelNotes.Location = New-Object System.Drawing.Point(10, 260)
        $labelNotes.Size = New-Object System.Drawing.Size(100, 20)
        $labelNotes.Text = "Notes:"
        $addCallLogForm.Controls.Add($labelNotes)

        $textBoxNotes = New-Object System.Windows.Forms.TextBox
        $textBoxNotes.Location = New-Object System.Drawing.Point(120, 260)
        $textBoxNotes.Size = New-Object System.Drawing.Size(250, 60)
        $textBoxNotes.Multiline = $true
        $addCallLogForm.Controls.Add($textBoxNotes)

        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(150, 330)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add"
        $addButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
        $addButton.ForeColor = [System.Drawing.Color]::White
        $addButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $addButton.FlatAppearance.BorderSize = 0
        $addButton.Add_Click({
            $item = New-Object System.Windows.Forms.ListViewItem($textBoxDate.Text)
            $item.SubItems.Add($comboBoxMachineId.SelectedItem)
            $item.SubItems.Add($comboBoxCause.SelectedItem)
            $item.SubItems.Add($comboBoxAction.SelectedItem)
            $item.SubItems.Add($comboBoxNoun.SelectedItem)
            $timeDown = "$($comboBoxTimeDownHour.SelectedItem):$($comboBoxTimeDownMinute.SelectedItem)"
            $timeUp   = "$($comboBoxTimeUpHour.SelectedItem):$($comboBoxTimeUpMinute.SelectedItem)"
            $item.SubItems.Add($timeDown)
            $item.SubItems.Add($timeUp)
            $item.SubItems.Add($textBoxNotes.Text)

            $script:listViewCallLogs.Items.Add($item)
            Save-Logs -listView $script:listViewCallLogs -filePath $callLogsFilePath
            $addCallLogForm.Close()
        })
        $addCallLogForm.Controls.Add($addButton)
        $addCallLogForm.ShowDialog()
    })
    $actionBar.Controls.Add($addCallLogButton)

    # ---------- Add Machine button ----------
    $addMachineButton = New-Object System.Windows.Forms.Button
    $addMachineButton.Size = New-Object System.Drawing.Size(150, 36)
    $addMachineButton.Text = "Add Machine"
    $addMachineButton.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $addMachineButton.ForeColor = [System.Drawing.Color]::White
    $addMachineButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addMachineButton.FlatAppearance.BorderSize = 0
    $addMachineButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $addMachineButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $addMachineButton.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
    $addMachineButton.Add_Click({
        $addMachineForm = New-Object System.Windows.Forms.Form
        $addMachineForm.Text = "Add New Machine"
        $addMachineForm.Size = New-Object System.Drawing.Size(300, 220)
        $addMachineForm.StartPosition = 'CenterParent'
        $addMachineForm.FormBorderStyle = 'FixedDialog'
        $addMachineForm.MaximizeBox = $false
        $addMachineForm.MinimizeBox = $false
        $addMachineForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        $labelAcronym = New-Object System.Windows.Forms.Label
        $labelAcronym.Location = New-Object System.Drawing.Point(10, 20)
        $labelAcronym.Size = New-Object System.Drawing.Size(120, 20)
        $labelAcronym.Text = "Machine Acronym:"
        $addMachineForm.Controls.Add($labelAcronym)

        $textBoxAcronym = New-Object System.Windows.Forms.TextBox
        $textBoxAcronym.Location = New-Object System.Drawing.Point(130, 20)
        $textBoxAcronym.Size = New-Object System.Drawing.Size(150, 20)
        $addMachineForm.Controls.Add($textBoxAcronym)

        $labelEquipmentNumber = New-Object System.Windows.Forms.Label
        $labelEquipmentNumber.Location = New-Object System.Drawing.Point(10, 50)
        $labelEquipmentNumber.Size = New-Object System.Drawing.Size(120, 20)
        $labelEquipmentNumber.Text = "Equipment Number:"
        $addMachineForm.Controls.Add($labelEquipmentNumber)

        $textBoxEquipmentNumber = New-Object System.Windows.Forms.TextBox
        $textBoxEquipmentNumber.Location = New-Object System.Drawing.Point(130, 50)
        $textBoxEquipmentNumber.Size = New-Object System.Drawing.Size(150, 20)
        $addMachineForm.Controls.Add($textBoxEquipmentNumber)

        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(100, 100)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add Machine"
        $addButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
        $addButton.ForeColor = [System.Drawing.Color]::White
        $addButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $addButton.FlatAppearance.BorderSize = 0
        $addButton.Add_Click({
            $machineAcronym = $textBoxAcronym.Text.Trim()
            $equipmentNumber = $textBoxEquipmentNumber.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($machineAcronym) -or [string]::IsNullOrWhiteSpace($equipmentNumber)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter both Machine Acronym and Equipment Number.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $machinesCsvPath = Join-Path $config.DropdownCsvsDirectory "Machines.csv"

            if (-not (Test-Path $machinesCsvPath)) {
                "Machine Acronym,Machine Number" | Out-File -FilePath $machinesCsvPath -Encoding utf8
                Write-Log "Created new Machines.csv file at $machinesCsvPath"
            }

            $existingData = Import-Csv -Path $machinesCsvPath

            if ($existingData | Where-Object { $_.'Machine Acronym' -eq $machineAcronym -and $_.'Machine Number' -eq $equipmentNumber }) {
                [System.Windows.Forms.MessageBox]::Show("This machine and equipment number combination already exists.", "Duplicate Entry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            "$machineAcronym,$equipmentNumber" | Out-File -FilePath $machinesCsvPath -Append -Encoding utf8
            Write-Log "Added new machine: $machineAcronym with equipment number: $equipmentNumber to $machinesCsvPath"

            $script:machinesUpdated = $true

            [System.Windows.Forms.MessageBox]::Show("Machine added successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $addMachineForm.Close()
        })
        $addMachineForm.Controls.Add($addButton)
        $addMachineForm.ShowDialog()
    })
    $actionBar.Controls.Add($addMachineButton)

    # ---------- Send Logs button ----------
    $sendLogsButton = New-Object System.Windows.Forms.Button
    $sendLogsButton.Size = New-Object System.Drawing.Size(150, 36)
    $sendLogsButton.Text = "Send Logs"
    $sendLogsButton.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $sendLogsButton.ForeColor = [System.Drawing.Color]::White
    $sendLogsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $sendLogsButton.FlatAppearance.BorderSize = 0
    $sendLogsButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $sendLogsButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $sendLogsButton.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
    $sendLogsButton.Add_Click({
        $sendLogsForm = New-Object System.Windows.Forms.Form
        $sendLogsForm.Text = "Send Logs"
        $sendLogsForm.Size = New-Object System.Drawing.Size(300, 200)
        $sendLogsForm.StartPosition = 'CenterParent'
        $sendLogsForm.FormBorderStyle = 'FixedDialog'
        $sendLogsForm.MaximizeBox = $false
        $sendLogsForm.MinimizeBox = $false
        $sendLogsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        $labelDate = New-Object System.Windows.Forms.Label
        $labelDate.Location = New-Object System.Drawing.Point(10, 20)
        $labelDate.Size = New-Object System.Drawing.Size(100, 20)
        $labelDate.Text = "Select Date:"
        $sendLogsForm.Controls.Add($labelDate)

        $dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
        $dateTimePicker.Location = New-Object System.Drawing.Point(120, 20)
        $dateTimePicker.Size = New-Object System.Drawing.Size(150, 20)
        $dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
        $sendLogsForm.Controls.Add($dateTimePicker)

        $sendButton = New-Object System.Windows.Forms.Button
        $sendButton.Location = New-Object System.Drawing.Point(100, 100)
        $sendButton.Size = New-Object System.Drawing.Size(100, 30)
        $sendButton.Text = "Save to File"
        $sendButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
        $sendButton.ForeColor = [System.Drawing.Color]::White
        $sendButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $sendButton.FlatAppearance.BorderSize = 0
        $sendButton.Add_Click({
            $selectedDate = $dateTimePicker.Value.ToString("yyyy-MM-dd")
            $logsForDate = $script:listViewCallLogs.Items | Where-Object { $_.SubItems[0].Text -eq $selectedDate }

            if ($logsForDate.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No logs found for the selected date.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $logContent = $logsForDate | ForEach-Object {
                $machine = $_.SubItems[1].Text
                $equipmentNumber = $_.SubItems[2].Text
                $cause = $_.SubItems[3].Text
                $action = $_.SubItems[4].Text
                $noun = $_.SubItems[5].Text
                $timeDown = $_.SubItems[6].Text
                $timeUp = $_.SubItems[7].Text
                $notes = $_.SubItems[8].Text
                "Machine: $machine`r`nEquipment Number: $equipmentNumber`r`nCause: $cause`r`nAction: $action`r`nNoun: $noun`r`nTime Down: $timeDown`r`nTime Up: $timeUp`r`nNotes: $notes`r`n`r`n"
            }

            $content = "Call Logs for $selectedDate`r`n`r`n" + ($logContent -join "`r`n")

            try {
                $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
                $saveFileDialog.FileName = "CallLogs_$selectedDate.txt"
                $saveFileDialog.Title = "Save Call Logs"

                if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $filePath = $saveFileDialog.FileName
                    $content | Out-File -FilePath $filePath -Encoding utf8
                    Write-Log "Call logs for $selectedDate saved to $filePath"
                    [System.Windows.Forms.MessageBox]::Show("Call logs have been saved to $filePath", "Logs Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    Write-Log "Log file save cancelled by user"
                }
            } catch {
                Write-Log "Error saving log file: $_"
                [System.Windows.Forms.MessageBox]::Show("Error saving log file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            $sendLogsForm.Close()
        })
        $sendLogsForm.Controls.Add($sendButton)
        $sendLogsForm.ShowDialog()
    })
    $actionBar.Controls.Add($sendLogsButton)

    Write-Log "Call Logs tab setup completed."
}

# Function to set up the Labor Log tab
function Setup-LaborLogTab {
    param($parentTab, $tabControl)

    Write-Log "Setting up Labor Log tab..."

    $laborLogPanel = New-Object System.Windows.Forms.Panel
    $laborLogPanel.Dock = 'Fill'
    $laborLogPanel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $laborLogPanel.Padding = New-Object System.Windows.Forms.Padding(12)
    $parentTab.Controls.Add($laborLogPanel)

    # Bottom button bar
    $buttonBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonBar.Dock = 'Bottom'
    $buttonBar.Height = 52
    $buttonBar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonBar.WrapContents = $false
    $buttonBar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $buttonBar.BackColor = [System.Drawing.Color]::Transparent
    $laborLogPanel.Controls.Add($buttonBar)

    # Labor Log ListView
    $script:listViewLaborLog = New-Object System.Windows.Forms.ListView
    $script:listViewLaborLog.Dock = 'Fill'
    $script:listViewLaborLog.Scrollable = $true
    $script:listViewLaborLog.ShowItemToolTips = $true
    $script:listViewLaborLog.Columns.Add("Date", 100)        | Out-Null
    $script:listViewLaborLog.Columns.Add("Work Order", 150)  | Out-Null
    $script:listViewLaborLog.Columns.Add("Description", 300)| Out-Null
    $script:listViewLaborLog.Columns.Add("Machine", 100)     | Out-Null
    $script:listViewLaborLog.Columns.Add("Duration", 100)    | Out-Null
    $script:listViewLaborLog.Columns.Add("Parts", 300)       | Out-Null
    $script:listViewLaborLog.Columns.Add("Notes", 180)       | Out-Null
    Set-ListViewStyle -ListView $script:listViewLaborLog
    $laborLogPanel.Controls.Add($script:listViewLaborLog)
    $script:listViewLaborLog.BringToFront()

    # Tooltip
    $script:listViewToolTip = New-Object System.Windows.Forms.ToolTip
    $script:listViewLaborLog.Add_MouseMove({
        param($sender, $e)
        $item = $script:listViewLaborLog.GetItemAt($e.X, $e.Y)
        if ($item -ne $null) {
            $script:listViewToolTip.SetToolTip($script:listViewLaborLog, $item.SubItems[5].Text)
        } else {
            $script:listViewToolTip.SetToolTip($script:listViewLaborLog, "")
        }
    })

    # Double-click details (unchanged)
    $script:listViewLaborLog.Add_DoubleClick({
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $item = $selectedItems[0]
            $workOrderNumber = $item.SubItems[1].Text
            if ($script:workOrderParts.ContainsKey($workOrderNumber)) {
                $parts = $script:workOrderParts[$workOrderNumber]
                $detailsForm = New-Object System.Windows.Forms.Form
                $detailsForm.Text = "Parts Details for Work Order #$workOrderNumber"
                $detailsForm.Size = New-Object System.Drawing.Size(600, 400)
                $detailsForm.StartPosition = 'CenterParent'
                $detailsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

                $detailsListView = New-Object System.Windows.Forms.ListView
                $detailsListView.Dock = 'Fill'
                $detailsListView.Columns.Add("Part Number", 100)
                $detailsListView.Columns.Add("OEM", 100)
                $detailsListView.Columns.Add("Quantity", 100)
                $detailsListView.Columns.Add("Location", 100)
                $detailsListView.Columns.Add("Source", 100)
                Set-ListViewStyle -ListView $detailsListView

                foreach ($part in $parts) {
                    $partItem = New-Object System.Windows.Forms.ListViewItem($part.PartNumber)
                    $partItem.SubItems.Add($part.PartNo)
                    $partItem.SubItems.Add($part.Quantity.ToString())
                    $partItem.SubItems.Add($part.Location)
                    $partItem.SubItems.Add($part.Source)
                    $detailsListView.Items.Add($partItem)
                }

                $detailsForm.Controls.Add($detailsListView)
                $detailsForm.ShowDialog()
            }
        }
    })

    # Notification icon (top-right of tab)
    $script:notificationIcon = New-Object System.Windows.Forms.Label
    $script:notificationIcon.Text = "•"
    $script:notificationIcon.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
    $script:notificationIcon.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $script:notificationIcon.AutoSize = $true
    $script:notificationIcon.BackColor = [System.Drawing.Color]::Transparent
    $script:notificationIcon.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:notificationIcon.Location = New-Object System.Drawing.Point(($parentTab.Width - 40), 6)
    $script:notificationIcon.Visible = $false
    $parentTab.Controls.Add($script:notificationIcon)
    $script:notificationIcon.BringToFront()

    # Load existing labor logs
    Load-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath

    # Helper for styled action buttons on the button bar
    $makeActionButton = {
        param([string]$Text, [System.Drawing.Color]$Back)
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $Text
        $b.Size = New-Object System.Drawing.Size(160, 36)
        $b.BackColor = $Back
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $b.FlatAppearance.BorderSize = 0
        $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $b.Cursor = [System.Windows.Forms.Cursors]::Hand
        $b.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
        return $b
    }

    # Add Labor Log Entry
    $addLaborLogButton = & $makeActionButton "Add Labor Log Entry" ([System.Drawing.Color]::FromArgb(52,152,219))
    $addLaborLogButton.Add_Click({
        $addLaborLogForm = New-Object System.Windows.Forms.Form
        $addLaborLogForm.Text = "Add Labor Log Entry"
        $addLaborLogForm.Size = New-Object System.Drawing.Size(400, 420)
        $addLaborLogForm.StartPosition = 'CenterParent'
        $addLaborLogForm.FormBorderStyle = 'FixedDialog'
        $addLaborLogForm.MaximizeBox = $false
        $addLaborLogForm.MinimizeBox = $false
        $addLaborLogForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
        $addLaborLogForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        $labelDate = New-Object System.Windows.Forms.Label; $labelDate.Location = New-Object System.Drawing.Point(10, 20); $labelDate.Size = New-Object System.Drawing.Size(100, 20); $labelDate.Text = "Date:"; $addLaborLogForm.Controls.Add($labelDate)
        $textBoxDate = New-Object System.Windows.Forms.TextBox; $textBoxDate.Location = New-Object System.Drawing.Point(120, 20); $textBoxDate.Size = New-Object System.Drawing.Size(250, 20); $textBoxDate.Text = Get-Date -Format "yyyy-MM-dd"; $addLaborLogForm.Controls.Add($textBoxDate)

        $labelWorkOrder = New-Object System.Windows.Forms.Label; $labelWorkOrder.Location = New-Object System.Drawing.Point(10, 50); $labelWorkOrder.Size = New-Object System.Drawing.Size(100, 20); $labelWorkOrder.Text = "Work Order #:"; $addLaborLogForm.Controls.Add($labelWorkOrder)
        $textBoxWorkOrder = New-Object System.Windows.Forms.TextBox; $textBoxWorkOrder.Location = New-Object System.Drawing.Point(120, 50); $textBoxWorkOrder.Size = New-Object System.Drawing.Size(250, 20); $addLaborLogForm.Controls.Add($textBoxWorkOrder)

        $labelTask = New-Object System.Windows.Forms.Label; $labelTask.Location = New-Object System.Drawing.Point(10, 80); $labelTask.Size = New-Object System.Drawing.Size(100, 20); $labelTask.Text = "Task:"; $addLaborLogForm.Controls.Add($labelTask)
        $textBoxTask = New-Object System.Windows.Forms.TextBox; $textBoxTask.Location = New-Object System.Drawing.Point(120, 80); $textBoxTask.Size = New-Object System.Drawing.Size(250, 60); $textBoxTask.Multiline = $true; $addLaborLogForm.Controls.Add($textBoxTask)

        $labelMachineId = New-Object System.Windows.Forms.Label; $labelMachineId.Location = New-Object System.Drawing.Point(10, 150); $labelMachineId.Size = New-Object System.Drawing.Size(100, 20); $labelMachineId.Text = "Machine ID:"; $addLaborLogForm.Controls.Add($labelMachineId)
        $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox; $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 150); $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20); $addLaborLogForm.Controls.Add($comboBoxMachineId); Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"

        $labelDuration = New-Object System.Windows.Forms.Label; $labelDuration.Location = New-Object System.Drawing.Point(10, 180); $labelDuration.Size = New-Object System.Drawing.Size(100, 20); $labelDuration.Text = "Duration:"; $addLaborLogForm.Controls.Add($labelDuration)
        $textBoxDuration = New-Object System.Windows.Forms.TextBox; $textBoxDuration.Location = New-Object System.Drawing.Point(120, 180); $textBoxDuration.Size = New-Object System.Drawing.Size(250, 20); $addLaborLogForm.Controls.Add($textBoxDuration)

        $labelNotes = New-Object System.Windows.Forms.Label; $labelNotes.Location = New-Object System.Drawing.Point(10, 210); $labelNotes.Size = New-Object System.Drawing.Size(100, 20); $labelNotes.Text = "Notes:"; $addLaborLogForm.Controls.Add($labelNotes)
        $textBoxNotes = New-Object System.Windows.Forms.TextBox; $textBoxNotes.Location = New-Object System.Drawing.Point(120, 210); $textBoxNotes.Size = New-Object System.Drawing.Size(250, 60); $textBoxNotes.Multiline = $true; $addLaborLogForm.Controls.Add($textBoxNotes)

        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(150, 320)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add"
        $addButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
        $addButton.ForeColor = [System.Drawing.Color]::White
        $addButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $addButton.FlatAppearance.BorderSize = 0
        $addButton.Add_Click({
            $workOrderNumber = if ([string]::IsNullOrWhiteSpace($textBoxWorkOrder.Text)) { "Need W/O #" } else { $textBoxWorkOrder.Text }
            $item = New-Object System.Windows.Forms.ListViewItem($textBoxDate.Text)
            $item.SubItems.Add($workOrderNumber)
            $item.SubItems.Add($textBoxTask.Text)
            $item.SubItems.Add($comboBoxMachineId.SelectedItem)
            $item.SubItems.Add($textBoxDuration.Text)
            $item.SubItems.Add("")
            $item.SubItems.Add($textBoxNotes.Text)
            $script:listViewLaborLog.Items.Add($item)

            if ($workOrderNumber -eq "Need W/O #") {
                $key = "$($textBoxDate.Text)_$($comboBoxMachineId.SelectedItem)_$($textBoxTask.Text)"
                $script:unacknowledgedEntries[$key] = $true
                Update-NotificationIcon
            }
            Save-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath
            $addLaborLogForm.Close()
        })
        $addLaborLogForm.Controls.Add($addButton)
        $addLaborLogForm.ShowDialog()
    })
    $buttonBar.Controls.Add($addLaborLogButton)

    # Edit Labor Log Entry (unchanged internals, restyled buttons)
    $editLaborLogButton = & $makeActionButton "Edit Labor Log Entry" ([System.Drawing.Color]::FromArgb(52,152,219))
    $editLaborLogButton.Add_Click({
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $item = $selectedItems[0]
            $editLaborLogForm = New-Object System.Windows.Forms.Form
            $editLaborLogForm.Text = "Edit Labor Log Entry"
            $editLaborLogForm.Size = New-Object System.Drawing.Size(400, 420)
            $editLaborLogForm.StartPosition = 'CenterParent'
            $editLaborLogForm.FormBorderStyle = 'FixedDialog'
            $editLaborLogForm.MaximizeBox = $false
            $editLaborLogForm.MinimizeBox = $false
            $editLaborLogForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
            $editLaborLogForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            $labelDate = New-Object System.Windows.Forms.Label; $labelDate.Location = New-Object System.Drawing.Point(10, 20); $labelDate.Size = New-Object System.Drawing.Size(100, 20); $labelDate.Text = "Date:"; $editLaborLogForm.Controls.Add($labelDate)
            $textBoxDate = New-Object System.Windows.Forms.TextBox; $textBoxDate.Location = New-Object System.Drawing.Point(120, 20); $textBoxDate.Size = New-Object System.Drawing.Size(250, 20); $textBoxDate.Text = $item.SubItems[0].Text; $editLaborLogForm.Controls.Add($textBoxDate)

            $labelWorkOrder = New-Object System.Windows.Forms.Label; $labelWorkOrder.Location = New-Object System.Drawing.Point(10, 50); $labelWorkOrder.Size = New-Object System.Drawing.Size(100, 20); $labelWorkOrder.Text = "Work Order #:"; $editLaborLogForm.Controls.Add($labelWorkOrder)
            $textBoxWorkOrder = New-Object System.Windows.Forms.TextBox; $textBoxWorkOrder.Location = New-Object System.Drawing.Point(120, 50); $textBoxWorkOrder.Size = New-Object System.Drawing.Size(250, 20); $textBoxWorkOrder.Text = $item.SubItems[1].Text; $editLaborLogForm.Controls.Add($textBoxWorkOrder)

            $labelTask = New-Object System.Windows.Forms.Label; $labelTask.Location = New-Object System.Drawing.Point(10, 80); $labelTask.Size = New-Object System.Drawing.Size(100, 20); $labelTask.Text = "Task:"; $editLaborLogForm.Controls.Add($labelTask)
            $textBoxTask = New-Object System.Windows.Forms.TextBox; $textBoxTask.Location = New-Object System.Drawing.Point(120, 80); $textBoxTask.Size = New-Object System.Drawing.Size(250, 60); $textBoxTask.Multiline = $true; $textBoxTask.Text = $item.SubItems[2].Text; $editLaborLogForm.Controls.Add($textBoxTask)

            $labelMachineId = New-Object System.Windows.Forms.Label; $labelMachineId.Location = New-Object System.Drawing.Point(10, 150); $labelMachineId.Size = New-Object System.Drawing.Size(100, 20); $labelMachineId.Text = "Machine ID:"; $editLaborLogForm.Controls.Add($labelMachineId)
            $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox; $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 150); $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20); $editLaborLogForm.Controls.Add($comboBoxMachineId); Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"; $comboBoxMachineId.Text = $item.SubItems[3].Text

            $labelDuration = New-Object System.Windows.Forms.Label; $labelDuration.Location = New-Object System.Drawing.Point(10, 180); $labelDuration.Size = New-Object System.Drawing.Size(100, 20); $labelDuration.Text = "Duration:"; $editLaborLogForm.Controls.Add($labelDuration)
            $textBoxDuration = New-Object System.Windows.Forms.TextBox; $textBoxDuration.Location = New-Object System.Drawing.Point(120, 180); $textBoxDuration.Size = New-Object System.Drawing.Size(250, 20); $textBoxDuration.Text = $item.SubItems[4].Text; $editLaborLogForm.Controls.Add($textBoxDuration)

            $labelNotes = New-Object System.Windows.Forms.Label; $labelNotes.Location = New-Object System.Drawing.Point(10, 210); $labelNotes.Size = New-Object System.Drawing.Size(100, 20); $labelNotes.Text = "Notes:"; $editLaborLogForm.Controls.Add($labelNotes)
            $textBoxNotes = New-Object System.Windows.Forms.TextBox; $textBoxNotes.Location = New-Object System.Drawing.Point(120, 210); $textBoxNotes.Size = New-Object System.Drawing.Size(250, 60); $textBoxNotes.Multiline = $true; $textBoxNotes.Text = $item.SubItems[6].Text; $editLaborLogForm.Controls.Add($textBoxNotes)

            $saveButton = New-Object System.Windows.Forms.Button
            $saveButton.Location = New-Object System.Drawing.Point(150, 320)
            $saveButton.Size = New-Object System.Drawing.Size(100, 30)
            $saveButton.Text = "Save"
            $saveButton.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
            $saveButton.ForeColor = [System.Drawing.Color]::White
            $saveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $saveButton.FlatAppearance.BorderSize = 0
            $saveButton.Add_Click({
                $partsToAdd = @()

                foreach ($item in $selectedListView.Items) {
                    $partsToAdd += $item.Tag
                }

                if ($partsToAdd.Count -gt 0) {
                    if (Save-PartsToWorkOrder -WorkOrderNumber $workOrderNumber -Parts $partsToAdd) {
                        $form.Close()
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "No parts selected to add to the work order.",
                        "Warning",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            })
            $editLaborLogForm.Controls.Add($saveButton)
            $editLaborLogForm.ShowDialog()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select an entry to edit.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $buttonBar.Controls.Add($editLaborLogButton)

    # Refresh button
    $refreshButton = & $makeActionButton "Refresh Labor Logs" ([System.Drawing.Color]::FromArgb(52,152,219))
    $refreshButton.Add_Click({
        $script:listViewLaborLog.Items.Clear()
        Load-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath
        Process-HistoricalLogs
        Write-Log "Labor Logs manually refreshed"
    })
    $buttonBar.Controls.Add($refreshButton)

    # Add Parts to Work Order
    $addPartsButton = & $makeActionButton "Add Parts to Work Order" ([System.Drawing.Color]::FromArgb(39,174,96))
    $addPartsButton.Add_Click({
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $workOrderItem = $selectedItems[0]
            $workOrderNumber = $workOrderItem.SubItems[1].Text
            Add-PartsToWorkOrder -workOrderNumber $workOrderNumber
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a work order to add parts.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $buttonBar.Controls.Add($addPartsButton)

    Write-Log "Labor Log tab setup completed."
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

# Combo box loader
function Load-ComboBoxData {
    param (
        [System.Windows.Forms.ComboBox]$comboBox,
        [string]$csvName
    )
    Write-Log "Starting Load-ComboBoxData for $csvName"
    $csvPath = Join-Path $config.DropdownCsvsDirectory "$csvName.csv"
    Write-Log "Attempting to load data from: $csvPath"
    
    if (-not (Test-Path $csvPath)) {
        Write-Log "CSV file not found: $csvPath"
        return
    }
    
    $data = Import-Csv -Path $csvPath
    Write-Log "Imported CSV data. Row count: $($data.Count)"

    if ($null -eq $comboBox) {
        Write-Log "Error: ComboBox is null for $csvName"
        return
    }
    $comboBox.Items.Clear()
    
    if ($csvName -eq "Machines") {
        Write-Log "Processing Machines data"
        $data | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_.'Machine Acronym') -and -not [string]::IsNullOrWhiteSpace($_.'Machine Number')) {
                $machineId = "$($_.'Machine Acronym') - $($_.'Machine Number')"
                $comboBox.Items.Add($machineId)
                Write-Log "Added machine to ComboBox: $machineId"
            } else {
                Write-Log "Skipped invalid machine entry"
            }
        }
    } else {
        $data | ForEach-Object { 
            if (-not [string]::IsNullOrWhiteSpace($_.Value)) {
                $comboBox.Items.Add($_.Value)
            }
        }
    }
    
    Write-Log "Loaded $($comboBox.Items.Count) items into $csvName ComboBox"
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
                $config | ConvertTo-Json | Set-Content -Path $script:configPath
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
        $config | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath

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

    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $script:configPath
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

    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $script:configPath
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

# Function to handle Labor Log entries dynamically
function Create-LaborLogEntry {
    param (
        $callLogDate,
        $machineId,
        $taskDescription,
        $timeDown,
        $timeUp
    )

    $duration = Get-TimeDifference -startTime $timeDown -endTime $timeUp
    $durationHours = [Math]::Round($duration / 60, 2)

    $laborLogItem = New-Object System.Windows.Forms.ListViewItem($callLogDate)| Out-Null
    $laborLogItem.SubItems.Add("Need W/O #")
    $laborLogItem.SubItems.Add("Maintenance Call - Details Needed")
    $laborLogItem.SubItems.Add($machineId)
    $laborLogItem.SubItems.Add($durationHours.ToString("F2"))
    $laborLogItem.SubItems.Add("Automatically added from Call Log")

    return $laborLogItem
}

$script:listViewLaborLog = New-Object System.Windows.Forms.ListView



# Ensure RootDirectory is set
if (-not $config.RootDirectory) {
    $config.RootDirectory = $PSScriptRoot
}

# Ensure all required paths are set
$requiredPaths = @('RootDirectory', 'LaborDirectory', 'CallLogsDirectory', 'PartsRoomDirectory', 'DropdownCsvsDirectory', 'PartsBooksDirectory')
foreach ($path in $requiredPaths) {
    if (-not $config.$path) {
        Write-Log "Error: $path is not set in the configuration"
        throw "$path is missing from the configuration"
    }
}

$script:unacknowledgedEntries = @{}



# Define and set default paths if not specified in config
$defaultPaths = @{
    LaborDirectory = "Labor"
    CallLogsDirectory = "Labor"
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

# Define file paths
$callLogsFilePath = Join-Path $config.CallLogsDirectory "CallLogs.csv"
$laborLogsFilePath = Join-Path $config.LaborDirectory "LaborLogs.csv"

Write-Log "CallLogsFilePath: $callLogsFilePath"
Write-Log "LaborLogsFilePath: $laborLogsFilePath"

# Check and create required directories
$requiredDirs = @($config.LaborDirectory, $config.CallLogsDirectory, $config.PartsRoomDirectory, $config.PartsBooksDirectory, $config.DropdownCsvsDirectory)
foreach ($dir in $requiredDirs) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force
        Write-Log "Created directory: $dir"
    }
}

# Save updated config if changes were made
if ($configUpdated) {
    $config | ConvertTo-Json | Set-Content -Path $configPath
    Write-Log "Config file updated with default paths"
}

################################################################################
#                           Main Execution                                     #
################################################################################

# Function to create and show the main form
function Show-MainForm {
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

    # ---- Call Logs Tab ----
    $callLogsTab = New-Object System.Windows.Forms.TabPage
    $callLogsTab.Text = "Call Logs"
    $callLogsTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $callLogsTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($callLogsTab)

    Setup-CallLogsTab -parentTab $callLogsTab

    # ---- Labor Log Tab ----
    $laborLogTab = New-Object System.Windows.Forms.TabPage
    $laborLogTab.Text = "Labor Log"
    $laborLogTab.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $laborLogTab.Padding = New-Object System.Windows.Forms.Padding(12)
    $tabControl.TabPages.Add($laborLogTab)

    Setup-LaborLogTab -parentTab $laborLogTab -tabControl $tabControl

    Process-HistoricalLogs

    # Form closing handler
    $form.Add_FormClosing({
        $script:listViewCallLogs = $callLogsTab.Controls | Where-Object { $_ -is [System.Windows.Forms.ListView] }
        if ($script:listViewCallLogs) {
            Save-Logs -listView $script:listViewCallLogs -filePath $callLogsFilePath
        }

        $listViewLaborLog = $laborLogTab.Controls | Where-Object { $_ -is [System.Windows.Forms.ListView] }
        if ($listViewLaborLog) {
            Save-LaborLogs -listView $listViewLaborLog -filePath $laborLogsFilePath
        }
    })

    # ---- Open Parts Room ----
    $openPartsRoomButton = New-Button "Open Parts Room" {
        $partsRoomFilePath = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.xlsx" | Select-Object -First 1 -ExpandProperty FullName
        if (Test-Path $partsRoomFilePath) {
            Start-Process $partsRoomFilePath
        } else {
            [System.Windows.Forms.MessageBox]::Show("Parts Room file not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } -Style 'Primary'
    $partsBookPanel.Controls.Add($openPartsRoomButton)

    # ---- Create Parts Book ----
    $createPartsBookButton = New-Button "Create Parts Book" {
        Write-Log "Creating Parts Book..."
        $scriptPath = Join-Path $PSScriptRoot "Parts-Books-Creator.ps1"
        Write-Log "Script path: $scriptPath"
        if (Test-Path $scriptPath) {
            Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"$scriptPath`""
            Write-Log "Started process to execute Parts-Books-Creator.ps1"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Parts Books Creator script not found at $scriptPath.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Parts Books Creator script not found at $scriptPath."
        }
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
