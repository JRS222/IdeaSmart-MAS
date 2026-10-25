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
#                   Responsive Layout Engine (PowerShell 5.1)                  #
################################################################################
# The main window rescales every tab proportionally when the user resizes the
# window: absolutely-positioned controls and their fonts scale with the client
# area, docked/filling elements keep adapting absolutely, toolbar strips keep a
# fixed size (and scroll horizontally when needed), and any tab content that
# overflows gets a scrollbar. This lets the window run anywhere from roughly
# 2.6" x 3.9" up to full screen.
#
# IMPORTANT (PowerShell 5.1): inside a .GetNewClosure() block, $script: refers
# to the CLOSURE's own (empty) script scope, not the main script scope. The
# event handlers below therefore capture everything they need as locals
# ($state, $Form, $timer), and every engine function takes the state object
# explicitly instead of relying on scope.

$script:responsiveLayout = $null

function Initialize-ResponsiveLayout {
    param([System.Windows.Forms.Form]$Form)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 90

    $state = [PSCustomObject]@{
        Form          = $Form
        DesignW       = $Form.Width
        DesignH       = $Form.Height
        MinScale      = 0.15
        MaxScale      = 4.0
        LastSx        = 1.0
        LastSy        = 1.0
        LastFontScale = 1.0
        Base          = @{}
        Timer         = $timer
        Busy          = $false
        Dirty         = $true
    }
    $script:responsiveLayout = $state

    # Debounced apply: the resize event restarts the timer; the tick does the work.
    # The design size is captured from the form's size right here: a freshly
    # created WinForms window does not raise a Resize event for its initial
    # size, so an event-based capture would keep the wrong baseline.
    $timer.Add_Tick({
        if ($null -eq $state -or $state.Busy) { return }
        $state.Busy = $true
        try { Invoke-ResponsiveRescale -State $state -Force } catch { }
        $state.Busy = $false
        $timer.Stop()
    }.GetNewClosure())

    $Form.Add_Resize({
        if ($null -eq $state -or $state.Busy) { return }
        try { $timer.Stop(); $timer.Start() } catch { }
    }.GetNewClosure())
}

function Invoke-ResponsiveRefresh {
    # Call after dynamically building tab content so new controls adopt the
    # current window scale immediately. (Runs in the main script scope.)
    $rl = $script:responsiveLayout
    if (-not $rl -or $rl.DesignW -le 0) { return }
    $rl.Dirty = $true
    Invoke-ResponsiveRescale -State $rl -Force
}

function Invoke-ResponsiveRescale {
    param([Parameter(Mandatory)]$State, [switch]$Force)

    if ($null -eq $State.Form -or $State.Form.IsDisposed) { return }
    if ($State.DesignW -le 0 -or $State.DesignH -le 0) { return }

    $cw = $State.Form.Width
    $ch = $State.Form.Height
    if ($cw -le 0 -or $ch -le 0) { return }

    $sx = [Math]::Max($State.MinScale, [Math]::Min($State.MaxScale, [double]$cw / $State.DesignW))
    $sy = [Math]::Max($State.MinScale, [Math]::Min($State.MaxScale, [double]$ch / $State.DesignH))
    $fs = [Math]::Max($State.MinScale, [Math]::Min($State.MaxScale, [Math]::Min($sx, $sy)))

    if (-not $Force -and -not $State.Dirty -and
        [Math]::Abs($sx - $State.LastSx) -lt 0.0005 -and [Math]::Abs($sy - $State.LastSy) -lt 0.0005) { return }

    $State.Busy = $true
    try {
        Scale-ResponsiveControlTree -State $State -Collection $State.Form.Controls -Sx $sx -Sy $sy -FontScale $fs
        Rescale-ResponsiveSplitters -Collection $State.Form.Controls -OldSx $State.LastSx -NewSx $sx -OldSy $State.LastSy -NewSy $sy
        $State.LastSx        = $sx
        $State.LastSy        = $sy
        $State.LastFontScale = $fs
        $State.Dirty         = $false
    } catch {
        try { Write-Log "Responsive layout error: $($_.Exception.Message)" } catch { }
    } finally {
        $State.Busy = $false
    }
}

function Get-ResponsiveBase {
    # Remember a control's design-time bounds/font size the first time we see
    # it, so every rescale is computed from the ORIGINAL layout (no drift).
    param($State, $Key, [System.Drawing.Rectangle]$Bounds, [float]$FontSize)

    if (-not $State.Base.ContainsKey($Key)) {
        $State.Base[$Key] = [PSCustomObject]@{
            Bounds   = $Bounds
            FontSize = $FontSize
        }
    }
}

function Scale-ResponsiveControlTree {
    param($State, $Collection, [double]$Sx, [double]$Sy, [double]$FontScale)

    foreach ($c in @($Collection)) {
        if ($null -eq $c) { continue }
        if ($c -is [System.Windows.Forms.Form]) { continue }

        $dock = [System.Windows.Forms.DockStyle]::None
        try { $dock = $c.Dock } catch { }

        # Horizontal toolbar strips keep their own absolute layout (they scroll).
        # Children of vertical (TopDown) flow panels — settings groups, action
        # buttons — scale with everything else so they never overflow their
        # container at any window size.
        $parentIsToolbar = ($null -ne $c.Parent) -and
            ($c.Parent -is [System.Windows.Forms.FlowLayoutPanel]) -and
            ($c.Parent.FlowDirection -eq [System.Windows.Forms.FlowDirection]::LeftToRight) -and
            (-not $c.Parent.WrapContents)

        # ---- Absolutely positioned controls: move + resize proportionally ----
        if ($dock -eq [System.Windows.Forms.DockStyle]::None -and -not $parentIsToolbar) {
            $fsNow = if ($c.Font) { [float]($c.Font.SizeInPoints) } else { 9.0 }
            Get-ResponsiveBase -State $State -Key $c -Bounds $c.Bounds -FontSize $fsNow
            $b  = $State.Base[$c].Bounds
            $x  = [int][Math]::Round($b.X * $Sx)
            $y  = [int][Math]::Round($b.Y * $Sy)
            $w  = [int][Math]::Max(1, [Math]::Round($b.Width  * $Sx))
            $h  = [int][Math]::Max(1, [Math]::Round($b.Height * $Sy))
            if ($c.Left -ne $x -or $c.Top -ne $y -or $c.Width -ne $w -or $c.Height -ne $h) {
                try { $c.SetBounds($x, $y, $w, $h) } catch { }
            }
        }
        elseif ($dock -eq [System.Windows.Forms.DockStyle]::Top -or $dock -eq [System.Windows.Forms.DockStyle]::Bottom) {
            # Docked header/footer bars: the dock owns the width; scale the height.
            if (-not ($c -is [System.Windows.Forms.FlowLayoutPanel]) -and -not ($c -is [System.Windows.Forms.TableLayoutPanel])) {
                $fsNow = if ($c.Font) { [float]($c.Font.SizeInPoints) } else { 9.0 }
                Get-ResponsiveBase -State $State -Key $c -Bounds $c.Bounds -FontSize $fsNow
                $bh = $State.Base[$c].Bounds.Height
                $nh = [int][Math]::Max(1, [Math]::Round($bh * $Sy))
                if ($c.Height -ne $nh) { try { $c.Height = $nh } catch { } }
            }
        }

        # ---- ListView columns scale with the horizontal factor ----
        if ($c -is [System.Windows.Forms.ListView]) {
            foreach ($col in @($c.Columns)) {
                Get-ResponsiveBase -State $State -Key $col -Bounds (New-Object System.Drawing.Rectangle(0, 0, $col.Width, 0)) -FontSize 9.0
                $bw = $State.Base[$col].Bounds.Width
                $nw = [int][Math]::Max(12, [Math]::Round($bw * $Sx))
                if ($col.Width -ne $nw) { try { $col.Width = $nw } catch { } }
            }
        }

        # ---- TableLayoutPanel absolute column/row styles scale with the grid ----
        if ($c -is [System.Windows.Forms.TableLayoutPanel]) {
            foreach ($cs in @($c.ColumnStyles)) {
                if ($cs.SizeType -eq [System.Windows.Forms.SizeType]::Absolute) {
                    Get-ResponsiveBase -State $State -Key $cs -Bounds (New-Object System.Drawing.Rectangle(0, 0, [int]$cs.Width, 0)) -FontSize 9.0
                    $bw = $State.Base[$cs].Bounds.Width
                    $nw = [float][Math]::Max(12, [Math]::Round($bw * $Sx))
                    if ([Math]::Abs([float]$cs.Width - $nw) -ge 1) { try { $cs.Width = $nw } catch { } }
                }
            }
            foreach ($rs in @($c.RowStyles)) {
                if ($rs.SizeType -eq [System.Windows.Forms.SizeType]::Absolute) {
                    Get-ResponsiveBase -State $State -Key $rs -Bounds (New-Object System.Drawing.Rectangle(0, 0, 0, [int]$rs.Height)) -FontSize 9.0
                    $bh = $State.Base[$rs].Bounds.Height
                    $nh = [float][Math]::Max(12, [Math]::Round($bh * $Sy))
                    if ([Math]::Abs([float]$rs.Height - $nh) -ge 1) { try { $rs.Height = $nh } catch { } }
                }
            }
        }

        # ---- Fonts scale with the smallest axis factor ----
        if (-not $parentIsToolbar -and $null -ne $c.Font) {
            $fsNow = [float]($c.Font.SizeInPoints)
            Get-ResponsiveBase -State $State -Key $c -Bounds $c.Bounds -FontSize $fsNow
            $baseSize = $State.Base[$c].FontSize
            $newSize  = [float][Math]::Max(3.0, $baseSize * $FontScale)
            if ([Math]::Abs($newSize - [float]($c.Font.SizeInPoints)) -ge 0.15) {
                try {
                    # Do not dispose the old font: some controls share inherited
                    # font instances, and Font has a finalizer anyway.
                    $newFont = New-Object System.Drawing.Font($c.Font.FontFamily, $newSize, $c.Font.Style)
                    $c.Font  = $newFont
                } catch { }
            }
        }

        # ---- Recurse ----
        if ($c -is [System.Windows.Forms.TabControl]) {
            foreach ($page in @($c.TabPages)) {
                Scale-ResponsiveControlTree -State $State -Collection $page.Controls -Sx $Sx -Sy $Sy -FontScale $FontScale
            }
        }
        elseif ($c.Controls -and $c.Controls.Count -gt 0) {
            Scale-ResponsiveControlTree -State $State -Collection $c.Controls -Sx $Sx -Sy $Sy -FontScale $FontScale
        }
    }
}

function Rescale-ResponsiveSplitters {
    # Splitter distances are scaled DELTA-style so the user's own splitter
    # drags survive window resizes.
    param($Collection, [double]$OldSx, [double]$NewSx, [double]$OldSy, [double]$NewSy)

    foreach ($c in @($Collection)) {
        if ($null -eq $c) { continue }
        if ($c -is [System.Windows.Forms.SplitContainer] -and $OldSx -gt 0 -and $NewSx -gt 0 -and $OldSy -gt 0 -and $NewSy -gt 0) {
            try {
                if ($c.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) {
                    $d = [int][Math]::Round($c.SplitterDistance * ($NewSx / $OldSx))
                } else {
                    $d = [int][Math]::Round($c.SplitterDistance * ($NewSy / $OldSy))
                }
                # Panel minimum sizes scale too, so neither panel can be
                # crushed out of existence at small window sizes.
                $m1 = [int][Math]::Max(20, [Math]::Min(400, [Math]::Round($c.Panel1MinSize * ($NewSy / $OldSy))))
                $m2 = [int][Math]::Max(20, [Math]::Min(400, [Math]::Round($c.Panel2MinSize * ($NewSy / $OldSy))))
                if ($c.Panel1MinSize -ne $m1) { try { $c.Panel1MinSize = $m1 } catch { } }
                if ($c.Panel2MinSize -ne $m2) { try { $c.Panel2MinSize = $m2 } catch { } }
                if ($d -lt $m1) { $d = $m1 }
                $maxD = if ($c.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) { $c.Width - $m2 } else { $c.Height - $m2 }
                if ($d -gt $maxD) { $d = $maxD }
                if ($d -lt $m1) { $d = $m1 }
                $c.SplitterDistance = $d
            } catch { }
        }
        if ($c.Controls -and $c.Controls.Count -gt 0) {
            Rescale-ResponsiveSplitters -Collection $c.Controls -OldSx $OldSx -NewSx $NewSx -OldSy $OldSy -NewSy $NewSy
        }
    }
}

################################################################################
#                          UI Theme Engine (PowerShell 5.1)                    #
################################################################################
# Three themes: Default (as designed), Dark, and Postal (USPS-style).
# Colors are swapped by exact-RGB mapping, so no existing layout code changes.
# BackColor and ForeColor are mapped independently (e.g., white button text
# stays white in every theme).

$script:theme = 'Default'
# Remembers every control's DESIGN-TIME colors so themes always translate from
# the original palette (switching Dark -> XP -> Default never compounds).
$script:themeBase = @{}

$script:themeBackMaps = @{
    Dark = @{
        '236,240,241' = [System.Drawing.Color]::FromArgb(30, 30, 30)     # window bg
        '245,247,250' = [System.Drawing.Color]::FromArgb(37, 37, 38)     # panels / pages
        '226,232,240' = [System.Drawing.Color]::FromArgb(60, 60, 62)     # progress track / status bars
        '255,255,255' = [System.Drawing.Color]::FromArgb(45, 45, 48)     # white surfaces (lists, textboxes)
        '220,225,232' = [System.Drawing.Color]::FromArgb(60, 60, 62)     # splitter back
        '189,195,199' = [System.Drawing.Color]::FromArgb(66, 66, 68)     # secondary / ghost buttons
        '236,240,242' = [System.Drawing.Color]::FromArgb(66, 66, 68)     # ghost buttons alt
        '240,240,240' = [System.Drawing.Color]::FromArgb(45, 45, 48)     # Control gray (tab strip area, group boxes)
    }
    Postal = @{
        '236,240,241' = [System.Drawing.Color]::FromArgb(232, 236, 242)  # light postal gray-blue window
        '245,247,250' = [System.Drawing.Color]::FromArgb(240, 243, 247)  # light postal panels
        '226,232,240' = [System.Drawing.Color]::FromArgb(214, 222, 232)  # track / status bars
        '255,255,255' = [System.Drawing.Color]::FromArgb(255, 255, 255)  # white surfaces stay white
        '220,225,232' = [System.Drawing.Color]::FromArgb(196, 206, 220)  # splitter back
        '189,195,199' = [System.Drawing.Color]::FromArgb(163, 178, 196)  # secondary buttons
        '236,240,242' = [System.Drawing.Color]::FromArgb(163, 178, 196)  # ghost buttons alt
        '240,240,240' = [System.Drawing.Color]::FromArgb(232, 236, 242)  # Control gray (tab strip area, group boxes)
        '52,152,219'  = [System.Drawing.Color]::FromArgb(0, 75, 135)     # Postal Blue (#004B87)
        '39,174,96'   = [System.Drawing.Color]::FromArgb(39, 174, 96)    # green kept
        '231,76,60'   = [System.Drawing.Color]::FromArgb(218, 41, 28)    # USPS red (#DA291C)
    }
}

$script:themeForeMaps = @{
    Dark = @{
        '44,62,80'    = [System.Drawing.Color]::FromArgb(225, 228, 232)  # dark text -> light
        '90,100,115'  = [System.Drawing.Color]::FromArgb(150, 156, 163)  # muted text
        '120,130,145' = [System.Drawing.Color]::FromArgb(150, 156, 163)  # muted text
        '160,170,185' = [System.Drawing.Color]::FromArgb(118, 124, 131)  # faint hint text
        '179,38,30'   = [System.Drawing.Color]::FromArgb(255, 130, 120)  # error text
        '0,0,0'       = [System.Drawing.Color]::FromArgb(230, 230, 230)  # black text
    }
    Postal = @{
        '44,62,80'    = [System.Drawing.Color]::FromArgb(27, 43, 64)     # postal navy text
        '90,100,115'  = [System.Drawing.Color]::FromArgb(96, 110, 128)
        '120,130,145' = [System.Drawing.Color]::FromArgb(96, 110, 128)
        '160,170,185' = [System.Drawing.Color]::FromArgb(130, 140, 155)
        '179,38,30'   = [System.Drawing.Color]::FromArgb(218, 41, 28)    # USPS red
    }
}

function Get-ThemeMappedColor {
    param([System.Drawing.Color]$Color, $Map)
    if ($null -eq $Color -or $null -eq $Map) { return $null }
    $key = "{0},{1},{2}" -f $Color.R, $Color.G, $Color.B
    if ($Map.ContainsKey($key)) { return $Map[$key] }
    return $null
}

################################################################################
# Owner-draw handlers: tab headers and ListView headers/rows are drawn by the
# OS, so they can only follow a theme when we draw them ourselves.
################################################################################

$script:themeTabPaint = {
    # TabControl.BackColor is ignored by the OS control in 5.1, so the strip
    # area behind the tabs is filled here instead.
    param($sender, $e)
    try {
        if ($script:theme -eq 'Default') { return }
        $c = if ($script:theme -eq 'Dark') { [System.Drawing.Color]::FromArgb(37, 37, 38) }
             else { [System.Drawing.Color]::FromArgb(232, 236, 242) }
        $b = New-Object System.Drawing.SolidBrush($c)
        $e.Graphics.FillRectangle($b, $sender.ClientRectangle)
        $b.Dispose()
    } catch { }
}

$script:themeTabDraw = {
    param($sender, $e)
    try {
        $g = $e.Graphics
        if ($e.Index -lt 0 -or $e.Index -ge $sender.TabPages.Count) { return }
        $page = $sender.TabPages[$e.Index]
        $r = $e.Bounds
        $selected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0
        switch ($script:theme) {
            'Dark' {
                $back = if ($selected) { [System.Drawing.Color]::FromArgb(70, 70, 74) } else { [System.Drawing.Color]::FromArgb(45, 45, 48) }
                $txt  = [System.Drawing.Color]::FromArgb(225, 228, 232)
                $edge = [System.Drawing.Color]::FromArgb(22, 22, 24)
                $acc  = [System.Drawing.Color]::FromArgb(52, 152, 219)
            }
            'Postal' {
                $back = if ($selected) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::FromArgb(214, 222, 232) }
                $txt  = [System.Drawing.Color]::FromArgb(27, 43, 64)
                $edge = [System.Drawing.Color]::FromArgb(163, 178, 196)
                $acc  = [System.Drawing.Color]::FromArgb(0, 75, 135)
            }
            default { return }
        }
        $brush = New-Object System.Drawing.SolidBrush($back)
        $pen   = New-Object System.Drawing.Pen($edge)
        try {
            $g.FillRectangle($brush, $r)
            if ($selected) {
                $ab = New-Object System.Drawing.SolidBrush($acc)
                $g.FillRectangle($ab, $r.X, $r.Y, $r.Width, 3)
                $ab.Dispose()
                $g.DrawLine($pen, $r.X, $r.Y + 1, $r.X, $r.Bottom - 1)
                $g.DrawLine($pen, $r.Right - 1, $r.Y + 1, $r.Right - 1, $r.Bottom - 1)
            }
            $fmt = New-Object System.Drawing.StringFormat
            $fmt.Alignment = [System.Drawing.StringAlignment]::Center
            $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
            $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
            $textBrush = New-Object System.Drawing.SolidBrush($txt)
            # DrawString needs a RectangleF in PowerShell 5.1 (a Rectangle fails to bind)
            $rf = New-Object System.Drawing.RectangleF($r.X, $r.Y, $r.Width, $r.Height)
            $g.DrawString($page.Text, $sender.Font, $textBrush, $rf, $fmt)
            $textBrush.Dispose(); $fmt.Dispose()
        } finally {
            $brush.Dispose(); $pen.Dispose()
        }
    } catch { }
}

$script:themeLvHeaderDraw = {
    param($sender, $e)
    try {
        $g = $e.Graphics
        $col = $sender.Columns[$e.ColumnIndex]
        $r = $e.Bounds
        switch ($script:theme) {
            'Dark'   { $back = [System.Drawing.Color]::FromArgb(58, 58, 60);  $txt = [System.Drawing.Color]::FromArgb(205, 209, 214); $line = [System.Drawing.Color]::FromArgb(88, 88, 92) }
            'Postal' { $back = [System.Drawing.Color]::FromArgb(0, 75, 135);   $txt = [System.Drawing.Color]::White;                  $line = [System.Drawing.Color]::FromArgb(0, 56, 102) }
            default  { $back = [System.Drawing.Color]::FromArgb(240, 242, 245); $txt = [System.Drawing.Color]::FromArgb(44, 62, 80);   $line = [System.Drawing.Color]::FromArgb(200, 205, 212) }
        }
        $brush = New-Object System.Drawing.SolidBrush($back)
        $g.FillRectangle($brush, $r)
        $brush.Dispose()
        $tb = New-Object System.Drawing.SolidBrush($txt)
        $fmt = New-Object System.Drawing.StringFormat
        $fmt.Alignment = [System.Drawing.StringAlignment]::Center
        $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
        $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
        # DrawString needs a RectangleF in PowerShell 5.1 (a Rectangle fails to bind)
        $rf = New-Object System.Drawing.RectangleF($r.X, $r.Y, $r.Width, $r.Height)
        $g.DrawString($col.Text, $sender.Font, $tb, $rf, $fmt)
        $tb.Dispose(); $fmt.Dispose()
        $pen = New-Object System.Drawing.Pen($line)
        $g.DrawLine($pen, $r.Right - 1, $r.Top, $r.Right - 1, $r.Bottom - 1)
        $g.DrawLine($pen, $r.Left, $r.Bottom - 1, $r.Right - 1, $r.Bottom - 1)
        $pen.Dispose()
    } catch { }
}

$script:themeLvItemDraw = {
    param($sender, $e)
    try {
        $g = $e.Graphics
        $r = $e.Bounds
        $item = $e.Item
        $grid = if ($script:theme -eq 'Dark') { [System.Drawing.Color]::FromArgb(70, 70, 72) }
                elseif ($script:theme -eq 'Postal') { [System.Drawing.Color]::FromArgb(214, 222, 232) }
                else { [System.Drawing.Color]::FromArgb(200, 205, 212) }
        if ($item.Selected) {
            $sel = if ($script:theme -eq 'Dark') { [System.Drawing.Color]::FromArgb(70, 96, 130) }
                   elseif ($script:theme -eq 'Postal') { [System.Drawing.Color]::FromArgb(0, 105, 180) }
                   else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
            $b = New-Object System.Drawing.SolidBrush($sel)
            $g.FillRectangle($b, $r)
            $b.Dispose()
        } else {
            $back = $item.BackColor
            if ($back.ToArgb() -eq 0) { $back = $sender.BackColor }
            $b = New-Object System.Drawing.SolidBrush($back)
            $g.FillRectangle($b, $r)
            $b.Dispose()
        }
        # checkbox for checkable lists (PM task lists)
        if ($sender.CheckBoxes) {
            $cb = New-Object System.Drawing.Rectangle(($r.X + 3), ($r.Y + [int](($r.Height - 13) / 2)), 13, 13)
            $boxCol = if ($script:theme -eq 'Dark') { [System.Drawing.Color]::FromArgb(200, 202, 206) } else { [System.Drawing.Color]::White }
            $wb = New-Object System.Drawing.SolidBrush($boxCol)
            $g.FillRectangle($wb, $cb)
            $wb.Dispose()
            $cp = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(120, 124, 130))
            $g.DrawRectangle($cp, $cb)
            $cp.Dispose()
            if ($item.Checked) {
                $xp = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 120, 60), 2)
                $g.DrawLine($xp, ($cb.X + 2), ($cb.Y + 7), ($cb.X + 5), ($cb.Y + 10))
                $g.DrawLine($xp, ($cb.X + 5), ($cb.Y + 10), ($cb.X + 11), ($cb.Y + 3))
                $xp.Dispose()
            }
        }
        $gp = New-Object System.Drawing.Pen($grid)
        $g.DrawLine($gp, $r.Left, $r.Bottom - 1, $r.Right - 1, $r.Bottom - 1)
        $gp.Dispose()
    } catch { }
}

$script:themeLvSubItemDraw = {
    param($sender, $e)
    try {
        $g = $e.Graphics
        $sub = $e.SubItem
        $r = $e.Bounds
        $txtColor = $sub.ForeColor
        if ($txtColor.ToArgb() -eq 0) { $txtColor = $sender.ForeColor }
        if ($e.Item.Selected -and $txtColor.ToArgb() -eq $sender.ForeColor.ToArgb()) {
            $txtColor = [System.Drawing.Color]::White
        }
        $fmt = New-Object System.Drawing.StringFormat
        $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
        $fmt.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
        $pad = if ($sender.CheckBoxes -and $e.ColumnIndex -eq 0) { 20 } else { 4 }
        $textRect = New-Object System.Drawing.RectangleF(($r.X + $pad), $r.Y, ($r.Width - $pad - 4), $r.Height)
        $tb = New-Object System.Drawing.SolidBrush($txtColor)
        $font = if ($sub.Font) { $sub.Font } else { $sender.Font }
        $g.DrawString($sub.Text, $font, $tb, $textRect, $fmt)
        $tb.Dispose(); $fmt.Dispose()
        $grid = if ($script:theme -eq 'Dark') { [System.Drawing.Color]::FromArgb(70, 70, 72) }
                elseif ($script:theme -eq 'Postal') { [System.Drawing.Color]::FromArgb(214, 222, 232) }
                else { [System.Drawing.Color]::FromArgb(200, 205, 212) }
        $pen = New-Object System.Drawing.Pen($grid)
        $g.DrawLine($pen, $r.Right - 1, $r.Top, $r.Right - 1, $r.Bottom - 1)
        $pen.Dispose()
    } catch { }
}

################################################################################
# Native window chrome (title bar, frame, scrollbars) — drawn by the OS, so it
# can only follow a theme through the DWM API (Windows 10 1809+ / Windows 11).
################################################################################

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ThemeNative
{
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref uint value, int size);
}
"@ -ErrorAction SilentlyContinue

function Set-ThemeWindowChrome {
    # Colors the title bar, window frame, and (via immersive dark mode) the
    # native scrollbars for one top-level window.
    param($Form)

    if ($null -eq $Form -or -not ($Form -is [System.Windows.Forms.Form])) { return }
    try {
        if (-not $Form.IsHandleCreated) { return }
    } catch { return }

    $hwnd = $Form.Handle
    try {
        if ($script:theme -eq 'Dark') {
            # Immersive dark mode: darkens the title bar, frame, AND scrollbars.
            $v = 1
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 20, [ref]$v, 4)   # 20H1+
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 19, [ref]$v, 4)   # pre-20H1
            # Windows 11 caption/border/text colors (ignored on older builds)
            $cap = [uint32]0x00262626   # COLORREF (BGR) dark charcoal
            $brd = [uint32]0x00101010
            $txt = [uint32]0x00E8E8E8
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 35, [ref]$cap, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 34, [ref]$brd, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 36, [ref]$txt, 4)
        }
        elseif ($script:theme -eq 'Postal') {
            $v = 0
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 20, [ref]$v, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 19, [ref]$v, 4)
            # Postal-blue caption + frame with white text (Windows 11)
            $cap = [uint32]0x00874B00   # BGR of RGB(0, 75, 135)
            $brd = [uint32]0x003E6FA9
            $txt = [uint32]0x00FFFFFF
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 35, [ref]$cap, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 34, [ref]$brd, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 36, [ref]$txt, 4)
        }
        else {
            # Default: back to the standard light chrome.
            $v = 0
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 20, [ref]$v, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 19, [ref]$v, 4)
            $cap = [uint32]0x00FFFFFF
            $brd = [uint32]0x00FFFFFF
            $txt = [uint32]0x00000000
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 35, [ref]$cap, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 34, [ref]$brd, 4)
            [void][ThemeNative]::DwmSetWindowAttribute($hwnd, 36, [ref]$txt, 4)
        }
    } catch { }
}

function Set-ThemedControlColors {
    # Themes ONE control (BackColor, ForeColor, button hover shades, progress
    # bar surfaces), translating from its stored design-time colors so themes
    # can be switched back and forth without compounding.
    param($c)

    if ($null -eq $c) { return }

    $isBar = $false
    try { $isBar = ($c -is [ModernProgressBar]) } catch { $isBar = $false }

    # Capture the control's original (design-time) colors on first sight.
    if (-not $script:themeBase.ContainsKey($c)) {
        $track  = if ($isBar) { $c.TrackColor }    else { [System.Drawing.Color]::Empty }
        $start  = if ($isBar) { $c.BarColorStart } else { [System.Drawing.Color]::Empty }
        $end    = if ($isBar) { $c.BarColorEnd }   else { [System.Drawing.Color]::Empty }
        $script:themeBase[$c] = [PSCustomObject]@{
            Back = $c.BackColor; Fore = $c.ForeColor
            Track = $track; BarStart = $start; BarEnd = $end
        }
    }
    $orig = $script:themeBase[$c]

    $mapBack = $script:themeBackMaps[$script:theme]
    $mapFore = $script:themeForeMaps[$script:theme]

    # BackColor: themed value when the design color has a mapping, otherwise
    # the original (this also restores the Default look).
    try {
        $nb = if ($mapBack) { Get-ThemeMappedColor -Color $orig.Back -Map $mapBack } else { $null }
        if ($null -eq $nb) { $nb = $orig.Back }
        if ($c.BackColor.ToArgb() -ne $nb.ToArgb()) { $c.BackColor = $nb }
    } catch { }
    try {
        $nf = if ($mapFore) { Get-ThemeMappedColor -Color $orig.Fore -Map $mapFore } else { $null }
        if ($null -eq $nf) { $nf = $orig.Fore }
        if ($c.ForeColor.ToArgb() -ne $nf.ToArgb()) { $c.ForeColor = $nf }
    } catch { }

    # Buttons: re-derive hover/pressed shades from the themed base color,
    # exactly like New-Button does.
    if ($c -is [System.Windows.Forms.ButtonBase]) {
        try {
            $b = $c.BackColor
            $c.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
                [Math]::Min(255, $b.R + 25), [Math]::Min(255, $b.G + 25), [Math]::Min(255, $b.B + 25))
            $c.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(
                [Math]::Max(0, $b.R - 25), [Math]::Max(0, $b.G - 25), [Math]::Max(0, $b.B - 25))
        } catch { }
    }

    # Custom progress bar surfaces
    if ($isBar) {
        try {
            $nt = if ($mapBack) { Get-ThemeMappedColor -Color $orig.Track -Map $mapBack } else { $null }
            if ($null -eq $nt) { $nt = $orig.Track }
            if ($c.TrackColor.ToArgb() -ne $nt.ToArgb()) { $c.TrackColor = $nt }
            $ns = if ($mapBack) { Get-ThemeMappedColor -Color $orig.BarStart -Map $mapBack } else { $null }
            if ($null -eq $ns) { $ns = $orig.BarStart }
            if ($c.BarColorStart.ToArgb() -ne $ns.ToArgb()) { $c.BarColorStart = $ns }
            $ne = if ($mapBack) { Get-ThemeMappedColor -Color $orig.BarEnd -Map $mapBack } else { $null }
            if ($null -eq $ne) { $ne = $orig.BarEnd }
            if ($c.BarColorEnd.ToArgb() -ne $ne.ToArgb()) { $c.BarColorEnd = $ne }
        } catch { }
    }

    # Tab strips and ListView headers/rows are OS-drawn; switch them to
    # owner-draw in themed modes so they follow the theme as well.
    if ($c -is [System.Windows.Forms.TabControl]) {
        try {
            if ($script:theme -eq 'Default') {
                if ($c.DrawMode -ne [System.Windows.Forms.TabDrawMode]::Normal) {
                    $c.DrawMode = [System.Windows.Forms.TabDrawMode]::Normal
                }
            } else {
                if ($c.DrawMode -ne [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed) {
                    $c.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
                    $c.ItemSize = New-Object System.Drawing.Size(112, [Math]::Max(24, $c.Font.Height + 14))
                }
                if (-not $script:themeBase[$c].TabWired) {
                    $c.Add_DrawItem($script:themeTabDraw)
                    $c.Add_Paint($script:themeTabPaint)
                    # Add-Member -Force: plain property assignment on a new
                    # member fails in PS 5.1 for objects held in hashtables.
                    $script:themeBase[$c] | Add-Member -NotePropertyName 'TabWired' -NotePropertyValue $true -Force
                }
            }
        } catch { }
    }
    if ($c -is [System.Windows.Forms.ListView]) {
        try {
            $want = ($script:theme -ne 'Default')
            if ($c.OwnerDraw -ne $want) { $c.OwnerDraw = $want }
            if ($want -and -not $script:themeBase[$c].LvWired) {
                $c.Add_DrawColumnHeader($script:themeLvHeaderDraw)
                $c.Add_DrawItem($script:themeLvItemDraw)
                $c.Add_DrawSubItem($script:themeLvSubItemDraw)
                $script:themeBase[$c] | Add-Member -NotePropertyName 'LvWired' -NotePropertyValue $true -Force
            }
        } catch { }
    }
}

function Apply-ThemeControlTree {
    param($Collection)

    foreach ($c in @($Collection)) {
        if ($null -eq $c) { continue }
        if ($c -is [System.Windows.Forms.Form]) { continue }

        Set-ThemedControlColors -c $c

        if ($c -is [System.Windows.Forms.TabControl]) {
            foreach ($page in @($c.TabPages)) {
                Set-ThemedControlColors -c $page   # theme the tab page surface itself
                Apply-ThemeControlTree -Collection $page.Controls
            }
        }
        elseif ($c.Controls -and $c.Controls.Count -gt 0) {
            Apply-ThemeControlTree -Collection $c.Controls
        }
    }
}

function Apply-UiTheme {
    param($Root = $null)

    if ($null -eq $Root) {
        if ($script:responsiveLayout -and $null -ne $script:responsiveLayout.Form) {
            $Root = $script:responsiveLayout.Form
        }
    }
    if ($null -eq $Root -or $Root.IsDisposed) { return }
    Set-ThemedControlColors -c $Root   # theme the window itself too
    Apply-ThemeControlTree -Collection $Root.Controls

    # Native chrome (title bar / frame / scrollbars) for top-level windows.
    if ($Root -is [System.Windows.Forms.Form]) {
        try {
            if (-not $script:themeBase[$Root].ChromeWired) {
                $Root.Add_HandleCreated({
                    Set-ThemeWindowChrome -Form $Root
                }.GetNewClosure())
                $script:themeBase[$Root] | Add-Member -NotePropertyName 'ChromeWired' -NotePropertyValue $true -Force
            }
            Set-ThemeWindowChrome -Form $Root
        } catch { }
    }

    try { $Root.Refresh() } catch { }
    Write-Log "Theme applied: $script:theme"
}

function Set-UiTheme {
    param([string]$Name)
    if ($Name -notin @('Default','Dark','Postal')) { $Name = 'Default' }
    $script:theme = $Name
    if ($script:config) {
        $script:config | Add-Member -NotePropertyName 'UiTheme' -NotePropertyValue $Name -Force
    }
    Apply-UiTheme
}

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

# ============================================================
# Automatic retrieval helpers (Sites + Parts Volumes)
# ============================================================
#
# Both the eMARS sites-criteria page and the MTSC parts-search page
# expose their dropdown data either inline (<select> in the served
# HTML) or via a lazy AJAX call.  These helpers fetch the page,
# look for the inline case, and if that fails scan the page's JS for
# candidate AJAX endpoints and try them in turn.

function Invoke-PmHttpGet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [int]$TimeoutSec = 30,
        [switch]$UseDefaultCredentials
    )

    $params = @{
        Uri             = $Url
        UseBasicParsing = $true
        TimeoutSec      = $TimeoutSec
    }
    if ($UseDefaultCredentials) { $params.UseDefaultCredentials = $true }

    Write-Log "Invoke-PmHttpGet: GET $Url"
    try {
        $resp = Invoke-WebRequest @params
    } catch {
        Write-Log "Invoke-PmHttpGet: FAILED $Url — $($_.Exception.Message)"
        throw
    }
    Write-Log "Invoke-PmHttpGet: $($resp.StatusCode) — $($resp.Content.Length) char(s)"
    return $resp.Content
}

function Invoke-PmHttpPost {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][hashtable]$Form,
        [int]$TimeoutSec = 30,
        [switch]$UseDefaultCredentials
    )

    $params = @{
        Uri             = $Url
        Method          = 'Post'
        Body            = $Form
        ContentType     = 'application/x-www-form-urlencoded'
        UseBasicParsing = $true
        TimeoutSec      = $TimeoutSec
    }
    if ($UseDefaultCredentials) { $params.UseDefaultCredentials = $true }

    Write-Log "Invoke-PmHttpPost: POST $Url"
    try {
        $resp = Invoke-WebRequest @params
    } catch {
        Write-Log "Invoke-PmHttpPost: FAILED $Url — $($_.Exception.Message)"
        throw
    }
    Write-Log "Invoke-PmHttpPost: $($resp.StatusCode) — $($resp.Content.Length) char(s)"
    return $resp.Content
}

function Resolve-RelativeUrl {
    param([string]$Base, [string]$Relative)
    if ([string]::IsNullOrWhiteSpace($Relative)) { return $null }

    $absUri = $null
    if ([System.Uri]::TryCreate($Relative, [System.UriKind]::Absolute, [ref]$absUri)) {
        return $absUri.AbsoluteUri
    }

    $baseUri = $null
    if ([System.Uri]::TryCreate($Base, [System.UriKind]::Absolute, [ref]$baseUri)) {
        $relUri = $null
        if ([System.Uri]::TryCreate($Relative, [System.UriKind]::Relative, [ref]$relUri)) {
            return ([System.Uri]::new($baseUri, $relUri)).AbsoluteUri
        }
    }
    return $Relative
}

function Save-DebugHtml {
    param([string]$Html, [string]$Tag)

    $dir = Join-Path $script:config.RootDirectory 'Debug Dumps'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $safe = ($Tag -replace '[^\w\-]', '_')
    $path = Join-Path $dir ("$safe-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.html')
    [System.IO.File]::WriteAllText($path, $Html, (New-Object System.Text.UTF8Encoding($false)))
    Write-Log "Save-DebugHtml: wrote $path"
    return $path
}

function Find-SelectBlock {
    # Returns the outer HTML of the first <select> whose opening tag or
    # body matches any of $Hints.  Returns $null if none match.
    param([string]$Html, [string[]]$Hints = @())
    if ([string]::IsNullOrWhiteSpace($Html)) { return $null }

    $blocks = [regex]::Matches($Html, '(?is)<select\b[^>]*>.*?</select>')
    if ($blocks.Count -eq 0) { return $null }

    # Pass 1: match against the opening tag (select name= / id=)
    foreach ($b in $blocks) {
        $openTag = [regex]::Match($b.Value, '(?is)^<select\b[^>]*>').Value
        foreach ($h in $Hints) {
            if ($openTag -match [regex]::Escape($h)) { return $b.Value }
        }
    }
    # Pass 2: match against the whole block (option values etc.)
    foreach ($b in $blocks) {
        foreach ($h in $Hints) {
            if ($b.Value -match [regex]::Escape($h)) { return $b.Value }
        }
    }
    return $null
}

function Find-AjaxEndpoints {
    param([string]$Html, [string]$Keyword = '')

    $found = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Html)) { return ,@() }

    $patterns = @(
        '\.(?:get|post|ajax)\s*\(\s*[''"](?<url>[^''"\s]{3,300})[''"]',
        '\burl\s*:\s*[''"](?<url>[^''"\s]{3,300})[''"]',
        '\bfetch\s*\(\s*[''"](?<url>[^''"\s]{3,300})[''"]',
        '\bopen\s*\(\s*[''"](?:GET|POST)[''"]\s*,\s*[''"](?<url>[^''"\s]{3,300})[''"]',
        '\.load\s*\(\s*[''"](?<url>[^''"\s]{3,300})[''"]'     # jQuery .load()
    )

    foreach ($p in $patterns) {
        foreach ($m in [regex]::Matches($Html, $p, 'IgnoreCase')) {
            $u = $m.Groups['url'].Value.Trim()
            if ($u.Length -lt 3) { continue }
            if ($Keyword -and ($u -notmatch [regex]::Escape($Keyword))) { continue }
            if (-not $found.Contains($u)) { [void]$found.Add($u) }
        }
    }
    return ,$found.ToArray()
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
            EmarsOrigin                = 'http://emarssu3.eng.usps.gov'
            EmarsUser                  = ''
            EmarsSite                  = ''
            EmarsSubsite               = '01'
            EmarsFacId                 = ''
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
        if ($config.WorkTracking) {
            $wtDefaults = (New-DefaultConfigObject -Root $config.RootDirectory -ScriptsDir $PSScriptRoot).WorkTracking
            foreach ($wtKey in @('EmarsOrigin', 'EmarsUser', 'EmarsSite', 'EmarsSubsite', 'EmarsFacId')) {
                if (-not ($config.WorkTracking.PSObject.Properties.Name -contains $wtKey)) {
                    $config.WorkTracking | Add-Member -NotePropertyName $wtKey -NotePropertyValue $wtDefaults.$wtKey
                    Write-Log "Backfilled missing WorkTracking.$wtKey key."
                }
            }
        }
        if (-not ($config.PSObject.Properties.Name -contains 'SupervisorEmail')) {
            $config | Add-Member -NotePropertyName 'SupervisorEmail' -NotePropertyValue 'default@example.com'
            Write-Log "Backfilled missing SupervisorEmail key."
        }
        if (-not ($config.PSObject.Properties.Name -contains 'UiTheme')) {
            $config | Add-Member -NotePropertyName 'UiTheme' -NotePropertyValue 'Default'
            Write-Log "Backfilled missing UiTheme key."
        }
        Write-Log "Config loaded successfully"
    } else {
        Write-Log "Config file not found. Using default configuration."
        $config = New-DefaultConfigObject -Root $PSScriptRoot -ScriptsDir $PSScriptRoot
    }

    return $config
}

$config = Initialize-Config
if ($config.PSObject.Properties['UiTheme'] -and "$($config.UiTheme)" -ne '') {
    $script:theme = "$($config.UiTheme)"
    # Migration: the old WindowsXP theme became the Postal theme.
    if ($script:theme -eq 'WindowsXP') { $script:theme = 'Postal' }
}
Write-Log "Active UI theme: $script:theme"

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

function Get-PartsBookSections {
    # Extracts section entries (sno + display name) from handbook tree HTML,
    # whether that HTML came from the phbk index page or from the
    # build_Tree_sections.php AJAX response.
    param([string]$Html)

    $out = @()
    if ([string]::IsNullOrWhiteSpace($Html)) { return $out }

    foreach ($m in [regex]::Matches($Html, '(?is)<li\b(?<attrs>[^>]*\bsno\s*=\s*["''](?<sno>\d+)[^>]*)>(?<body>.*?)</li>')) {
        $span = [regex]::Match($m.Groups['body'].Value, '(?is)<span\b[^>]*>(?<t>.*?)</span>')
        if (-not $span.Success) { continue }
        $name = [System.Net.WebUtility]::HtmlDecode(($span.Groups['t'].Value -replace '(?is)<[^>]+>', ''))
        $name = ($name -replace '\s+', ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $out += [PSCustomObject]@{ Number = $m.Groups['sno'].Value; Name = $name }
    }
    return $out
}

function Get-PartsBookFigures {
    # Extracts figure entries (figno + name) from a section_figures.php
    # AJAX response.
    param([string]$Html)

    $out = @()
    if ([string]::IsNullOrWhiteSpace($Html)) { return $out }

    foreach ($m in [regex]::Matches($Html, '(?is)<li\b(?<attrs>[^>]*\bfigno\s*=\s*["''](?<figno>[\w.-]+)[^>]*)>(?<body>.*?)</li>')) {
        $span = [regex]::Match($m.Groups['body'].Value, '(?is)<span\b[^>]*>(?<t>.*?)</span>')
        if (-not $span.Success) { continue }
        $name = [System.Net.WebUtility]::HtmlDecode(($span.Groups['t'].Value -replace '(?is)<[^>]+>', ''))
        $name = ($name -replace '\s+', ' ').Trim()
        $name = $name -replace '^[\w]+-\d+\s+', ''
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $out += [PSCustomObject]@{ Figno = $m.Groups['figno'].Value; Name = $name }
    }
    return $out
}

function Get-PartsBookTreeFromWeb {
    # Auto-fetches the handbook tree (sections + figures) for one book and
    # returns a synthetic HTML document shaped like the pasted page, so the
    # existing ConvertFrom-PartsBookTree / -Fallback parsers work unchanged.
    #
    # Endpoints (from the phbk viewer's own treeControl.js / main.js):
    #   GET  index.php?msbookno=N&volno=V        — shell; may render sections
    #   POST content/build_Tree_sections.php     — section <li>s  {booknum, vol}
    #   POST ajax/section_figures.php            — figure <li>s   {booknum, vol, section}
    param(
        [Parameter(Mandatory)][string]$MsBookNo,
        [Parameter(Mandatory)][string]$Volume,
        [int]$TimeoutSec = 30
    )

    $phbkBase = 'https://www1.mtsc.usps.gov/apps/phbk'

    # 1. Sections — index page first, then the AJAX tree builder as fallback.
    $sections = @()
    $sectionHtml = ''
    try {
        $page = Invoke-PmHttpGet -Url "$phbkBase/index.php?msbookno=$MsBookNo&volno=$Volume" -TimeoutSec $TimeoutSec
        $sectionHtml = $page
        $sections = @(Get-PartsBookSections -Html $page)
    } catch {
        Write-Log "Get-PartsBookTreeFromWeb: index page failed: $($_.Exception.Message)"
    }

    if ($sections.Count -eq 0) {
        $treeResp = Invoke-PmHttpPost -Url "$phbkBase/content/build_Tree_sections.php" -Form @{ booknum = $MsBookNo; vol = $Volume } -TimeoutSec $TimeoutSec
        $sectionHtml = $treeResp
        $sections = @(Get-PartsBookSections -Html $treeResp)
    }
    if ($sections.Count -eq 0) {
        $dump = Save-DebugHtml -Html $sectionHtml -Tag 'phbk-tree'
        throw "No sections found in the handbook tree for MS-$MsBookNo Vol $Volume. Debug dump: $dump"
    }
    Write-Log "Get-PartsBookTreeFromWeb: found $($sections.Count) section(s)"

    # 2. Figures per section via the lazy-load endpoint.
    $sb = New-Object System.Text.StringBuilder
    $figureCount = 0
    foreach ($sec in $sections) {
        [void]$sb.Append('<li sno="' + $sec.Number + '" special_char="' + $sec.Number + '"><div class="hitarea"></div><span>' + [System.Net.WebUtility]::HtmlEncode($sec.Name) + '</span><ul>')
        # Section 1 is the handbook front page (section_1.php) and carries no
        # figures, so only query the figure endpoint for sections >= 2.
        if ([int]$sec.Number -ge 2) {
            try {
                $figResp = Invoke-PmHttpPost -Url "$phbkBase/ajax/section_figures.php" -Form @{ booknum = $MsBookNo; vol = $Volume; section = $sec.Number } -TimeoutSec $TimeoutSec
                foreach ($fig in @(Get-PartsBookFigures -Html $figResp)) {
                    [void]$sb.Append('<li figno="' + $fig.Figno + '"><span class="go_fig">' + $fig.Figno + ' ' + [System.Net.WebUtility]::HtmlEncode($fig.Name) + '</span></li>')
                    $figureCount++
                }
            } catch {
                Write-Log "Get-PartsBookTreeFromWeb: figures for section $($sec.Number) failed: $($_.Exception.Message)"
            }
        }
        [void]$sb.Append('</ul></li>')
    }

    if ($figureCount -eq 0) {
        $dump = Save-DebugHtml -Html $sectionHtml -Tag 'phbk-tree'
        throw "No figures found for MS-$MsBookNo Vol $Volume. Debug dump: $dump"
    }

    Write-Log "Get-PartsBookTreeFromWeb: fetched $figureCount figure(s)"
    return '<html><body><div id="Ryan_fault"><ul id="phbk_tree" class="treeview">' + $sb.ToString() + '</ul></div></body></html>'
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
        [string]$PartsBookName,
        $ProgressHandle = $null
    )

    $mergeStart = Get-Date

    $combinedDir = Join-Path $SourceDir "CombinedSections"
    if (-not (Test-Path $combinedDir)) { New-Item -ItemType Directory -Path $combinedDir -Force | Out-Null }

    $csvFiles = @(Get-ChildItem -Path (Join-Path $SourceDir "HTML and CSV Files") -Filter "Figure *.csv" -ErrorAction SilentlyContinue)
    Write-Log "Merge: found $($csvFiles.Count) figure CSV(s) to combine"

    # Numeric ascending group sort (Section 2, 3, ..., 9, 10, ..., 20)
    $groups = @(
        $csvFiles |
            Group-Object { $_.BaseName -replace 'Figure (\d+)-\d+', '$1' } |
            Sort-Object -Property @{ Expression = { [int]$_.Name } }
    )
    Write-Log "Merge: $($groups.Count) section group(s) to process"

    # ---- Load site data and pre-add columns ONCE ----
    $siteData = $null
    $bookCol  = $null
    if ($SiteCsvPath -and (Test-Path $SiteCsvPath)) {
        $loadStart = Get-Date
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
        Write-Log "Merge: loaded $($siteData.Count) site row(s) in $([int]((Get-Date)-$loadStart).TotalSeconds)s; bookCol='$bookCol'"
    }

    $groupIdx = 0
    foreach ($group in $groups) {
        $groupIdx++
        $groupStart = Get-Date
        $figNum = $group.Name
        $sectionFile = Join-Path $combinedDir "Section $figNum.csv"

        # Pump the message queue so the UI doesn't go gray
        Update-PartsBookProgress -Handle $ProgressHandle `
            -Message "Merging section $figNum ($groupIdx of $($groups.Count))..." `
            -Percent (($groupIdx - 1) / [double]$groups.Count * 100) `
            -Book $PartsBookName

        Write-Log "Merge: section $figNum start ($($group.Group.Count) figure file(s))"

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
        Write-Log "Merge: section $figNum done — $($allRows.Count) row(s) in $([int]((Get-Date)-$groupStart).TotalSeconds)s"
    }

    if ($siteData -and (Test-Path $SiteCsvPath)) {
        $siteData | Export-Csv -Path $SiteCsvPath -NoTypeInformation
        Write-Log "Merge-PartsBookSections: updated $SiteCsvPath"
    }

    Write-Log "Merge-PartsBookSections: TOTAL $([int]((Get-Date)-$mergeStart).TotalSeconds)s"
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
			$ws = $workbook.Sheets.Add([System.Reflection.Missing]::Value,
									   $workbook.Sheets.Item($workbook.Sheets.Count))
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
                        # REF. stores the CSV filename (e.g. "Figure 2-1.csv").
                        # Strip the extension so we look for the HTML that
                        # Export-PartsBookFigures wrote alongside it.
                        $figName  = ("$v" -replace '\.csv$', '').Trim()
                        $htmlPath = Join-Path $SourceDir "HTML and CSV Files\$figName.html"
                        if (Test-Path $htmlPath) {
                            # file:/// URI so paths with spaces
                            # ("HTML and CSV Files") survive intact.
                            try {
                                $uri = (New-Object System.Uri -ArgumentList $htmlPath).AbsoluteUri
                            } catch {
                                $uri = $htmlPath
                            }
                            $cell = $ws.Cells.Item($r, $refIdx)
                            $ws.Hyperlinks.Add($cell, $uri, "", "", $figName) | Out-Null
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

function Repair-PartsBookFigureHyperlinks {
    # Walks every worksheet in the workbook, finds the REF. column, and
    # normalizes any figure hyperlink to point at the real ".html" file
    # sitting next to its ".csv".  Handles all four historical states:
    #   (a) bare text "Figure 2-1.csv", no hyperlink      -> add hyperlink
    #   (b) hyperlink -> "Figure 2-1.csv"                 -> retarget to .html
    #   (c) hyperlink -> "Figure 2-1.html", display dirty -> clean display
    #   (d) hyperlink -> "Figure 2-1.html", display clean -> skip
    # Returns the number of cells repaired.  Idempotent.
    param(
        [string]$ExcelPath,
        [string]$SourceDir
    )

    if (-not (Test-Path $ExcelPath)) { return 0 }
    $htmlCsvDir = Join-Path $SourceDir "HTML and CSV Files"
    if (-not (Test-Path $htmlCsvDir)) { return 0 }

    $repairs = 0
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($ExcelPath)

        foreach ($ws in @($wb.Worksheets)) {
            # Locate the REF. column (case-insensitive, tolerant of trailing space)
            $refCol = 0
            $lastCol = [Math]::Max(20, [int]$ws.UsedRange.Columns.Count)
            for ($c = 1; $c -le $lastCol; $c++) {
                $hdr = "$($ws.Cells.Item(1, $c).Value2)".Trim()
                if ($hdr -ieq 'REF.') { $refCol = $c; break }
            }
            if ($refCol -eq 0) { continue }

            $lastRow = [int]$ws.UsedRange.Rows.Count
            if ($lastRow -lt 2) { continue }

            for ($r = 2; $r -le $lastRow; $r++) {
                if (($r % 100) -eq 0) {
                    try { [System.Windows.Forms.Application]::DoEvents() } catch { }
                }

                $cell = $ws.Cells.Item($r, $refCol)
                $val  = "$($cell.Value2)"
                if ($val -notmatch 'Figure\s+\d+-\d+') { continue }

                # Canonical display text: "Figure 2-1" (no extension)
                $figName = ($val -replace '\.csv$', '' -replace '\.html$', '').Trim()
                $htmlPath = Join-Path $htmlCsvDir "$figName.html"
                if (-not (Test-Path $htmlPath)) { continue }   # nothing to link to

                try {
                    $uri = (New-Object System.Uri -ArgumentList $htmlPath).AbsoluteUri
                } catch {
                    $uri = $htmlPath
                }

                # Inspect the current state of this cell
                $existing = $null
                try {
                    if ($cell.Hyperlinks.Count -gt 0) { $existing = $cell.Hyperlinks.Item(1) }
                } catch { }

                $needsFix = $false
                if ($null -eq $existing) {
                    $needsFix = $true
                } else {
                    $target = "$($existing.Address)"
                    $targetIsHtml = ($target -match '\.html($|\?)')
                    $displayClean = ($val -eq $figName)
                    if (-not $targetIsHtml -or -not $displayClean) { $needsFix = $true }
                }

                if (-not $needsFix) { continue }

                # Delete any existing hyperlink(s) on the cell, then add a
                # clean one.  Excel won't let us mutate the address in place.
                try {
                    if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() }
                } catch { }

                $ws.Hyperlinks.Add($cell, $uri, "", "", $figName) | Out-Null
                $cell.Value2 = $figName        # belt & suspenders
                $repairs++
            }
        }

        if ($repairs -gt 0) { $wb.Save() }
    } catch {
        Write-Log "Repair-PartsBookFigureHyperlinks error: $($_.Exception.Message)"
    } finally {
        if ($wb) {
            try { $wb.Close($false) } catch { }    # we already Save()'d above if needed
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    return $repairs
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

                # Try the automatic tree fetch first; fall back to the
                # browser + paste flow only if that fails.
                $html = $null
                try {
                    $html = Get-PartsBookTreeFromWeb -MsBookNo "$($row.'MS Book No')" -Volume "$($row.Volume)"
                    Write-Log "Parts Book Creator: auto-fetched tree for $bookName"
                } catch {
                    Write-Log "Parts Book Creator: tree auto-fetch failed for $bookName — $($_.Exception.Message); falling back to manual paste"
                    $url = "https://www1.mtsc.usps.gov/apps/phbk/index.php?msbookno=$($row.'MS Book No')&volno=$($row.Volume)"
                    Start-Process $url
                    $html = Get-PartsBookTreeHtml -BookName $bookName
                }
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

                $combined = Merge-PartsBookSections -SourceDir $bookDir -SiteCsvPath $partsRoomCsv -PartsBookName $safeName -ProgressHandle $progress

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

    if ([string]::IsNullOrWhiteSpace($Html))       { throw "No HTML supplied." }
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { throw "No target path supplied." }

    # The closing </option> tag is optional in the wild (eMARS/MTSC pages emit
    # unclosed <option> elements); stop at the next <option>, </select>, or EOF.
    $optionMatches = [regex]::Matches(
        $Html,
        '(?is)<option\b(?<attrs>[^>]*)>(?<text>.*?)(?=</option>|<option\b|</select>|\z)')

    Write-Log "New-PartsVolumesCsv: matched $($optionMatches.Count) <option> element(s)"
    if ($optionMatches.Count -eq 0) {
        throw "No <option> elements matched. Confirm you pasted the entire <select> block."
    }

    $dir = Split-Path -Path $TargetPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("MS Book No,Volume,Full Name")

    $written = 0
    $seen = @{}
    foreach ($m in $optionMatches) {
        $attrs = $m.Groups['attrs'].Value
        $text  = $m.Groups['text'].Value

        $msBook = $null
        if     ($attrs -match '\bvalue\s*=\s*"([^"]*)"')  { $msBook = $matches[1] }
        elseif ($attrs -match "\bvalue\s*=\s*'([^']*)'")  { $msBook = $matches[1] }
        elseif ($attrs -match '\bvalue\s*=\s*([^\s>]+)')  { $msBook = $matches[1] }

        $volNo = $null
        if     ($attrs -match '\bvolno\s*=\s*"([^"]*)"')  { $volNo = $matches[1] }
        elseif ($attrs -match "\bvolno\s*=\s*'([^']*)'")  { $volNo = $matches[1] }
        elseif ($attrs -match '\bvolno\s*=\s*([^\s>]+)')  { $volNo = $matches[1] }

        if ([string]::IsNullOrWhiteSpace($msBook) -or [string]::IsNullOrWhiteSpace($volNo)) { continue }
        # Skip placeholder entries ("All MS Handbooks" / "--Select--").
        if ($msBook -ieq 'all' -or $volNo -ieq 'all') { continue }

        $display = [regex]::Replace($text, '(?is)<[^>]+>', '')
        $display = [System.Net.WebUtility]::HtmlDecode($display)
        $display = ($display -replace '\s+', ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($display)) { continue }

        $key = "$msBook|$volNo"
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        $esc = $display.Replace('"', '""')
        $lines.Add("$msBook,$volNo,""$esc""")
        $written++
    }

    $lines | Out-File -LiteralPath $TargetPath -Encoding UTF8
    Write-Log "New-PartsVolumesCsv: wrote $written row(s) to $TargetPath"
    return $written
}

# ---- Parts Volumes ----

function Get-PartsVolumesCsvFromWeb {
    param([Parameter(Mandatory)][string]$TargetPath)

    $override = ''
    if ($script:config.PSObject.Properties['PartsVolumesEndpointUrl']) {
        $override = "$($script:config.PartsVolumesEndpointUrl)".Trim()
    }

    $baseUrl = 'https://www1.mtsc.usps.gov/apps/mtsc/index.php'
    $targetUrl = $baseUrl
    if ($override) { $targetUrl = $override }
    $html = Invoke-PmHttpGet -Url $targetUrl

    # (a) inline <select>?
    $selectHtml = Find-SelectBlock -Html $html -Hints @('msbookSelect', 'msbookno', 'volno', 'booklist', 'Volume', 'Handbook')
    if ($selectHtml) {
        Write-Log "Get-PartsVolumesCsvFromWeb: found inline <select>"
        return (New-PartsVolumesCsv -Html $selectHtml -TargetPath $TargetPath)
    }

    # (a2) known lazy-load content pages — the MTSC shell loads these via the
    # menu; they are server-rendered and carry the handbook dropdown inline.
    foreach ($page in @(
        'contents/partsSearch_information.php',
        'contents/msbook_information.php'
    )) {
        $pageUrl = Resolve-RelativeUrl -Base $baseUrl -Relative $page
        try {
            $pageHtml = Invoke-PmHttpGet -Url $pageUrl -TimeoutSec 20
        } catch {
            Write-Log "  content page $pageUrl failed: $($_.Exception.Message)"
            continue
        }
        $sel = Find-SelectBlock -Html $pageHtml -Hints @('msbookSelect', 'msbookno', 'volno', 'booklist', 'Volume', 'Handbook')
        if ($sel) {
            Write-Log "Get-PartsVolumesCsvFromWeb: content page $pageUrl yielded a <select>"
            return (New-PartsVolumesCsv -Html $sel -TargetPath $TargetPath)
        }
        if ($pageHtml -match '(?i)<option\b' -and $pageHtml -match '(?i)volno|msbook') {
            Write-Log "Get-PartsVolumesCsvFromWeb: content page $pageUrl yielded raw <option>s"
            return (New-PartsVolumesCsv -Html $pageHtml -TargetPath $TargetPath)
        }
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    $addCandidate = {
        param([string]$u)
        if (-not [string]::IsNullOrWhiteSpace($u) -and -not $candidates.Contains($u)) {
            [void]$candidates.Add($u)
        }
    }

    # (b) URLs already sitting in the shell page
    foreach ($u in (Find-AjaxEndpoints -Html $html)) { & $addCandidate $u }

    # (c) walk every linked script and scan it too
    $scriptSrcs = @()
    foreach ($m in [regex]::Matches($html, '(?is)<script\b[^>]*\bsrc\s*=\s*[''"](?<src>[^''"]+)[''"]')) {
        $scriptSrcs += $m.Groups['src'].Value
    }
    Write-Log "Get-PartsVolumesCsvFromWeb: $($scriptSrcs.Count) linked script(s) to scan"

    foreach ($src in $scriptSrcs) {
        $lower = $src.ToLowerInvariant()
        # Chase anything plausibly parts-search related; skip pure vendor libs.
        if ($lower -match '(jquery|datatables|highcharts|colorbox|layerslider|maskedinput|simplepagination|hash\.js|feedback\.js|videos\.js)') { continue }

        $jsUrl = Resolve-RelativeUrl -Base $baseUrl -Relative $src
        try {
            $js = Invoke-PmHttpGet -Url $jsUrl -TimeoutSec 15
        } catch {
            Write-Log "  script scan failed $jsUrl : $($_.Exception.Message)"
            continue
        }
        $hits = Find-AjaxEndpoints -Html $js
        if ($hits.Count -gt 0) {
            Write-Log "  $jsUrl -> $($hits.Count) URL(s)"
            foreach ($h in $hits) { Write-Log "     $h" }
        }
        foreach ($u in $hits) { & $addCandidate $u }
    }

    # (d) final fallback: the known content pages plus a handful of plausible paths
    foreach ($u in @(
        'contents/partsSearch_information.php',
        'contents/msbook_information.php',
        'contents/parts_search/parts_search.php',
        'contents/parts_search/partssearch.php',
        'contents/partsSearch/partsSearch.php',
        'contents/partssearch.php',
        'contents/mtsc/parts_search.php'
    )) { & $addCandidate $u }

    Write-Log "Get-PartsVolumesCsvFromWeb: $($candidates.Count) candidate URL(s)"
    foreach ($c in $candidates) { Write-Log "  -> $c" }

    $dumpPath = Save-DebugHtml -Html $html -Tag 'parts-search'
    Write-Log "Get-PartsVolumesCsvFromWeb: page dump -> $dumpPath"

    foreach ($candidate in $candidates) {
        # Ignore anything obviously not content
        if ($candidate -match '\.(js|css|png|jpg|jpeg|gif|ico|svg|woff\d?)(\?|$)') { continue }
        if ($candidate -match '^javascript:|^#') { continue }

        $resolved = Resolve-RelativeUrl -Base $baseUrl -Relative $candidate
        if (-not $resolved) { continue }

        try {
            $body = Invoke-PmHttpGet -Url $resolved -TimeoutSec 20
        } catch {
            Write-Log "  candidate $resolved failed: $($_.Exception.Message)"
            continue
        }

        $sel = Find-SelectBlock -Html $body -Hints @('msbookSelect', 'msbookno', 'volno', 'booklist', 'Volume')
        if ($sel) {
            Write-Log "  candidate $resolved yielded a <select>"
            return (New-PartsVolumesCsv -Html $sel -TargetPath $TargetPath)
        }
        if ($body -match '(?i)<option\b' -and $body -match '(?i)volno|msbookno') {
            Write-Log "  candidate $resolved yielded raw <option>s"
            return (New-PartsVolumesCsv -Html $body -TargetPath $TargetPath)
        }
        try {
            $json = $body | ConvertFrom-Json
            $count = New-PartsVolumesCsvFromJson -Json $json -TargetPath $TargetPath
            if ($count -gt 0) {
                Write-Log "  candidate $resolved yielded JSON with $count row(s)"
                return $count
            }
        } catch { }
    }

    throw "Could not locate the parts-book dropdown. Debug dump: $dumpPath"
}

function New-PartsVolumesCsvFromJson {
    # Best-effort JSON parser for the parts-book dropdown payload.
    # Looks for objects that carry a book number, a volume, and a name.
    param($Json, [string]$TargetPath)

    $rows = New-Object System.Collections.ArrayList
    $seen = @{}

    $stack = New-Object System.Collections.Stack
    $stack.Push($Json)
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        if ($null -eq $node) { continue }

        if ($node -is [System.Collections.IEnumerable] -and $node -isnot [string]) {
            foreach ($child in $node) { $stack.Push($child) }
            continue
        }
        if ($node -isnot [PSCustomObject]) { continue }

        $props = @($node.PSObject.Properties)
        $bProp = $props | Where-Object { $_.Name -match '(?i)ms.?book|^book$|bookno' } | Select-Object -First 1
        $vProp = $props | Where-Object { $_.Name -match '(?i)^vol|volume|volno'      } | Select-Object -First 1
        $nProp = $props | Where-Object { $_.Name -match '(?i)full.?name|^name$|label|text' } | Select-Object -First 1

        if ($bProp -and $vProp -and $nProp) {
            $key = "$($bProp.Value)|$($vProp.Value)"
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$rows.Add([PSCustomObject]@{
                    MsBook   = "$($bProp.Value)".Trim()
                    Volume   = "$($vProp.Value)".Trim()
                    FullName = "$($nProp.Value)".Trim()
                })
            }
        }
    }

    if ($rows.Count -eq 0) { return 0 }

    $dir = Split-Path -Path $TargetPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('MS Book No,Volume,Full Name')
    foreach ($r in $rows) {
        $esc = $r.FullName.Replace('"', '""')
        $lines.Add("$($r.MsBook),$($r.Volume),`"$esc`"")
    }
    $lines | Out-File -LiteralPath $TargetPath -Encoding UTF8
    Write-Log "New-PartsVolumesCsvFromJson: wrote $($rows.Count) row(s) to $TargetPath"
    return $rows.Count
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

    # ---- Try automatic fetch first ----
    try {
        $count = Get-PartsVolumesCsvFromWeb -TargetPath $targetPath
        [System.Windows.Forms.MessageBox]::Show(
            "Fetched and wrote $count entries to:`r`n$targetPath",
            "Parts Volumes List", "OK", "Information")
        return $true
    } catch {
        Write-Log "Show-PartsVolumesWizard: auto-fetch failed: $($_.Exception.Message)"
        $fallback = [System.Windows.Forms.MessageBox]::Show(
            "Automatic fetch failed.`r`n`r`n$($_.Exception.Message)`r`n`r`nFall back to the manual paste workflow?",
            "Parts Volumes List",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($fallback -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    # ---- Manual paste fallback (existing flow, unchanged) ----
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

    # Match every <option …>…</option>.  Attributes are captured as one
    # blob and parsed separately, so we tolerate: uppercase tags, attribute
    # order, single/double/no quotes, multiline display text, and unclosed
    # <option> tags (the closing tag is optional in HTML).
    $optionMatches = [regex]::Matches(
        $Html,
        '(?is)<option\b(?<attrs>[^>]*)>(?<text>.*?)(?=</option>|<option\b|</select>|\z)')

    Write-Log "New-SitesCsv: matched $($optionMatches.Count) <option> element(s)"
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
        $attrs = $m.Groups['attrs'].Value
        $text  = $m.Groups['text'].Value

        $val = $null
        if     ($attrs -match '\bvalue\s*=\s*"([^"]*)"')   { $val = $matches[1] }
        elseif ($attrs -match "\bvalue\s*=\s*'([^']*)'")   { $val = $matches[1] }
        elseif ($attrs -match '\bvalue\s*=\s*([^\s>]+)')   { $val = $matches[1] }

        if ([string]::IsNullOrWhiteSpace($val)) { continue }

        $display = [regex]::Replace($text, '(?is)<[^>]+>', '')
        $display = [System.Net.WebUtility]::HtmlDecode($display)
        $display = ($display -replace '\s+', ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($display)) { continue }

        $siteId = $val.Trim()
        if ([string]::IsNullOrWhiteSpace($siteId)) { continue }
        if ($seen.ContainsKey($siteId)) { continue }
        $seen[$siteId] = $true

        $safeName = $display.Replace('"', '""')
        $lines.Add("$siteId,""$safeName""")
        $written++
    }

    $lines | Out-File -LiteralPath $TargetPath -Encoding UTF8
    Write-Log "New-SitesCsv: wrote $written row(s) to $TargetPath"
    return $written
}

# ---- Sites ----

function Get-SitesCsvFromCriteriaPage {
    # Automatic retrieval of the sites list.  Returns the row count, or throws.
    param([Parameter(Mandatory)][string]$TargetPath)

    # Config override wins, if the user has hardcoded a discovered endpoint
    $override = ''
    if ($script:config.PSObject.Properties['SitesListEndpointUrl']) {
        $override = "$($script:config.SitesListEndpointUrl)".Trim()
    }

    $url = if ($override) {
        $override
    } else {
        'http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_criteria'
    }

    $html = Invoke-PmHttpGet -Url $url -UseDefaultCredentials

    # (a) inline <select>?
    $selectHtml = Find-SelectBlock -Html $html -Hints @(
        'Site ID', 'site_id', 'p_site_id', 'Full Name', 'FullName', 'siteid'
    )
    if ($selectHtml) {
        Write-Log "Get-SitesCsvFromCriteriaPage: found inline <select> ($($selectHtml.Length) chars)"
        try {
            return (New-SitesCsv -Html $selectHtml -TargetPath $TargetPath)
        } catch {
            $dump = Save-DebugHtml -Html $selectHtml -Tag 'sites-select'
            Write-Log "Get-SitesCsvFromCriteriaPage: select parse failed — dump saved: $dump"
            throw
        }
    }

    # (b) AJAX endpoint scan
    Write-Log "Get-SitesCsvFromCriteriaPage: no inline <select>; scanning JS for endpoints"
    $candidates = @(
        Find-AjaxEndpoints -Html $html -Keyword 'site'
        Find-AjaxEndpoints -Html $html -Keyword 'stockroom'
    ) | Select-Object -Unique

    Write-Log "Get-SitesCsvFromCriteriaPage: $($candidates.Count) candidate endpoint(s):"
    foreach ($c in $candidates) { Write-Log "  -> $c" }

    $dumpPath = Save-DebugHtml -Html $html -Tag 'sites-criteria'
    Write-Log "Get-SitesCsvFromCriteriaPage: page dump -> $dumpPath"

    foreach ($candidate in $candidates) {
        $resolved = Resolve-RelativeUrl -Base $url -Relative $candidate
        if (-not $resolved) { continue }
        try {
            $body = Invoke-PmHttpGet -Url $resolved -UseDefaultCredentials
        } catch {
            Write-Log "  candidate $resolved failed: $($_.Exception.Message)"
            continue
        }

        $sel = Find-SelectBlock -Html $body -Hints @(
            'Site ID', 'site_id', 'p_site_id', 'Full Name'
        )
        if ($sel) {
            Write-Log "  candidate $resolved yielded a <select>"
            return (New-SitesCsv -Html $sel -TargetPath $TargetPath)
        }
        if ($body -match '(?i)<option\b') {
            Write-Log "  candidate $resolved yielded raw <option>s"
            return (New-SitesCsv -Html $body -TargetPath $TargetPath)
        }
    }

    throw "Could not locate the sites dropdown. Debug dump: $dumpPath"
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

    # ---- Try automatic fetch first ----
    try {
        $count = Get-SitesCsvFromCriteriaPage -TargetPath $targetPath
        [System.Windows.Forms.MessageBox]::Show(
            "Fetched and wrote $count entries to:`r`n$targetPath",
            "Sites List", "OK", "Information")
        return $true
    } catch {
        Write-Log "Show-SitesWizard: auto-fetch failed: $($_.Exception.Message)"
        $fallback = [System.Windows.Forms.MessageBox]::Show(
            "Automatic fetch failed.`r`n`r`n$($_.Exception.Message)`r`n`r`nFall back to the manual paste workflow?",
            "Sites List",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($fallback -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    # ---- Manual paste fallback (existing flow, unchanged) ----
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
