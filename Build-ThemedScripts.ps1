<#
.SYNOPSIS
    Generates the hardcoded themed copies of UI-Script.ps1.

.DESCRIPTION
    UI-Script.ps1 is the single master script.  Its BUILD-TIME THEME
    CONSTANT block ($script:theme = 'Default') is the only place a theme
    is chosen.  This script stamps that one line and writes:

        UI-Script.ps1            (master — Default, never overwritten)
        UI-Script-Dark.ps1       -> $script:theme = 'Dark'
        UI-Script-Postal.ps1     -> $script:theme = 'Postal'

    The copies differ from the master ONLY on that line, so edits are
    made exclusively in the master and the copies are regenerated.

    The runtime theme picker was removed from the app because re-theming
    a live WinForms window in PowerShell is too slow; run the themed
    script you want instead.

.EXAMPLE
    .\Build-ThemedScripts.ps1

.PARAMETER Themes
    Themes to generate.  Default: Dark, Postal.
#>
[CmdletBinding()]
param(
    [string[]]$Themes = @('Dark', 'Postal')
)

$ErrorActionPreference = 'Stop'

$srcPath = Join-Path $PSScriptRoot 'UI-Script.ps1'
if (-not (Test-Path $srcPath)) { throw "Master script not found: $srcPath" }

$content = [System.IO.File]::ReadAllText($srcPath)

$marker = 'BUILD-TIME THEME CONSTANT'
$mPos = $content.IndexOf($marker)
if ($mPos -lt 0) { throw "Theme marker not found in $srcPath — is the master up to date?" }

# The first theme assignment at/after the marker block.
$after = $content.Substring($mPos)
$m = [regex]::Match($after, '(?m)\$script:theme\s*=\s*''[^'']*''')
if (-not $m.Success) { throw 'Could not locate the $script:theme assignment after the marker.' }

$absStart = $mPos + $m.Index

foreach ($theme in $Themes) {
    if ($theme -notin @('Dark', 'Postal')) {
        Write-Warning "Skipping unknown theme '$theme' (expected Dark or Postal)."
        continue
    }
    $themed = $content.Substring(0, $absStart) +
              "`$script:theme = '$theme'" +
              $content.Substring($absStart + $m.Length)
    $outPath = Join-Path $PSScriptRoot "UI-Script-$theme.ps1"
    [System.IO.File]::WriteAllText($outPath, $themed, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Wrote $(Split-Path $outPath -Leaf) — theme '$theme'"
}

Write-Host 'Done.'
