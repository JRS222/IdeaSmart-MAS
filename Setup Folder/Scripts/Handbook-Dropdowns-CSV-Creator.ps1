# Handbook-Dropdowns-CSV-Creator.ps1
#
# Parses the MTSC parts handbook HTML into Parsed-Parts-Volumes.csv and drops
# it in the setup folder root, next to the other setup CSVs (sites.csv, etc.).
#
# No config required. No folder prompts. Idempotent.

Add-Type -AssemblyName System.Windows.Forms

################################################################################
#                       Resolve the setup folder                               #
################################################################################

$scriptDir = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

# If we're inside a "Scripts" subfolder, the setup root is one level up.
# Otherwise, we're already at the setup root.
if ((Split-Path -Leaf $scriptDir) -ieq 'Scripts') {
    $setupDir = Split-Path -Parent $scriptDir
} else {
    $setupDir = $scriptDir
}

$csvFilePath = Join-Path $setupDir 'Parsed-Parts-Volumes.csv'

Write-Host "Setup folder: $setupDir"
Write-Host "Target CSV:   $csvFilePath"

################################################################################
#                          Idempotency check                                   #
################################################################################

if (Test-Path -LiteralPath $csvFilePath) {
    [System.Windows.Forms.MessageBox]::Show(
        "Parsed-Parts-Volumes.csv already exists at:`n$csvFilePath`n`nNothing to do.",
        "File Exists",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    return
}

################################################################################
#                          Collect HTML from user                              #
################################################################################

$url = "https://www1.mtsc.usps.gov/apps/mtsc/index.php#Doc&partssearch&0&NA"
Start-Process $url

[System.Windows.Forms.MessageBox]::Show(
    "Navigate to the URL and copy the entire HTML content of the relevant <select> element. " +
    "Paste the copied content in the next prompt.",
    "Instructions",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

$form = New-Object System.Windows.Forms.Form
$form.Text = "Paste HTML Content"
$form.Width = 600
$form.Height = 400

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = 'Vertical'
$textBox.Width = 550
$textBox.Height = 300
$textBox.Top = 10
$textBox.Left = 10
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Top = 320
$okButton.Left = 250
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

$form.Add_Shown({ $form.Activate() })
$form.ShowDialog() | Out-Null

$htmlContent = $textBox.Text
if ([string]::IsNullOrWhiteSpace($htmlContent)) {
    [System.Windows.Forms.MessageBox]::Show(
        "No HTML content was pasted. Aborting.",
        "Nothing to Parse",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    return
}

################################################################################
#                             Parse <option> tags                              #
################################################################################

$regex = '<option\s+value="(?<msbookno>[^"]+)"\s+volno="(?<volno>[^"]+)">(?<fullname>.*?)<\/option>'
$matches = [regex]::Matches($htmlContent, $regex)

if ($matches.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "No <option> elements matched. Check that you copied the entire <select> block.",
        "No Matches",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    return
}

################################################################################
#                               Write CSV                                      #
################################################################################

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("MS Book No,Volume,Full Name")

foreach ($m in $matches) {
    $msbookno = $m.Groups['msbookno'].Value
    $volno    = $m.Groups['volno'].Value
    $fullname = $m.Groups['fullname'].Value.Replace('"', '""')   # escape embedded quotes
    $lines.Add("$msbookno,$volno,""$fullname""")
}

$lines | Out-File -LiteralPath $csvFilePath -Encoding UTF8

Write-Host "Wrote $($matches.Count) entries to $csvFilePath"

[System.Windows.Forms.MessageBox]::Show(
    "Parsed-Parts-Volumes.csv created at:`n$csvFilePath`n`n$($matches.Count) entries written.",
    "Success",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null