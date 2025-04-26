# PowerShell script to parse HTML <select> options and save to CSV

# Load or create the configuration
if (Test-Path $configFilePath) {
    $config = Get-Content -Path $configFilePath | ConvertFrom-Json
} else {
    $config = [PSCustomObject]@{ 
        PartsBooksDir = ""; 
        ScriptsDir = ""; 
        RequiredCsvsDir = "";
        Books = @{}
    }
}

# Ensure ParsedPartsVolumesCsvPath property exists
if (-not $config.PSObject.Properties.Match('ParsedPartsVolumesCsvPath')) {
    Add-Member -InputObject $config -MemberType NoteProperty -Name 'ParsedPartsVolumesCsvPath' -Value ''
}


# Define the CSV file path
$csvFilePath = Join-Path -Path $config.RequiredCsvsDir -ChildPath "Parsed-Parts-Volumes.csv"

if (-not [string]::IsNullOrEmpty($csvFilePath) -and (Test-Path $csvFilePath)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("CSV file already exists at $csvFilePath. The script will now terminate.", "File Exists", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
}

# If the CSV file does not exist, create it
$url = "https://www1.mtsc.usps.gov/apps/mtsc/index.php#Doc&partssearch&0&NA"

# Open the URL in the default web browser
Start-Process $url

# Prompt user to navigate to the webpage and copy the relevant HTML content
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Navigate to the URL and copy the entire HTML content of the relevant <select> element. Paste the copied content in the next prompt.", "Instructions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Create a form with a text box for the user to paste the HTML content
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
$okButton.Add_Click({$form.Close()})
$form.Controls.Add($okButton)

$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

$htmlContent = $textBox.Text

# Use regular expression to extract details from each <option> element
$regex = '<option\s+value="(?<msbookno>[^"]+)"\s+volno="(?<volno>[^"]+)">(?<fullname>.*?)<\/option>'
$matches = [regex]::Matches($htmlContent, $regex)

# Create an array to store CSV lines
$csvLines = "MS Book No,Volume,Full Name" # Header row

# Loop through each match and build CSV lines
foreach ($match in $matches) {
    $msbookno = $match.Groups["msbookno"].Value
    $volno = $match.Groups["volno"].Value
    $fullname = $match.Groups["fullname"].Value

    # Add to CSV lines
    $csvLines += "`n$msbookno,$volno,""$fullname"""
}

# Define the CSV file path
$csvFilePath = Join-Path -Path $config.RequiredCsvsDir -ChildPath "Parsed-Parts-Volumes.csv"

# Output CSV content to file
$csvLines | Out-File -FilePath $csvFilePath -Encoding ASCII

Write-Host "CSV file has been created at $csvFilePath"

# Update the JSON configuration with the new CSV file path
$config.ParsedPartsVolumesCsvPath = $csvFilePath
$config | ConvertTo-Json -Depth 4 | Set-Content -Path $configFilePath -Force

# Verify that the property was set correctly
$updatedConfig = Get-Content -Path $configFilePath | ConvertFrom-Json
if ($updatedConfig.ParsedPartsVolumesCsvPath -eq $csvFilePath) {
    Write-Host "Configuration updated successfully."
} else {
    Write-Host "Warning: Configuration may not have updated correctly."
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("CSV file has been created at $csvFilePath and the path has been updated in the configuration.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)