################################################################################
#                                                                              #
#                          Parts Books Creator                                 #
#                                                                              #
################################################################################

################################################################################
#                          Required .NET Assemblies                            #
################################################################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

################################################################################
#                           Global Variables                                   #
################################################################################

# Configuration variables
$global:configPath = Join-Path $PSScriptRoot "Config.json"  # Path to configuration
$script:ScriptDirectory = $PSScriptRoot    # Script directory path
$script:configPath = Join-Path $script:ScriptDirectory "Config.json"  # Config path

# UI Controls
$global:listView = New-Object System.Windows.Forms.ListView  # ListView for books

# Progress UI elements
$global:progressForm = New-Object System.Windows.Forms.Form  # Progress form
$global:progressBar = New-Object System.Windows.Forms.ProgressBar  # Progress bar
$global:progressLabel = New-Object System.Windows.Forms.Label  # Progress description
$global:bookLabel = New-Object System.Windows.Forms.Label  # Current book label

# State tracking
$script:existingBooks = $null             # Existing books collection 
$script:machinesUpdated = $false          # Flag for machines list update

$global:configPath = Join-Path $PSScriptRoot "Config.json"
$global:listView = New-Object System.Windows.Forms.ListView


# Load configuration
$configPath = Join-Path $PSScriptRoot "Config.json"
Write-Host "Attempting to load config from: $configPath" 


if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath | ConvertFrom-Json
    Write-Host "Config loaded successfully"
} else {
    Write-Host "Error: Config file not found at $configPath"
    exit
}

$partsBooksDirPath = $config.PartsBooksDirectory
Write-Host "Parts Books Directory: $partsBooksDirPath"

if (-not $partsBooksDirPath) {
    Write-Host "Error: PartsBooksDirectory is not set in the config file."
    exit
}

$global:progressForm = New-Object System.Windows.Forms.Form
$global:progressForm.Text = "Processing Parts Books"
$global:progressForm.Size = New-Object System.Drawing.Size(400,150)
$global:progressForm.StartPosition = 'CenterScreen'

$global:progressBar = New-Object System.Windows.Forms.ProgressBar
$global:progressBar.Size = New-Object System.Drawing.Size(360,20)
$global:progressBar.Location = New-Object System.Drawing.Point(10,10)
$global:progressForm.Controls.Add($global:progressBar)

$global:progressLabel = New-Object System.Windows.Forms.Label
$global:progressLabel.Size = New-Object System.Drawing.Size(360,40)
$global:progressLabel.Location = New-Object System.Drawing.Point(10,40)
$global:progressForm.Controls.Add($global:progressLabel)

$global:bookLabel = New-Object System.Windows.Forms.Label
$global:bookLabel.Size = New-Object System.Drawing.Size(360,20)
$global:bookLabel.Location = New-Object System.Drawing.Point(10,90)
$global:progressForm.Controls.Add($global:bookLabel)

# Paths
$HandbookSelectionCSV = Join-Path $config.DropdownCsvsDirectory "Parsed-Parts-Volumes.csv"
Write-Host "HandbookSelectionCSV path: $HandbookSelectionCSV"

if (-not (Test-Path $HandbookSelectionCSV)) {
    Write-Host "Error: CSV file not found at path: $HandbookSelectionCSV"
    Write-Host "Files in DropdownCsvsDirectory:"
    Get-ChildItem $config.DropdownCsvsDirectory | ForEach-Object { Write-Host $_.Name }
    exit
}

try {
    $csvData = Import-Csv -Path $HandbookSelectionCSV
    Write-Host "CSV data loaded successfully. Row count: $($csvData.Count)"
    Write-Host "First row: $($csvData[0] | Out-String)"
} catch {
    Write-Host "Error loading CSV data: $_"
    exit
}

################################################################################
#                       Configuration Management                               #
################################################################################

function Get-ExistingBooks {
    Write-Host "Config path in Get-ExistingBooks: $script:configPath"
    if (-not (Test-Path $script:configPath)) {
        Write-Host "Config file not found at $script:configPath"
        return @{}  # Return an empty hashtable
    }
    $config = Get-Content -Path $script:configPath | ConvertFrom-Json
    if ($null -eq $config) {
        Write-Host "Failed to load config from $script:configPath"
        return @{}
    }
    Write-Host "Config loaded in Get-ExistingBooks"
    $books = @{}
    if ($config.PSObject.Properties.Name -contains "Books" -and $null -ne $config.Books) {
        Write-Host "Config contains 'Books' property"
        foreach ($key in $config.Books.PSObject.Properties.Name) {
            Write-Host "Adding book '$key' to existingBooks"
            $books[$key] = $config.Books.$key
        }
    } else {
        Write-Host "Config does not contain 'Books' property or it's null"
    }
    return $books
}

 
# Initialize $existingBooks
$existingBooks = Get-ExistingBooks
if ($null -eq $existingBooks) {
    $existingBooks = @{}
    Write-Host "Warning: existingBooks was null, initializing empty hashtable"
}

Write-Host "Existing books: $($existingBooks.Keys -join ', ')"

foreach ($row in $csvData) {
    Write-Host "Processing row: $($row.'Full Name')"
    if (-not $existingBooks.ContainsKey($row.'Full Name')) {
        $item = New-Object System.Windows.Forms.ListViewItem($row.'Full Name')
        $item.SubItems.Add($row.'MS Book No')
        $item.SubItems.Add($row.'Volume')
        $listView.Items.Add($item)
        Write-Host "Added item to ListView: $($row.'Full Name')"
    } else {
        Write-Host "Skipped existing book: $($row.'Full Name')"
    }
}
Write-Host "ListView items count: $($global:listView.Items.Count)"

Write-Host "Script is running from: "
Write-Host "DropdownCsvsDirectory: $($config.DropdownCsvsDirectory)"
Write-Host "PartsBooksDirectory: $($config.PartsBooksDirectory)"
Write-Host "PartsRoomDirectory: $($config.PartsRoomDirectory)"

function Update-ConfigBooks($newBooks) {
    Write-Host "Config path in Update-ConfigBooks: $script:configPath"
    if (-not (Test-Path $script:configPath)) {
        Write-Host "Config file not found at $script:configPath"
        return
    }
    $config = Get-Content -Path $script:configPath | ConvertFrom-Json

    if (-not $config.PSObject.Properties['Books']) {
        $config | Add-Member -NotePropertyName 'Books' -NotePropertyValue @{}
    }

    foreach ($book in $newBooks.GetEnumerator()) {
        if (-not $config.Books.PSObject.Properties[$book.Key]) {
            $config.Books | Add-Member -NotePropertyName $book.Key -NotePropertyValue $book.Value
        } else {
            $config.Books.$($book.Key) = $book.Value
        }
    }

    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $script:configPath
    Write-Host "Updated Config.json with new books: $($newBooks.Keys -join ', ')"
}

# Function to update config with new book information
function Update-Config($bookName, $volumesToUrlPath, $sectionNamesCsvPath) {
    # Read the current config
    $config = Get-Content -Path $configPath | ConvertFrom-Json

    # Ensure $config.Books is initialized as a hashtable
    if (-not ($config.PSObject.Properties.Name -contains "Books") -or $null -eq $config.Books) {
        $config | Add-Member -NotePropertyName "Books" -NotePropertyValue @{} -Force
    } elseif ($config.Books -isnot [System.Collections.IDictionary]) {
        $existingBooks = @{}
        foreach ($prop in $config.Books.PSObject.Properties) {
            $existingBooks[$prop.Name] = $prop.Value
        }
        $config.Books = $existingBooks
    }

    # Add or update the book in the config
    $config.Books[$bookName] = @{
        VolumesToUrlCsvPath = $volumesToUrlPath
        SectionNamesCsvPath = $sectionNamesCsvPath
    }

    # Write the updated config back to the file
    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configPath
    Write-Host "Config updated for book: $bookName"
}

################################################################################
#                      UI and Progress Reporting                               #
################################################################################

# Function to update progress
function Update-Progress($stepName, $percentComplete, $currentBook) {
    $global:progressBar.Value = $percentComplete
    $global:progressLabel.Text = $stepName
    $global:bookLabel.Text = "Current Book: $currentBook"
    $global:progressForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# Show a message box (only used for final message now)
function Show-Message ($message, $title = "Information") {
    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

################################################################################
#                           Utility Functions                                  #
################################################################################

# Function to sanitize file/folder names
function Sanitize-Name($name) {
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars()
    foreach ($char in $invalidChars) {
        $name = $name -replace [regex]::Escape($char), '-'
    }
    return $name
}

################################################################################
#                       HTML and Web Processing                                #
################################################################################

# Function to get HTML content
function Get-VerifiedHtmlContent($handbookName) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Open HTML for $handbookName"
    $form.Size = New-Object System.Drawing.Size(600, 400)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(580, 40)
    $label.Text = "For $handbookName, please follow these steps:
1. Go to the URL that opened in your browser
2. Use right-click > 'View Page Source' or press Ctrl+U
3. Press Ctrl+A to select all, then Ctrl+C to copy
4. Paste the HTML below:"
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.RichTextBox
    $textBox.Size = New-Object System.Drawing.Size(570, 250)
    $textBox.Location = New-Object System.Drawing.Point(10, 60)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250, 320)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = 'OK'
    $okButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($okButton)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        Write-Host "Operation cancelled by user."
        exit
    }
}

function Process-HTMLContent {
    param (
        [string]$htmlContent,
        [PSObject]$selectedRow,
        [string]$directoryPath
    )
    
    # Define a helper function to dump the HTML content to a file for troubleshooting
    function Save-DebugContent {
        param (
            [string]$content,
            [string]$path
        )
        
        $debugPath = Join-Path -Path $path -ChildPath "debug_html.txt"
        $content | Out-File -FilePath $debugPath -Encoding UTF8
        Write-Host "Saved debug HTML content to $debugPath for troubleshooting."
    }

    Write-Host "Starting Process-HTMLContent..."
    # Create the output directory if it doesn't exist
    if (-not (Test-Path -Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath | Out-Null
        Write-Host "Created directory: $directoryPath"
    }

    # Initialize HTML document object
    $htmlDoc = New-Object -ComObject "HTMLFile"
    
    Write-Host "Parsing HTML content..."
    # Handle different methods of writing to the HTMLFile object based on PowerShell version
    try {
        # For PowerShell 5.1 with newer IE versions
        $htmlDoc.IHTMLDocument2_write($htmlContent)
    } catch {
        try {
            # For PowerShell 5.1 with older IE versions
            $enc = [System.Text.Encoding]::Unicode
            $byteArray = $enc.GetBytes($htmlContent)
            $htmlDoc.write($byteArray)
        } catch {
            # If all methods fail, try a regex-based fallback approach
            Write-Host "Warning: DOM methods failed. Using regex fallback for HTML parsing."
            return Process-HTMLContentWithRegex -htmlContent $htmlContent -selectedRow $selectedRow -directoryPath $directoryPath
        }
    }

    Write-Host "HTML parsed successfully."
    # Get MS Book No and Volume from selectedRow or extract from the HTML if available
    $msBookNo = $selectedRow.'MS Book No'
    $volume = $selectedRow.Volume

    # If we don't have the MS Book No or Volume from selectedRow, try to extract from HTML
    if ([string]::IsNullOrEmpty($msBookNo) -or [string]::IsNullOrEmpty($volume)) {
        Write-Host "Getting book info from HTML content..."
        $bookTitleElement = $htmlDoc.getElementById("book_title")
        if ($bookTitleElement) {
            $bookTitleText = $bookTitleElement.innerText
            # Assuming format is "MS{number} VOLUME {letter}"
            if ($bookTitleText -match "MS(\d+)\s+VOLUME\s+([A-Z])") {
                $msBookNo = $matches[1]
                $volume = $matches[2]
                Write-Host "Extracted from HTML: MS Book No: $msBookNo, Volume: $volume"
            }
        }
    }

    # If HTML parsing fails, save the content for debugging
    if ([string]::IsNullOrEmpty($msBookNo) -or [string]::IsNullOrEmpty($volume)) {
        Write-Host "Warning: Could not extract MS Book No and Volume from HTML content."
        Write-Host "Saving HTML content for debugging..."
        Save-DebugContent -content $htmlContent -path $directoryPath
        
        # Try to get information from the filename or path as a last resort
        $dirName = Split-Path -Path $directoryPath -Leaf
        if ($dirName -match "MS-(\d+).*\(Vol\.\s*([A-Z])\)") {
            $msBookNo = $matches[1]
            $volume = $matches[2]
            Write-Host "Extracted from directory name: MS Book No: $msBookNo, Volume: $volume"
        } else {
            Write-Host "Error: Could not determine MS Book No and Volume. Exiting function."
            return
        }
    }

    # Create paths for output files
    $volumesToUrlPath = Join-Path -Path $directoryPath -ChildPath "Volumes-to-URL.csv"
    $sectionNamesPath = Join-Path -Path $directoryPath -ChildPath "SectionNames.txt"

    Write-Host "Output files will be:"
    Write-Host "  Volumes-to-URL.csv: $volumesToUrlPath"
    Write-Host "  SectionNames.txt: $sectionNamesPath"

    # Create CSV writer for Volumes-to-URL.csv
    $csvHeader = "Figure No.,Name,Section No.,MS Book No,Volume,URL"
    $csvHeader | Out-File -FilePath $volumesToUrlPath -Encoding UTF8
    
    # Initialize array for section names
    $sectionNames = @()

    # Try to get the Ryan_fault div
    $ryanFaultDiv = $htmlDoc.getElementById("Ryan_fault")
    
    # If not found directly, try searching through all divs with matching style
    if (-not $ryanFaultDiv) {
        Write-Host "Ryan_fault div not found by ID, searching by attributes..."
        $allDivs = $htmlDoc.getElementsByTagName("div")
        foreach ($div in $allDivs) {
            $style = $div.getAttribute("style")
            if ($style -and $style -like "*height:100%*overflow-y:scroll*") {
                Write-Host "Found div with matching style attributes, using this instead."
                $ryanFaultDiv = $div
                break
            }
        }
    }
    
    if ($ryanFaultDiv) {
        Write-Host "Found content div, processing content..."
        # Get the UL with ID "phbk_tree" which contains all sections and figures
        $treeUl = $ryanFaultDiv.getElementsByTagName("ul") | Where-Object { $_.id -eq "phbk_tree" -or $_.className -eq "treeview" } | Select-Object -First 1
        
        if ($treeUl) {
            Write-Host "Found phbk_tree, extracting sections and figures..."
            # Loop through each LI (section) in the tree
            $sectionElements = $treeUl.getElementsByTagName("li") | Where-Object { $_.getAttribute("sno") }
            
            Write-Host "Found $($sectionElements.Count) sections."
            $totalFigures = 0
            
            foreach ($sectionElement in $sectionElements) {
                $sectionNo = $sectionElement.getAttribute("sno")
                $sectionSpan = $sectionElement.getElementsByTagName("span") | Where-Object { $_.className -ne "go_fig" } | Select-Object -First 1
                
                if ($sectionSpan) {
                    # Extract the section name, removing any 'Section X' prefix that might already be in the text
                    $sectionText = $sectionSpan.innerText -replace '^Section \d+\s+', ''
                    $sectionName = "Section $sectionNo $sectionText"
                    Write-Host "Processing $sectionName"
                    $sectionNames += $sectionName
                    
                    # Find figures within this section
                    $figureList = $sectionElement.getElementsByTagName("ul") | Select-Object -First 1
                    
                    if ($figureList) {
                        $figureElements = $figureList.getElementsByTagName("li") | Where-Object { $_.getAttribute("figno") }
                        
                        Write-Host "  Found $($figureElements.Count) figures in section $sectionNo."
                        
                        foreach ($figureElement in $figureElements) {
                            $figno = $figureElement.getAttribute("figno")
                            
                            if ($figno) {
                                $figureSpan = $figureElement.getElementsByTagName("span") | Where-Object { $_.className -eq "go_fig" } | Select-Object -First 1
                                
                                if ($figureSpan) {
                                    $figureName = $figureSpan.innerText
                                    # Clean up the figure name by removing the prefix like "2-1 "
                                    $cleanFigureName = $figureName -replace "^\d+-\d+\s+", ""
                                    
                                    # Construct the URL for the figure
                                    $figureUrl = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figno&viewerflag=d&layout=L11"
                                    
                                    # Write to CSV
                                    $csvLine = "$figno,`"$cleanFigureName`",$sectionNo,$msBookNo,$volume,$figureUrl"
                                    $csvLine | Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Append
                                    
                                    Write-Host "    Added figure: $figno - $cleanFigureName"
                                    $totalFigures++
                                }
                            }
                        }
                    } else {
                        Write-Host "  No figures found in section $sectionNo."
                    }
                }
            }
            
            Write-Host "Processed $totalFigures total figures."
            
            # Write section names to text file
            $sectionNames | Out-File -FilePath $sectionNamesPath -Encoding UTF8
            Write-Host "Section names written to $sectionNamesPath"
        } else {
            Write-Host "Warning: Could not find phbk_tree UL element in the HTML."
        }
    } else {
        Write-Host "Warning: Could not find appropriate content div in the HTML."
        Write-Host "Attempting to extract data directly using broader patterns..."
        
        # Try a more aggressive approach to find section and figure data
        # Look for any list items with section numbers
        $sectionPattern = '<li[^>]*sno="(\d+)"[^>]*>.*?<span[^>]*>(.*?)</span>'
        $sectionMatches = [regex]::Matches($htmlContent, $sectionPattern)
        
        if ($sectionMatches.Count -gt 0) {
            Write-Host "Found $($sectionMatches.Count) sections using fallback pattern."
            
            foreach ($sectionMatch in $sectionMatches) {
                $sectionNo = $sectionMatch.Groups[1].Value
                $sectionText = $sectionMatch.Groups[2].Value -replace '^Section \d+\s+', ''
                $sectionName = "Section $sectionNo $sectionText"
                
                Write-Host "Processing $sectionName"
                $sectionNames += $sectionName
                
                # Find figures for this section - look anywhere in the HTML
                $figurePattern = '<li[^>]*figno="' + $sectionNo + '-(\d+)"[^>]*>.*?<span[^>]*>(.*?)</span>'
                $figureMatches = [regex]::Matches($htmlContent, $figurePattern)
                
                Write-Host "  Found $($figureMatches.Count) figures in section $sectionNo."
                
                foreach ($figureMatch in $figureMatches) {
                    $figureNum = $figureMatch.Groups[1].Value
                    $figno = "$sectionNo-$figureNum"
                    $figureName = $figureMatch.Groups[2].Value
                    
                    # Clean up the figure name by removing the prefix like "2-1 "
                    $cleanFigureName = $figureName -replace "^\d+-\d+\s+", ""
                    
                    # Construct the URL for the figure
                    $figureUrl = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figno&viewerflag=d&layout=L11"
                    
                    # Write to CSV
                    $csvLine = "$figno,`"$cleanFigureName`",$sectionNo,$msBookNo,$volume,$figureUrl"
                    $csvLine | Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Append
                    
                    Write-Host "    Added figure: $figno - $cleanFigureName"
                }
            }
            
            # Write section names to text file
            $sectionNames | Out-File -FilePath $sectionNamesPath -Encoding UTF8
        } else {
            Write-Host "Error: Could not find any section data in the HTML."
        }
    }
    
    # Return the paths to the created files
    return @{
        VolumesToUrlPath = $volumesToUrlPath
        SectionNamesPath = $sectionNamesPath
    }
}

# Fallback function that uses regex if DOM methods fail
function Process-HTMLContentWithRegex {
    param (
        [string]$htmlContent,
        [PSObject]$selectedRow,
        [string]$directoryPath
    )
    
    Write-Host "Using regex-based HTML parsing..."
    
    # Get MS Book No and Volume from selectedRow or extract from the HTML if available
    $msBookNo = $selectedRow.'MS Book No'
    $volume = $selectedRow.Volume

    # If we don't have the MS Book No or Volume from selectedRow, try to extract from HTML
    if ([string]::IsNullOrEmpty($msBookNo) -or [string]::IsNullOrEmpty($volume)) {
        $bookTitleMatch = $htmlContent -match '<span[^>]*id="book_title"[^>]*>([^<]+)</span>'
        if ($matches) {
            $bookTitleText = $matches[1].Trim()
            if ($bookTitleText -match "MS(\d+)\s+VOLUME\s+([A-Z])") {
                $msBookNo = $matches[1]
                $volume = $matches[2]
                Write-Host "Extracted from HTML: MS Book No: $msBookNo, Volume: $volume"
            }
        }
    }

    # If we still don't have valid MSBookNo and Volume, exit the function
    if ([string]::IsNullOrEmpty($msBookNo) -or [string]::IsNullOrEmpty($volume)) {
        Write-Host "Error: Could not determine MS Book No and Volume. Exiting function."
        return
    }

    # Create paths for output files
    $volumesToUrlPath = Join-Path -Path $directoryPath -ChildPath "Volumes-to-URL.csv"
    $sectionNamesPath = Join-Path -Path $directoryPath -ChildPath "SectionNames.txt"

    # Create CSV writer for Volumes-to-URL.csv
    $csvHeader = "Figure No.,Name,Section No.,MS Book No,Volume,URL"
    $csvHeader | Out-File -FilePath $volumesToUrlPath -Encoding UTF8
    
    # Initialize array for section names
    $sectionNames = @()

    # Find the div with id="Ryan_fault"
    if ($htmlContent -match '<div\s+id="Ryan_fault"[^>]*>(.*?)</div>') {
        $ryanFaultContent = $matches[1]
        
        # Extract sections
        $sectionPattern = '<li[^>]*sno="(\d+)"[^>]*><div[^>]*></div><span[^>]*>(.*?)</span>'
        $sectionMatches = [regex]::Matches($ryanFaultContent, $sectionPattern)
        
        Write-Host "Found $($sectionMatches.Count) sections."
        
        foreach ($sectionMatch in $sectionMatches) {
            $sectionNo = $sectionMatch.Groups[1].Value
            $sectionText = $sectionMatch.Groups[2].Value -replace '^Section \d+\s+', ''
            $sectionName = "Section $sectionNo $sectionText"
            
            Write-Host "Processing $sectionName"
            $sectionNames += $sectionName
            
            # Find figures for this section
            $figurePattern = '<li[^>]*figno="([^"]+)"[^>]*><span\s+class="go_fig">(.*?)</span>'
            $figureMatches = [regex]::Matches($ryanFaultContent, $figurePattern)
            
            $sectionFigures = $figureMatches | Where-Object { $_.Groups[1].Value -match "^$sectionNo-" }
            Write-Host "  Found $($sectionFigures.Count) figures in section $sectionNo."
            
            foreach ($figureMatch in $sectionFigures) {
                $figno = $figureMatch.Groups[1].Value
                $figureName = $figureMatch.Groups[2].Value
                
                # Clean up the figure name by removing the prefix like "2-1 "
                $cleanFigureName = $figureName -replace "^\d+-\d+\s+", ""
                
                # Construct the URL for the figure
                $figureUrl = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figno&viewerflag=d&layout=L11"
                
                # Write to CSV
                $csvLine = "$figno,`"$cleanFigureName`",$sectionNo,$msBookNo,$volume,$figureUrl"
                $csvLine | Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Append
                
                Write-Host "    Added figure: $figno - $cleanFigureName"
            }
        }
        
        # Write section names to text file
        $sectionNames | Out-File -FilePath $sectionNamesPath -Encoding UTF8
    } else {
        Write-Host "Error: Could not find Ryan_fault div in the HTML."
    }
    
    # Return the paths to the created files
    return @{
        VolumesToUrlPath = $volumesToUrlPath
        SectionNamesPath = $sectionNamesPath
    }
}

# Function to process HTML to CSV
function Process-HTMLToCSV($htmlContent, $htmlFilePath) {
    $csvFilePath = [System.IO.Path]::ChangeExtension($htmlFilePath, '.csv')

    # Load HTML into DOM
    try {
        $htmlDoc = New-Object -ComObject "HTMLFile"
        $htmlDoc.IHTMLDocument2_write($htmlContent)
    } catch {
        $htmlDoc.write([System.Text.Encoding]::UTF8.GetBytes($htmlContent))
    }

    # Find the target table with corrected bordercolor check
    $dataTable = $htmlDoc.getElementsByTagName("TABLE") | 
        Where-Object { 
            $_.border -eq "1" -and 
            $_.cols -eq "5" -and 
            ($_.getAttribute("bordercolor") -in @("#808080", "808080"))
        } | Select-Object -First 1

    if (-not $dataTable) {
        Write-Host "Data table not found."
        return
    }

    # Process rows
    $tableRows = @('"NO.","PART DESCRIPTION","REF.","STOCK NO.","PART NO.","CAGE"')
    for ($i = 2; $i -lt $dataTable.rows.length; $i++) {
        $row = $dataTable.rows[$i]
        $cells = @($row.cells)
        $cleanCells = $cells | ForEach-Object { 
            '"' + ($_.innerText.Trim() -replace '\s+', ' ' -replace '&nbsp;', '') + '"'
        }
        if ($cleanCells -join '' -ne '""""""""""') {
            $tableRows += $cleanCells -join ','
        }
    }

    $tableRows | Out-File $csvFilePath -Encoding UTF8
    Write-Host "CSV created successfully!"
}

################################################################################
#                           File Operations                                    #
################################################################################

# Function to download and process HTML files
function Download-And-Process-HTML($volumesToUrlData, $directoryPath, $currentBook) {
    $HTMLCSVDirectoryPath = Join-Path $directoryPath "HTML and CSV Files"
    if (-not (Test-Path -Path $HTMLCSVDirectoryPath)) {
        New-Item -Path $HTMLCSVDirectoryPath -ItemType Directory | Out-Null
        Write-Host "Created directory: $HTMLCSVDirectoryPath"
    }

    $totalFiles = $volumesToUrlData.Count
    $filesDownloaded = 0

    if ($totalFiles -eq 0) {
        Write-Host "No files to process. Please check the Volumes-to-URL.csv file."
        return
    }

    foreach ($row in $volumesToUrlData) {
        $filesDownloaded++
        $percentComplete = ($filesDownloaded / $totalFiles) * 25
        Update-Progress "Downloading and processing HTML files ($filesDownloaded of $totalFiles)" $percentComplete $currentBook

        try {
            $figureNo = $row.'Figure No.' -replace '[^\w\d-]', '_'
            $htmlFilePath = Join-Path $HTMLCSVDirectoryPath "Figure $figureNo.html"
            
            # Check if URL is provided
            if ([string]::IsNullOrWhiteSpace($row.URL)) {
                Write-Host "Skipping Figure $figureNo - No URL provided"
                continue
            }

            # Fetch webpage and save HTML
            $htmlContent = Invoke-WebRequest -Uri $row.URL -UseBasicParsing
            if ($htmlContent.StatusCode -eq 200) {
                Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8
                Write-Host "Saved HTML file to: $htmlFilePath"

                # Process HTML to CSV
                Process-HTMLToCSV $htmlContent.Content $htmlFilePath
            } else {
                Write-Host "Failed to download Figure $figureNo - Status code: $($htmlContent.StatusCode)"
            }
        } catch {
            Write-Host "Error processing URL for Figure $figureNo : $($row.URL) - $_"
        }
    }

    if ($filesDownloaded -eq 0) {
        Write-Host "No files were downloaded. Please check your internet connection and the URLs in the CSV file."
    } else {
        Write-Host "$filesDownloaded file(s) have been downloaded and processed."
    }
}

# Function to combine CSV files into Section CSVs
function Combine-CSVFiles($sourceDir, $siteCsvPath, $partsBookName) {
    $NewPartsBooksDirectory = Join-Path -Path $sourceDir -ChildPath "CombinedSections"
    if (-Not (Test-Path -Path $NewPartsBooksDirectory)) {
        New-Item -Path $NewPartsBooksDirectory -ItemType Directory | Out-Null
    }

    $csvFiles = Get-ChildItem -Path (Join-Path $sourceDir "HTML and CSV Files") -Filter "Figure *.csv"
    $groupedFiles = $csvFiles | Group-Object { $_.BaseName -replace 'Figure (\d+)-\d+', '$1' }

    $siteData = $null
    if (Test-Path $siteCsvPath) {
        # Read the site data
        $siteData = Import-Csv -Path $siteCsvPath

        # Ensure the "Changed Part (NSN)" column exists in the site data
        if (-not $siteData[0].PSObject.Properties['Changed Part (NSN)']) {
            $siteData | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'Changed Part (NSN)' -Value '' }
        }

        # Add new column for the parts book being created
        $partsBookColumnName = $partsBookName -replace '[^\w\s-]', '' -replace '\s+', ' '
        if (-not $siteData[0].PSObject.Properties[$partsBookColumnName]) {
            $siteData | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name $partsBookColumnName -Value '' }
        }
    } else {
        Write-Host "Site CSV file not found at $siteCsvPath. Proceeding without site data comparison."
    }

    $totalGroups = $groupedFiles.Count
    $processedGroups = 0

    foreach ($group in $groupedFiles) {
        $processedGroups++
        $percentComplete = 25 + ($processedGroups / $totalGroups) * 25  # Allocate 25% of progress to this step
        Update-Progress "Combining CSV files ($processedGroups of $totalGroups)" $percentComplete $partsBookName

        $figureNumber = $group.Name
        $sectionFile = Join-Path -Path $NewPartsBooksDirectory -ChildPath "Section $figureNumber.csv"
        Write-Host "Processing Section $figureNumber"

        $allRows = @()
        foreach ($file in $group.Group) {
            Write-Host "  Processing file: $($file.Name)"
            $csvContent = Import-Csv -Path $file.FullName

            # Add columns if they don't exist
            if (-not $csvContent[0].PSObject.Properties['Location']) {
                $csvContent | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'Location' -Value '' }
            }
            if (-not $csvContent[0].PSObject.Properties['QTY']) {
                $csvContent | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'QTY' -Value '' }
            }

            foreach ($row in $csvContent) {
                if (-not ($row.'STOCK NO.' -eq "" -and $row.'PART NO.' -eq "" -and $row.'CAGE' -eq "" -and $row.'Location' -eq "")) {
                    if (-not $row.'REF.') {
                        $row.'REF.' = $file.Name
                    }

                    if ($siteData) {
                        $stockNo = $row.'STOCK NO.'
                        $partNo = $row.'PART NO.'

                        $siteMatch = $siteData | Where-Object { $_.'Part (NSN)' -eq $stockNo }
                        if ($siteMatch) {
                            $row.'Location' = $siteMatch.'Location'
                            $row.'QTY' = $siteMatch.'QTY'
                            $siteMatchIndex = $siteData.IndexOf($siteMatch)
                            
                            # Extract figure number from the REF. column
                            $refValue = $row.'REF.' -replace '^Figure\s+', ''
                            $figureNumber = $refValue -replace '\.csv$', ''
                            
                            # Update the parts book column with the figure number
                            if ([string]::IsNullOrEmpty($siteData[$siteMatchIndex].$partsBookColumnName)) {
                                $siteData[$siteMatchIndex].$partsBookColumnName = "Figure $figureNumber"
                            } else {
                                $existingRefs = $siteData[$siteMatchIndex].$partsBookColumnName -split '\s*\|\s*'
                                $newRefs = @("Figure $figureNumber")
                                foreach ($ref in $existingRefs) {
                                    if ($ref -notmatch [regex]::Escape("Figure $figureNumber")) {
                                        if ($ref -match '^(See\s+)?Figure\s+\d+-\d+') {
                                            $newRefs += $ref -replace '\.csv$', ''
                                        } else {
                                            $newRefs += $ref
                                        }
                                    }
                                }
                                $siteData[$siteMatchIndex].$partsBookColumnName = $newRefs -join ' | '
                            }
                            Write-Host "Updated $partsBookColumnName for Part (NSN): $($siteMatch.'Part (NSN)') with figure: Figure $figureNumber"
                        } else {
                            $oemMatch = $siteData | Where-Object { $_.'OEM 1' -eq $partNo -or $_.'OEM 2' -eq $partNo -or $_.'OEM 3' -eq $partNo }
                            if ($oemMatch) {
                                $oemMatchIndex = $siteData.IndexOf($oemMatch)
                                $previousPartNSN = $oemMatch.'Part (NSN)'
                                if ($row.'STOCK NO.' -eq "NSL" -or [string]::IsNullOrEmpty($row.'STOCK NO.')) {
                                    $siteData[$oemMatchIndex].'Changed Part (NSN)' = "No standard NSN"
                                } else {
                                    $siteData[$oemMatchIndex].'Changed Part (NSN)' = $previousPartNSN
                                    $siteData[$oemMatchIndex].'Part (NSN)' = $row.'STOCK NO.'
                                }
                                $row.'Location' = $siteData[$oemMatchIndex].'Location'
                                $row.'QTY' = $siteData[$oemMatchIndex].'QTY'
                                
                                # Extract figure number from the REF. column
                                $refValue = $row.'REF.' -replace '^Figure\s+', ''
                                $figureNumber = $refValue -replace '\.csv$', ''
                                
                                # Update the parts book column with the figure number
                                if ([string]::IsNullOrEmpty($siteData[$oemMatchIndex].$partsBookColumnName)) {
                                    $siteData[$oemMatchIndex].$partsBookColumnName = "Figure $figureNumber"
                                } else {
                                    $existingRefs = $siteData[$oemMatchIndex].$partsBookColumnName -split '\s*\|\s*'
                                    $newRefs = @("Figure $figureNumber")
                                    foreach ($ref in $existingRefs) {
                                        if ($ref -notmatch [regex]::Escape("Figure $figureNumber")) {
                                            if ($ref -match '^(See\s+)?Figure\s+\d+-\d+') {
                                                $newRefs += $ref -replace '\.csv$', ''
                                            } else {
                                                $newRefs += $ref
                                            }
                                        }
                                    }
                                    $siteData[$oemMatchIndex].$partsBookColumnName = $newRefs -join ' | '
                                }
                                Write-Host "Updated $partsBookColumnName for OEM Part: $partNo with figure: Figure $figureNumber"
                            } else {
                                $row.'Location' = "Not Stocked Locally"
                            }
                        }
                    } else {
                        $row.'Location' = "Site data not available"
                    }

                    $allRows += $row
                }
            }
            Write-Host "  Processed $($csvContent.Count) rows from $($file.Name)"
        }

        $allRows | Export-Csv -Path $sectionFile -NoTypeInformation
        Write-Host "Saved combined section to: $sectionFile"
    }

    if ($siteData -and (Test-Path $siteCsvPath)) {
        # Save the updated Site CSV
        $siteData | Export-Csv -Path $siteCsvPath -NoTypeInformation
        Write-Host "Updated Site CSV: $siteCsvPath"

        # Update the Excel file
        $siteName = [System.IO.Path]::GetFileNameWithoutExtension($siteCsvPath)
        $partsRoomDir = Split-Path $siteCsvPath -Parent
    }

    Write-Host "CSV files combined and saved to $NewPartsBooksDirectory"
    return $NewPartsBooksDirectory
}

################################################################################
#                          Excel Operations                                    #
################################################################################

function Create-ExcelWorkbook($sourceDir, $combinedCsvDir) {
    $startTime = Get-Date
    Write-Host "Starting Excel workbook creation for $sourceDir"
    
    $excelWorkbookPath = Join-Path $sourceDir "$((Split-Path $sourceDir -Leaf)).xlsx"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()

    try {
        # Remove default sheets
        while ($workbook.Sheets.Count -gt 1) {
            $workbook.Sheets.Item(1).Delete()
        }

        # Verify combined CSV directory
        if (-not (Test-Path $combinedCsvDir)) {
            Write-Host "Combined CSV directory not found: $combinedCsvDir"
            throw "Combined CSV directory not found"
        }

        # Get CSV files and sort them
        Update-Progress "Finding section CSV files..." 50 (Split-Path $sourceDir -Leaf)
        $sectionCsvFiles = Get-ChildItem -Path $combinedCsvDir -Filter "Section *.csv" -ErrorAction SilentlyContinue
        
        if ($null -eq $sectionCsvFiles -or $sectionCsvFiles.Count -eq 0) {
            Write-Host "No Section CSV files found in $combinedCsvDir"
            throw "No Section CSV files found"
        }

        # Safer sorting with error handling
        $sectionCsvFiles = $sectionCsvFiles | Sort-Object { 
            if ($_.BaseName -match 'Section (\d+)') {
                return [int]$matches[1]
            } else {
                return 0
            }
        } -Descending
        
        Write-Host "Found $($sectionCsvFiles.Count) section CSV files"
        
        $csvLoadTime = (Get-Date) - $startTime
        Write-Host "CSV files located in $($csvLoadTime.TotalSeconds) seconds"

        $totalSheets = [int]$sectionCsvFiles.Count
        $processedSheets = 0
        
        # Timing data
        $sheetTimes = @()

        foreach ($file in $sectionCsvFiles) {
            $sheetStartTime = Get-Date
            $processedSheets = [int]$processedSheets + 1
            $baseProgress = [double]50  # Starting point for this phase
            $progressWeight = [double]25  # Weight of this phase as percentage
            $sheetProgress = [double]$processedSheets / [double]$totalSheets
            $percentComplete = [double]$baseProgress + ([double]$sheetProgress * [double]$progressWeight)
            
            $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            Update-Progress "Loading CSV for sheet: $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)

            # Load CSV content
            $csvStart = Get-Date
            $csvContent = Import-Csv -Path $file.FullName
            $csvEnd = Get-Date
            $csvDuration = ($csvEnd - $csvStart).TotalSeconds
            
            # Create worksheet
            Update-Progress "Creating worksheet: $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
            $worksheet = $workbook.Sheets.Add()
            $worksheet.Name = $worksheetName

            # Define column order
            $orderedHeaders = @('NO.', 'STOCK NO.', 'PART DESCRIPTION', 'PART NO.', 'REF.', 'QTY', 'LOCATION', 'CAGE')

            # Write headers
            for ($i = 0; $i -lt $orderedHeaders.Count; $i++) {
                $worksheet.Cells.Item(1, $i + 1) = $orderedHeaders[$i]
            }

            # Prepare 2D array for data
            $rowCount = if ($null -eq $csvContent) { 0 } else { $csvContent.Count }
            Update-Progress "Preparing data for $rowCount rows in $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
            
            # Optimize for larger data sets - use array approach
            if ($rowCount -gt 0) {
                $dataArray = New-Object 'object[,]' $rowCount, $orderedHeaders.Count
                
                # Fill the array with data
                $totalRows = [int]$rowCount
                $rowsProcessed = 0
                $dataStart = Get-Date
                
                foreach ($rowItem in $csvContent) {
                    $rowsProcessed = [int]$rowsProcessed + 1
                    # Calculate percentage within this sheet
                    $rowProgress = [double]$rowsProcessed / [double]$totalRows
                    # Calculate sheet contribution to overall progress
                    $sheetContribution = [double]$rowProgress / [double]$totalSheets
                    # Calculate final percentage
                    $detailedProgress = [double]$baseProgress + ([double]$sheetProgress * [double]$progressWeight) + ([double]$sheetContribution * [double]$progressWeight)
                    
                    if ($rowsProcessed % 50 -eq 0 -or $rowsProcessed -eq $totalRows) {  # Update progress every 50 rows
                        Update-Progress "Processing sheet ${worksheetName}: row $rowsProcessed of $totalRows" [Math]::Min([int]$detailedProgress, 100) (Split-Path $sourceDir -Leaf)
                    }
                    
                    # Fill array row
                    for ($i = 0; $i -lt $orderedHeaders.Count; $i++) {
                        $header = $orderedHeaders[$i]
                        $cellValue = $rowItem.$header
                        
                        # Special handling for REF. column with hyperlinks
                        if ($header -eq 'REF.' -and $cellValue -match 'Figure \d+-\d+') {
                            $dataArray[$rowsProcessed-1, $i] = $cellValue -replace '\.csv$', ''
                        } else {
                            $dataArray[$rowsProcessed-1, $i] = $cellValue
                        }
                    }
                }
                $dataEnd = Get-Date
                $dataDuration = ($dataEnd - $dataStart).TotalSeconds
                
                # Calculate position for data range
                $startRow = 2  # Start after header row
                $startCol = 1
                $endRow = $startRow + $rowCount - 1
                $endCol = $startCol + $orderedHeaders.Count - 1
                
                # Write data to worksheet in one operation
                Update-Progress "Writing data to worksheet $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
                $writeStart = Get-Date
                
                $startCell = $worksheet.Cells.Item($startRow, $startCol)
                $endCell = $worksheet.Cells.Item($endRow, $endCol)
                $dataRange = $worksheet.Range($startCell, $endCell)
                $dataRange.Value2 = $dataArray
                
                $writeEnd = Get-Date
                $writeDuration = ($writeEnd - $writeStart).TotalSeconds
                
                # Add hyperlinks for REF. column
                Update-Progress "Adding hyperlinks to worksheet $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
                $linkStart = Get-Date
                
                # Find REF. column index
                $refColIndex = -1
                for ($i = 0; $i -lt $orderedHeaders.Count; $i++) {
                    if ($orderedHeaders[$i] -eq 'REF.') {
                        $refColIndex = $i + 1  # Excel columns are 1-based
                        break
                    }
                }
                
                if ($refColIndex -gt 0) {
                    for ($row = 2; $row -le $rowCount + 1; $row++) {
                        $cellValue = $worksheet.Cells.Item($row, $refColIndex).Value2
                        if ($cellValue -and $cellValue -match 'Figure \d+-\d+') {
                            $figureNumber = $cellValue
                            $htmlFileName = "$figureNumber.html"
                            $htmlFilePath = Join-Path -Path $sourceDir -ChildPath "HTML and CSV Files\$htmlFileName"
                            
                            if (Test-Path $htmlFilePath) {
                                $cell = $worksheet.Cells.Item($row, $refColIndex)
                                $worksheet.Hyperlinks.Add($cell, $htmlFilePath, "", "", $figureNumber) | Out-Null
                            }
                            
                            # Update progress occasionally to keep UI responsive
                            if (($row - 2) % 50 -eq 0 -or $row -eq $rowCount + 1) {
                                $linkProgress = [double]($row - 2) / [double]$rowCount
                                $linkPercent = [double]$baseProgress + ([double]$sheetProgress * [double]$progressWeight) + ([double]$linkProgress * [double]$progressWeight * 0.2)
                                Update-Progress "Adding hyperlinks: $($row-1) of $rowCount in $worksheetName" [Math]::Min([int]$linkPercent, 100) (Split-Path $sourceDir -Leaf)
                            }
                        }
                    }
                }
                
                $linkEnd = Get-Date
                $linkDuration = ($linkEnd - $linkStart).TotalSeconds
            } else {
                $dataDuration = 0
                $writeDuration = 0
                $linkDuration = 0
            }
            
            # Format as table
            Update-Progress "Formatting table for $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
            $formatStart = Get-Date
            
            $range = $worksheet.UsedRange
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $range, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
            $listObject.Name = "$($worksheet.Name)Table"
            $listObject.TableStyle = "TableStyleMedium2"
            
            $formatEnd = Get-Date
            $formatDuration = ($formatEnd - $formatStart).TotalSeconds
            
            # Format columns
            Update-Progress "Formatting columns for $worksheetName" $percentComplete (Split-Path $sourceDir -Leaf)
            $columnStart = Get-Date
            
            foreach ($column in $listObject.ListColumns) {
                $column.Range.EntireColumn.AutoFit()
                $column.Range.VerticalAlignment = -4108 # xlCenter
                if ($column.Name -eq "PART DESCRIPTION") {
                    $column.Range.Cells(1, 1).HorizontalAlignment = -4108 # xlCenter
                    $column.Range.Offset(1, 0).HorizontalAlignment = -4131 # xlLeft
                } else {
                    $column.Range.HorizontalAlignment = -4108 # xlCenter
                }
            }
            
            $columnEnd = Get-Date
            $columnDuration = ($columnEnd - $columnStart).TotalSeconds
            
            # Record timing data for this sheet
            $sheetEndTime = Get-Date
            $sheetDuration = ($sheetEndTime - $sheetStartTime).TotalSeconds
            $sheetTimes += [PSCustomObject]@{
                Worksheet = $worksheetName
                Rows = $rowCount
                TotalTime = $sheetDuration
                CsvLoad = $csvDuration
                DataPrep = $dataDuration
                WriteData = $writeDuration
                Hyperlinks = $linkDuration
                TableFormat = $formatDuration
                ColumnFormat = $columnDuration
            }
            
            Write-Host "Worksheet $worksheetName created in $sheetDuration seconds ($rowCount rows)"
        }
        
        # Save the workbook
        Update-Progress "Saving Excel workbook..." 95 (Split-Path $sourceDir -Leaf)
        $saveStart = Get-Date
        $workbook.SaveAs($excelWorkbookPath)
        $saveEnd = Get-Date
        $saveDuration = ($saveEnd - $saveStart).TotalSeconds
        
        # Log all timing data
        Write-Host "Excel workbook created and saved to $excelWorkbookPath in $saveDuration seconds"
        
        $endTime = Get-Date
        $totalDuration = ($endTime - $startTime).TotalSeconds
        Write-Host "Total Excel workbook creation time: $totalDuration seconds"
        Write-Host "Sheet timing details:"
        
        # Fix for the Write-Host error - separate the operations
        $formattedOutput = $sheetTimes | Format-Table -AutoSize | Out-String
        Write-Host $formattedOutput
        
        return $true
    }
    catch {
        Write-Host "Error in Create-ExcelWorkbook: $_"
        throw $_
    }
    finally {
        Update-Progress "Cleaning up Excel resources..." 100 (Split-Path $sourceDir -Leaf)
        if ($workbook) {
            try { $workbook.Close($false) } catch { Write-Host "Error closing workbook: $_" }
        }
        if ($excel) {
            try { $excel.Quit() } catch { Write-Host "Error quitting Excel: $_" }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to rename worksheets with improved reliability
function Rename-Worksheets {
    param (
        [string]$excelFilePath,
        [hashtable]$sectionNames
    )
    
    Write-Host "Starting worksheet renaming for $excelFilePath..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelFilePath)

    # Create a mapping table that translates between truncated and full names
    $nameMapping = @{}
    
    # Generate truncated versions of all section names for matching
    foreach ($sectionNum in $sectionNames.Keys) {
        $fullSectionName = $sectionNames[$sectionNum]
        $truncatedName = $fullSectionName.Substring(0, [Math]::Min(31, $fullSectionName.Length)) -replace '[:\\/?*\[\]]', ''
        $nameMapping[$truncatedName] = $fullSectionName
        # Also create a mapping from the section number format
        if ($fullSectionName -match '^(Section \d+)') {
            $sectionNumberOnly = $matches[1]
            $nameMapping[$sectionNumberOnly] = $fullSectionName
        }
    }
    
    $totalSheets = $workbook.Sheets.Count
    $processedSheets = 0

    foreach ($sheet in $workbook.Sheets) {
        $processedSheets++
        $percentComplete = 75 + ($processedSheets / $totalSheets) * 25
        Update-Progress "Renaming worksheets ($processedSheets of $totalSheets)" $percentComplete $currentBook

        $currentName = $sheet.Name
        
        # Skip Sheet1
        if ($currentName -eq "Sheet1") {
            Write-Host "Deleting default Sheet1"
            $sheet.Delete()
            continue
        }
        
        $renamed = $false
        
        # Check exact matches first
        if ($sectionNames.ContainsKey($currentName)) {
            $newName = $sectionNames[$currentName]
            $newName = $newName.Substring(0, [Math]::Min(31, $newName.Length)) -replace '[:\\/?*\[\]]', ''
            if (![string]::IsNullOrWhiteSpace($newName)) {
                $sheet.Name = $newName
                Write-Host "Renamed '$currentName' to '$newName' (direct match)"
                $renamed = $true
            }
        }
        
        # If not renamed yet, check against truncated versions
        if (-not $renamed -and $nameMapping.ContainsKey($currentName)) {
            $fullName = $nameMapping[$currentName]
            $newName = $fullName.Substring(0, [Math]::Min(31, $fullName.Length)) -replace '[:\\/?*\[\]]', ''
            if (![string]::IsNullOrWhiteSpace($newName)) {
                $sheet.Name = $newName
                Write-Host "Renamed '$currentName' to '$newName' (mapping match)"
                $renamed = $true
            }
        }
        
        # Try matching just the section number portion
        if (-not $renamed -and $currentName -match '^Section (\d+)') {
            $sectionPrefix = $matches[0]
            # Find any section that starts with this prefix
            $matchingSections = $sectionNames.Values | Where-Object { $_ -match "^$sectionPrefix" }
            if ($matchingSections.Count -gt 0) {
                $newName = $matchingSections[0].Substring(0, [Math]::Min(31, $matchingSections[0].Length)) -replace '[:\\/?*\[\]]', ''
                if (![string]::IsNullOrWhiteSpace($newName)) {
                    $sheet.Name = $newName
                    Write-Host "Renamed '$currentName' to '$newName' (prefix match)"
                    $renamed = $true
                }
            }
        }
        
        if (-not $renamed) {
            Write-Host "No matching section name found for sheet '$currentName', skipping rename."
        }
    }

    Write-Host "Saving workbook after renaming worksheets..."
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Clean up COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Worksheet renaming completed."
}

# Function to create and format Excel file from CSV
function Create-ExcelFromCsv {
    param(
        [string]$siteName,
        [string]$csvDirectory,
        [string]$excelDirectory,
        [string]$tableName = "My_Parts_Room"
    )
    try {
        Write-Log "Starting to create Excel file from CSV for $siteName..."
        
        # Add timing diagnostics
        $startTime = Get-Date
        
        # Show progress form if not already visible
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Creating Parts Room Excel"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = 'CenterScreen'
        
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(360,20)
        $progressBar.Location = New-Object System.Drawing.Point(10,10)
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
        $rowCount = if ($null -eq $csvData) { 0 } else { $csvData.Count }
        $colCount = if ($null -eq $headers) { 0 } else { $headers.Count }
        
        if ($rowCount -gt 0 -and $colCount -gt 0) {
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
                    $percent = [double]30 + ([double]$rowIdx / [double]$rowCount * [double]20)  # Scale from 30% to 50%
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
        }
        
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
        $descriptionColumn = $null
        try {
            $descriptionColumn = $listObject.ListColumns | Where-Object { $_.Name -eq "Description" }
            if ($descriptionColumn) {
                $descriptionColumn.Range.Offset(1, 0).HorizontalAlignment = -4131 # xlLeft
            }
        } catch {
            Write-Log "Error while formatting Description column: $_"
        }

        # Remove any columns named "Importing data..."
        try {
            for ($col = $headers.Count; $col -ge 1; $col--) {
                $columnHeader = $worksheet.Cells.Item(1, $col).Value2
                if ($columnHeader -eq "Importing data...") {
                    $column = $worksheet.Columns.Item($col)
                    $column.Delete()
                    Write-Log "Removed 'Importing data...' column"
                }
            }
        } catch {
            Write-Log "Error while removing columns: $_"
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
        # Clean up even if there's an error
        if ($excel) {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch {
                Write-Log "Error releasing Excel COM object: $_"
            }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

################################################################################
#                           Main Execution                                     #
################################################################################

# Main execution
$existingBooks = Get-ExistingBooks
Write-Host "Existing books: $($existingBooks.Keys -join ', ')"

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select Handbooks'
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(780, 20)
$label.Text = 'Please select one or more handbooks:'
$form.Controls.Add($label)

# Create a ListView instead of a DataGridView
$global:listView.Location = New-Object System.Drawing.Point(10,40)
$global:listView.Size = New-Object System.Drawing.Size(760, 450)
$global:listView.View = [System.Windows.Forms.View]::Details
$global:listView.FullRowSelect = $true
$global:listView.CheckBoxes = $true
$form.Controls.Add($global:listView)

# Add columns to the ListView
$global:listView.Columns.Add("Full Name", 300)
$global:listView.Columns.Add("MS Book No", 100)
$global:listView.Columns.Add("Volume", 100)

Write-Host "ListView items count: $($global:listView.Items.Count)"
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10, 500)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = 'Process Selected'
$form.Controls.Add($button)

# Button click event
$button.Add_Click({
    $selectedItems = $global:listView.CheckedItems

    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one handbook.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $handbooks = @()

    foreach ($item in $selectedItems) {
        $selectedName = $item.Text
        $selectedRow = $csvData | Where-Object { $_.'Full Name' -eq $selectedName }
        $url = "https://www1.mtsc.usps.gov/apps/phbk/index.php?msbookno=$($selectedRow.'MS Book No')&volno=$($selectedRow.Volume)"
        
        $handbooks += @{
            Name = $selectedName
            Row = $selectedRow
            Url = $url
        }

        Start-Process $url
    }

    $form.Close()

    # Show the progress form at the start of processing
    $global:progressForm.Show()
    $global:progressForm.Refresh()

    # Collect HTML content for all handbooks
    $handbooksWithContent = @()
    foreach ($handbook in $handbooks) {
        $htmlContent = Get-VerifiedHtmlContent -handbookName $handbook.Name
        $handbooksWithContent += @{
            Name = $handbook.Name
            Row = $handbook.Row
            HtmlContent = $htmlContent
        }
    }

    $newBooks = @{}
    # Process each handbook
    foreach ($handbook in $handbooksWithContent) {
        $selectedNameSafe = Sanitize-Name $handbook.Name
        $directoryPath = Join-Path -Path $partsBooksDirPath -ChildPath $selectedNameSafe

        # Create directory for the new parts book
        if (-not (Test-Path $directoryPath)) {
            New-Item -ItemType Directory -Path $directoryPath | Out-Null
        }
        Write-Host "Processing: $($handbook.Name)"
        Write-Host "Directory created/updated: $directoryPath"

        # Process HTML content
        $processedContent = Process-HTMLContent -htmlContent $handbook.HtmlContent -selectedRow $handbook.Row -directoryPath $directoryPath

        # Update config with new book information
        $newBooks[$selectedNameSafe] = @{
            VolumesToUrlCsvPath = $processedContent.VolumesToUrlPath
            SectionNamesCsvPath = $processedContent.SectionNamesPath
        }

        # Download and process HTML files
        $volumesToUrlData = Import-Csv -Path $processedContent.VolumesToUrlPath
        Download-And-Process-HTML -volumesToUrlData $volumesToUrlData -directoryPath $directoryPath -currentBook $handbook.Name

        # Determine the site CSV path
        $partsRoomDir = $config.PartsRoomDirectory
        $siteCsvPath = Get-ChildItem -Path $partsRoomDir -Filter "*.csv" | Select-Object -First 1 -ExpandProperty FullName
        
        if (-not $siteCsvPath) {
            Write-Host "No CSV file found in the Parts Room directory. Proceeding without site data comparison."
            $siteCsvPath = $null
        } else {
            Write-Host "Using site CSV file: $siteCsvPath"
        }
        
        # Combine CSV files
        $combinedCsvDir = Combine-CSVFiles -sourceDir $directoryPath -siteCsvPath $siteCsvPath -partsBookName $selectedNameSafe

        # Create Excel workbook
        Create-ExcelWorkbook -sourceDir $directoryPath -combinedCsvDir $combinedCsvDir

        # Rename worksheets using extracted section names
        $excelFilePath = Join-Path $directoryPath "$selectedNameSafe.xlsx"
        $sectionNamesContent = Get-Content -Path $processedContent.SectionNamesPath
        $sectionNames = @{}
        foreach ($line in $sectionNamesContent) {
            if ($line -match '^(Section \d+) (.+)$') {
                $sectionNames[$matches[1]] = $line
            }
        }
        Rename-Worksheets -excelFilePath $excelFilePath -sectionNames $sectionNames

        Write-Host "Completed processing: $($handbook.Name)"
    }

    # Update config with all books
    Update-ConfigBooks $newBooks

    # After processing all handbooks, create the Excel file for the parts room
    $partsRoomDir = $config.PartsRoomDirectory
    $siteCsvPath = Get-ChildItem -Path $partsRoomDir -Filter "*.csv" | Select-Object -First 1 -ExpandProperty FullName

    if ($siteCsvPath) {
        $siteName = [System.IO.Path]::GetFileNameWithoutExtension($siteCsvPath)
        Create-ExcelFromCsv -siteName $siteName -csvDirectory $partsRoomDir -excelDirectory $partsRoomDir
        Write-Host "Created Excel file for site data"
    } else {
        Write-Host "No CSV file found in the Parts Room directory. Excel file not created."
    }

    # Close progress form
    $global:progressForm.Close()

    Write-Host "All selected handbooks have been processed successfully!"
})

# Show the form
$form.ShowDialog()
