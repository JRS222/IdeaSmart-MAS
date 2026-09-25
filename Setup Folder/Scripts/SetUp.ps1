################################################################################
#                                                                              #
#                          Parts Book Manager Setup                            #
#                                                                              #
################################################################################

################################################################################
#                          Required .NET Assemblies                            #
################################################################################

# Load required assemblies for the Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

################################################################################
#                           Global Variables                                   #
################################################################################

# Configuration variables
$currentDir = $PSScriptRoot         #Current script directory
$setupFolder = Split-Path -Parent $PSScriptRoot   #Parent directory
$config = $null                     #Configuration object

# File paths
$configPath = $null                 # Path to configuration file
$logPath = "setup_log.log"          # Path to log file

################################################################################
#                            Core Utilities                                    #
################################################################################

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp - $message"
    # Optionally, you can also write to a log file:
    "$timestamp - $message" | Out-File -Append -FilePath "setup_log.log"
}

# Function to show a folder browser dialog
function Show-FolderBrowserDialog {
    param([string]$Description)
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

################################################################################
#                        Configuration Functions                               #
################################################################################

# Function to set up the initial configuration
function Set-InitialConfiguration {
    Write-Log "Starting Set-InitialConfiguration"

    # --- Locate script directory robustly ---
    $currentDir = if ($PSScriptRoot) {
        $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        (Get-Location).Path
    }
    $setupFolder = Split-Path -Parent $currentDir

    Write-Log "Current Directory: $currentDir"
    Write-Log "Setup Folder: $setupFolder"

    # --- Ask where the installation root should live ---
    $parentDir = Show-FolderBrowserDialog -Description "Select the parent directory where PartsBookManagerRootDirectory will be created"
    if (-not $parentDir) {
        Write-Log "Setup cancelled by user."
        return $null
    }

    $rootDir = Join-Path $parentDir "PartsBookManagerRootDirectory"

    if (Test-Path $rootDir) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "A PartsBookManagerRootDirectory already exists at:`n`n$rootDir`n`n" +
            "Setup creates fresh installations. To modify an existing install, " +
            "use the app's Actions tab instead.`n`nAbort setup?",
            "Existing Installation Detected",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Setup aborted: root exists at $rootDir"
            return $null
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Overwriting will DELETE the existing Config.json, Call Logs, Labor Logs, " +
            "and Parts Books index. Continue?",
            "Confirm Overwrite",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Setup aborted at overwrite confirmation."
            return $null
        }

        Write-Log "User chose to overwrite existing root at $rootDir"
        Remove-Item -Path $rootDir -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $rootDir | Out-Null
    Write-Log "Created Root Directory at $rootDir"

    # --- Resolve every directory path up front ---
    $partsBooksDir   = Join-Path $rootDir "Parts Books"
    $scriptsDir      = Join-Path $rootDir "Scripts"
    $dropdownCsvsDir = Join-Path $rootDir "Dropdown CSVs"
    $partsRoomDir    = Join-Path $rootDir "Parts Room"
    $laborDir        = Join-Path $rootDir "Labor"
    $callLogsDir     = Join-Path $rootDir "Call Logs"

    foreach ($d in @($partsBooksDir, $scriptsDir, $dropdownCsvsDir, $partsRoomDir, $laborDir, $callLogsDir)) {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
        Write-Log "Created directory: $d"
    }

    $sameDayPartsRoomDir = Join-Path $partsRoomDir "Same Day Parts Room"
    New-Item -ItemType Directory -Force -Path $sameDayPartsRoomDir | Out-Null
    Write-Log "Created Same Day Parts Room directory: $sameDayPartsRoomDir"

    # --- Seed prerequisite CSVs with headers so Import-Csv doesn't blow up on empty files ---
    $callLogsPath  = Join-Path $callLogsDir     "CallLogs.csv"
    $laborLogsPath = Join-Path $laborDir        "LaborLogs.csv"
    $machinesPath  = Join-Path $dropdownCsvsDir "Machines.csv"

    $seedFiles = @(
        @{ Path = $callLogsPath;  Headers = "Date,Machine,Cause,Action,Noun,Time Down,Time Up,Notes" },
        @{ Path = $laborLogsPath; Headers = "Date,Work Order,Description,Machine,Duration,Notes,Parts" },
        @{ Path = $machinesPath;  Headers = "Machine Acronym,Machine Number" }
    )

    foreach ($f in $seedFiles) {
        if (-not (Test-Path $f.Path)) {
            $f.Headers | Out-File -FilePath $f.Path -Encoding UTF8
            Write-Log "Created file with headers: $($f.Path)"
        }
    }

    # --- Build the config object ---
    #
    # This schema is the contract with UI-Script.ps1. Every key here is
    # read by the UI at startup. Do not rename without updating the UI.
    #
    # Use [ordered]@{} so ConvertTo-Json preserves the key order (nice
    # for diffing and for humans reading the file).
    #
    # Notes on the two "state" keys:
    #   Books             -> hashtable, empty by default. Serializes to {}
    #                        The UI iterates it via .PSObject.Properties.
    #   SameDayPartsRooms -> ARRAY, empty by default. Serializes to [].
    #                        The UI uses .Count, +=, and foreach directly,
    #                        all of which break on a hashtable.
    #
    $config = [ordered]@{
        # Identity / paths
        RootDirectory         = $rootDir
        SetupFolder           = $setupFolder
        ScriptsDirectory      = $scriptsDir
        DropdownCsvsDirectory = $dropdownCsvsDir
        LaborDirectory        = $laborDir
        CallLogsDirectory     = $callLogsDir
        PartsRoomDirectory    = $partsRoomDir
        PartsBooksDirectory   = $partsBooksDir

        # Runtime state, populated during setup and later by the UI
        Books                 = @{}
        SameDayPartsRooms     = @()

        # Named prerequisite files. The UI reads these directly:
        #   $config.PrerequisiteFiles.LaborLogs
        #   $config.PrerequisiteFiles.CallLogs
        #   $config.PrerequisiteFiles.Machines
        PrerequisiteFiles     = [ordered]@{
            LaborLogs = $laborLogsPath
            CallLogs  = $callLogsPath
            Machines  = $machinesPath
        }

        # User-configurable
        SupervisorEmail       = "default@example.com"
    }

    # --- Persist to the Scripts directory, which is where the UI looks ---
    $configFilePath = Join-Path $scriptsDir "Config.json"
    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $configFilePath -Encoding UTF8
    Write-Log "Configuration saved at $configFilePath"

    return $config
}

################################################################################
#                          Setup Operations                                    #
################################################################################

# Function to copy setup files to the necessary directories
function Copy-SetupFiles {
    param($config)
    Write-Log "Starting Copy-SetupFiles"
    
    $setupDir = $config.SetupFolder
    Write-Log "Setup Directory: $setupDir"

    # Copy Scripts
    Write-Log "Copying Scripts..."
    $scriptsSourceDir = Join-Path $setupDir "Scripts"
    if (Test-Path $scriptsSourceDir) {
        Copy-Item -Path "$scriptsSourceDir\*" -Destination $config.ScriptsDirectory -Recurse -Force
        Write-Log "Copied Scripts"
    } else {
        Write-Log "Warning: Scripts directory not found in the SetupFolder: $scriptsSourceDir"
    }

    # Copy Dropdown CSVs
    $csvFiles = @("sites.csv", "Parsed-Parts-Volumes.csv", "Causes.csv", "Actions.csv", "Nouns.csv")
    foreach ($csvFile in $csvFiles) {
        $sourcePath = Join-Path $setupDir $csvFile
        $destPath = Join-Path $config.DropdownCsvsDirectory $csvFile
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-Log "Copied $csvFile to $destPath"
            # Always add to PrerequisiteFiles when successfully copied
            $config.PrerequisiteFiles[$csvFile -replace "\.csv", ""] = $destPath
        } else {
            Write-Log "Warning: $csvFile not found in setup directory: $sourcePath"
            if ($csvFile -in @("sites.csv", "Parsed-Parts-Volumes.csv")) {
                New-Item -ItemType File -Path $destPath -Force | Out-Null
                Write-Log "Created empty file: $destPath"
                $config.PrerequisiteFiles[$csvFile -replace "\.csv", ""] = $destPath
            }
        }
    }

    # Save configuration in the Scripts directory
    $configFilePath = Join-Path $config.ScriptsDirectory "Config.json"
    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $configFilePath -Encoding UTF8
    Write-Log "Configuration saved at $configFilePath"

    [System.Windows.Forms.MessageBox]::Show(
        "Setup files have been copied successfully.",
        "Files Copied",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to parse and export HTML data into CSV format using DOM
function Set-PartsRoom {
    param($config)

    # Load the sites from the CSV
    $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
    Write-Host "Loading Sites from CSV at $sitesPath..."
    $sites = Import-Csv -Path $sitesPath

    # Check if the CSV was loaded successfully
    if ($sites.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data found in the Sites.csv file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "No data found in the Sites.csv file. Exiting..."
        return
    }

    # Determine the correct column names
    $SiteIDColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
    $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
    Write-Host "Site ID Column: $SiteIDColumn, Full Name Column: $fullNameColumn"

    # Create the form for site selection
    Write-Host "Creating form for site selection..."
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
    Write-Host "Dropdown populated with site names."
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
    Write-Host "Showing form for site selection..."
    $form.ShowDialog()
    $selectedSite = $form.Tag
    
    if ($selectedSite) {
        Write-Host "Selected Site: $selectedSite"
        $selectedRow = $sites | Where-Object { $_.$fullNameColumn -eq $selectedSite }
        $url = "http://emarssu3.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$($selectedRow.$SiteIDColumn)&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

        # Download HTML content
        Write-Host "Downloading HTML content for $selectedSite..."
        try {
            $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
            $htmlFilePath = Join-Path $config.PartsRoomDirectory "$selectedSite.html"
            Write-Host "Saving HTML content to $htmlFilePath..."
            Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Downloaded HTML for $selectedSite", "Download Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            $errorMessage = "Failed to download HTML: $($_.Exception.Message)"
            Write-Host $errorMessage -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        Write-Host "Processing the downloaded HTML file for $selectedSite using DOM..."

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

            Write-Host "HTML document loaded into DOM successfully."
            
            # Find the main table with the parts data
            $tables = $htmlDoc.getElementsByTagName("table")
            $mainTable = $null
            
            Write-Host "Found $($tables.length) tables in the document."
            
            # Try multiple methods to find the main table
            foreach ($table in $tables) {
                # Try with className
                if ($table.className -eq "MAIN") {
                    $mainTable = $table
                    Write-Host "Found main table using className property."
                    break
                }
                
                # Try with getAttribute
                try {
                    if ($table.getAttribute("class") -eq "MAIN") {
                        $mainTable = $table
                        Write-Host "Found main table using getAttribute method."
                        break
                    }
                } catch {
                    # Ignore errors with getAttribute
                }
                
                # Check if it's a wide table with borders and multiple columns
                try {
                    if ($table.border -eq "1" -and $table.summary -match "stock") {
                        $mainTable = $table
                        Write-Host "Found main table by border and summary attributes."
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
                            Write-Host "Found table with $($headerRow.cells.length) columns, using as main table."
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
            Write-Host "Number of rows found: $($rows.length)"

            # Create progress bar form
            $progressForm = New-Object System.Windows.Forms.Form
            $progressForm.Text = "Processing Parts Data"
            $progressForm.Size = New-Object System.Drawing.Size(400, 100)
            $progressForm.StartPosition = 'CenterScreen'
            $progressForm.FormBorderStyle = 'FixedDialog'
            $progressForm.MaximizeBox = $false
            $progressForm.MinimizeBox = $false

            $progressBar = New-Object System.Windows.Forms.ProgressBar
            $progressBar.Size = New-Object System.Drawing.Size(360, 20)
            $progressBar.Location = New-Object System.Drawing.Point(10, 20)
            $progressBar.Minimum = 0
            $progressBar.Maximum = $rows.length - 1  # Skip header row
            $progressForm.Controls.Add($progressBar)

            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Size = New-Object System.Drawing.Size(360, 20)
            $statusLabel.Location = New-Object System.Drawing.Point(10, 50)
            $statusLabel.Text = "Processing data..."
            $progressForm.Controls.Add($statusLabel)

            # Show the progress form
            $progressForm.Show()
            $progressForm.Refresh()
            
            # Initialize array to hold parsed data
            $parsedData = @()
            
            # Process each row (skip first row which is the header)
            for ($i = 1; $i -lt $rows.length; $i++) {
                try {
                    # Update progress bar
                    $progressBar.Value = $i
                    $progressForm.Refresh()
                    [System.Windows.Forms.Application]::DoEvents()
                    $statusLabel.Text = "Processing row $i of $($rows.length - 1)..."
                    
                    $row = $rows.item($i)
                    if ($null -eq $row) {
                        # Continue processing instead of writing to host
                        continue
                    }
                    
                    # Skip rows that don't have the MAIN class or enough cells
                    $rowClass = try { $row.className } catch { "" }
                    if ($rowClass -ne "MAIN" -and $rowClass -ne "HILITE") {
                        continue
                    }
                    
                    # Get all cells in the row
                    $cells = $row.getElementsByTagName("td")
                    
                    if ($null -eq $cells -or $cells.length -lt 6) {
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
                                        # Ignore errors on OEM processing
                                    }
                                }
                            }
                        }
                    } catch {
                        # Ignore errors on OEM cell processing
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
                    
                } catch {
                    # Log error to file but don't display to console
                    "Error processing row $i : $($_.Exception.Message)" | Out-File -FilePath $logPath -Append
                }
            }
            
            # Close the progress form
            $progressForm.Close()
            
            $statusLabel.Text = "Finalizing data..."
            $progressBar.Value = $progressBar.Maximum
            
            # Show final message with results count
            Write-Host "Processed $($parsedData.Count) entries from $($rows.length - 1) rows" -ForegroundColor Green
            
            
            # Export parsed data to CSV
            $csvFilePath = Join-Path $config.PartsRoomDirectory "$selectedSite.csv"
            $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation
            
            if (Test-Path $csvFilePath) {
                Write-Host "CSV file created successfully at $csvFilePath"
                [System.Windows.Forms.MessageBox]::Show("CSV file has been created at: $csvFilePath", "CSV Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                throw "Failed to create CSV file."
            }
            
        } catch {
            Write-Host "Error: $($_.Exception.Message)"
            Write-Host "Stack Trace: $($_.ScriptStackTrace)"
            $errorMessage = "Error: $($_.Exception.Message)`r`nStack Trace: $($_.ScriptStackTrace)"
            $errorMessage | Out-File -FilePath $logPath -Append
            [System.Windows.Forms.MessageBox]::Show("An error occurred. Please check the error log at $logPath for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            # Clean up COM objects
            if ($null -ne $htmlDoc) {
                try {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
                } catch {
                    Write-Host "Warning: Failed to release COM object: $($_.Exception.Message)"
                }
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } else {
        Write-Host "No site was selected. Exiting Set-PartsRoom function..."
    }
}

################################################################################
#                       Integration Functions                                  #
################################################################################


# Function to run Parts-Books-Creator.ps1
function Run-PartsBookCreator {
    param($config)
    $partsBookCreatorPath = Join-Path $config.ScriptsDirectory "Parts-Books-Creator.ps1"
    
    if (Test-Path $partsBookCreatorPath) {
        Write-Host "Running Parts-Books-Creator.ps1 from $partsBookCreatorPath"
        
        # Check for the required CSV file
        $requiredCsvPath = Join-Path $config.DropdownCsvsDirectory "Parsed-Parts-Volumes.csv"
        Write-Host "Checking for required CSV file at: $requiredCsvPath"
        
        if (-not (Test-Path $requiredCsvPath)) {
            Write-Host "ERROR: Required CSV file not found: $requiredCsvPath" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "The required file 'Parsed-Parts-Volumes.csv' was not found in the Dropdown CSVs directory.`n`nThis file is needed for the Parts Books Creator to function. Please ensure this file exists at:`n$requiredCsvPath",
                "Missing Required File",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
                
            return
        }

        & $partsBookCreatorPath
    } else {
        # Show message box when script is NOT found
        [System.Windows.Forms.MessageBox]::Show(
            "The Parts Books Creator script was not found at:`n$partsBookCreatorPath`n`nPlease check that the script exists in the Scripts directory.",
            "Script Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

function Select-InitialBooks {
    param($config)

    $parsedCsvPath = Join-Path $config.DropdownCsvsDirectory "Parsed-Parts-Volumes.csv"
    if (-not (Test-Path $parsedCsvPath) -or (Get-Item $parsedCsvPath).Length -eq 0) {
        Write-Log "Parsed-Parts-Volumes.csv missing or empty. Skipping book selection."
        return @{}
    }

    $csvData = Import-Csv $parsedCsvPath
    if ($csvData.Count -eq 0) { return @{} }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Parts Books to Install"
    $form.Size = New-Object System.Drawing.Size(720, 520)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(690, 40)
    $label.Text = "Select the parts books to install now. You can add or remove books later from the app."
    $form.Controls.Add($label)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(10, 60)
    $checkedList.Size = New-Object System.Drawing.Size(690, 380)
    $checkedList.CheckOnClick = $true

    foreach ($row in $csvData) {
        $display = "$($row.'Full Name')   MS$($row.'MS Book No') Vol $($row.Volume)"
        $checkedList.Items.Add($display) | Out-Null
    }
    $form.Controls.Add($checkedList)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Install Selected"
    $okButton.Location = New-Object System.Drawing.Point(520, 450)
    $okButton.Size = New-Object System.Drawing.Size(180, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $skipButton = New-Object System.Windows.Forms.Button
    $skipButton.Text = "Skip configure later"
    $skipButton.Location = New-Object System.Drawing.Point(320, 450)
    $skipButton.Size = New-Object System.Drawing.Size(180, 30)
    $skipButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($skipButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $skipButton

    $books = @{}
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($idx in $checkedList.CheckedIndices) {
            $row = $csvData[$idx]
            $bookName = ($row.'Full Name' -replace '[^\w\s-]', '' -replace '\s+', ' ').Trim()
            $bookDir  = Join-Path $config.PartsBooksDirectory $bookName
            $books[$bookName] = @{
                VolumesToUrlCsvPath = Join-Path $bookDir "Volumes-to-URL.csv"
                SectionNamesCsvPath = Join-Path $bookDir "SectionNames.txt"
            }
        }
        Write-Log "Seeded $($books.Count) book(s): $($books.Keys -join ', ')"
    } else {
        Write-Log "User skipped initial book selection."
    }
    return $books
}

function Select-InitialSameDaySites {
    param($config)

    $sitesCsvPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
    if (-not (Test-Path $sitesCsvPath) -or (Get-Item $sitesCsvPath).Length -eq 0) {
        Write-Log "Sites.csv missing or empty. Skipping Same Day selection."
        return @()
    }

    $sites = Import-Csv $sitesCsvPath
    if ($sites.Count -eq 0) { return @() }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Same Day Parts Room Sites"
    $form.Size = New-Object System.Drawing.Size(720, 520)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(690, 40)
    $label.Text = "Select sites to mirror into your Same Day Parts Room. You can add or remove sites later from the app."
    $form.Controls.Add($label)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(10, 60)
    $checkedList.Size = New-Object System.Drawing.Size(690, 380)
    $checkedList.CheckOnClick = $true

    foreach ($site in $sites) {
        $checkedList.Items.Add("$($site.'Site ID')   $($site.'Full Name')") | Out-Null
    }
    $form.Controls.Add($checkedList)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Add Selected Sites"
    $okButton.Location = New-Object System.Drawing.Point(520, 450)
    $okButton.Size = New-Object System.Drawing.Size(180, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $skipButton = New-Object System.Windows.Forms.Button
    $skipButton.Text = "Skip configure later"
    $skipButton.Location = New-Object System.Drawing.Point(320, 450)
    $skipButton.Size = New-Object System.Drawing.Size(180, 30)
    $skipButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($skipButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $skipButton

    $selected = @()
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($idx in $checkedList.CheckedIndices) {
            $site = $sites[$idx]
            $selected += @{
                SiteID   = $site.'Site ID'
                FullName = $site.'Full Name'
                Email    = ""
            }
        }
        Write-Log "Seeded $($selected.Count) Same Day site(s)."
    } else {
        Write-Log "User skipped initial Same Day selection."
    }
    return $selected
}

################################################################################
#                           Main Execution                                     #
################################################################################

# Main setup logic
function Run-Setup {
    Write-Log "Starting Run-Setup"
    try {
        $config = Set-InitialConfiguration
        if ($null -eq $config) {
            Write-Log "Setup cancelled by user."
            return
        }

        Copy-SetupFiles -config $config

        # --- Seed Same Day sites ---
        # Book selection happens inside Parts-Books-Creator.ps1, not here.
        $config.SameDayPartsRooms = Select-InitialSameDaySites -config $config

        # Persist seeded values
        $configPath = Join-Path $config.ScriptsDirectory "Config.json"
        $config | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding UTF8
        Write-Log "Seeded SameDayPartsRooms ($($config.SameDayPartsRooms.Count))."

        # Downloads + book downloader
        Set-PartsRoom        -config $config
        Run-PartsBookCreator -config $config   # <-- prompts once, for books

        Write-Log "Setup completed successfully!"
    }
    catch {
        Write-Log "FATAL: $($_.Exception.Message)"
        Write-Log "Stack: $($_.ScriptStackTrace)"
    }
    Read-Host -Prompt "Setup completed. Press Enter to exit"
}

# Run the setup
Run-Setup