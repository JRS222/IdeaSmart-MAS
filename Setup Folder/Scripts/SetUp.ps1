# Load required assemblies for the Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# Function to set up the initial configuration
function Set-InitialConfiguration {
    Write-Log "Starting Set-InitialConfiguration"
    
    # Get the current script's directory
    $currentDir = $PSScriptRoot
    if (-not $currentDir) {
        $currentDir = (Get-Location).Path
        Write-Log "Warning: Unable to determine script directory. Using current directory: $currentDir"
    }
    $setupFolder = Split-Path -Parent $currentDir

    Write-Log "Current Directory: $currentDir"
    Write-Log "Setup Folder: $setupFolder"
    
    # Select parent directory
    $parentDir = Show-FolderBrowserDialog -Description "Select parent directory for PartsBookManagerRootDirectory"
    if (-not $parentDir) { 
        Write-Log "Setup cancelled by user."
        exit 
    }

    # Create root directory
    $rootDir = Join-Path $parentDir "PartsBookManagerRootDirectory"
    New-Item -ItemType Directory -Force -Path $rootDir | Out-Null
    Write-Log "Created Root Directory at $rootDir"

    # Define subdirectories in the config
    $config = @{
        RootDirectory        = $rootDir
        PartsBooksDirectory  = Join-Path $rootDir "Parts Books"
        ScriptsDirectory     = Join-Path $rootDir "Scripts"
        DropdownCsvsDirectory = Join-Path $rootDir "Dropdown CSVs"
        PartsRoomDirectory   = Join-Path $rootDir "Parts Room"
        LaborDirectory       = Join-Path $rootDir "Labor"
        CallLogsDirectory    = Join-Path $rootDir "Call Logs"
        Books                = @{}
        PrerequisiteFiles    = @{}
        SupervisorEmail      = "default@example.com"
        SetupFolder          = $setupFolder
    }

    # Create subdirectories
    Write-Log "Creating subdirectories..."
    $subDirs = @("PartsBooksDirectory", "ScriptsDirectory", "DropdownCsvsDirectory", "PartsRoomDirectory", "LaborDirectory", "CallLogsDirectory")
    foreach ($dir in $subDirs) {
        New-Item -ItemType Directory -Force -Path $config[$dir] | Out-Null
        Write-Log "Created directory: $($config[$dir])"
    }

    # Create Same Day Parts Room directory
    $sameDayPartsRoomDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    New-Item -ItemType Directory -Force -Path $sameDayPartsRoomDir | Out-Null
    Write-Log "Created Same Day Parts Room directory: $sameDayPartsRoomDir"

    # Create empty prerequisite files
    $prerequisiteFiles = @(
        @{Name="CallLogs"; Path=Join-Path $config.CallLogsDirectory "CallLogs.csv"},
        @{Name="LaborLogs"; Path=Join-Path $config.LaborDirectory "LaborLogs.csv"},
        @{Name="Machines"; Path=Join-Path $config.DropdownCsvsDirectory "Machines.csv"},
        @{Name="Causes"; Path=Join-Path $config.DropdownCsvsDirectory "Causes.csv"},
        @{Name="Actions"; Path=Join-Path $config.DropdownCsvsDirectory "Actions.csv"},
        @{Name="Nouns"; Path=Join-Path $config.DropdownCsvsDirectory "Nouns.csv"}
    )

    foreach ($file in $prerequisiteFiles) {
        if (-not (Test-Path $file.Path)) {
            New-Item -ItemType File -Path $file.Path -Force | Out-Null
            Write-Log "Created empty file: $($file.Path)"
        }
        $config.PrerequisiteFiles[$file.Name] = $file.Path
    }

    # Save configuration in the Scripts directory
    $configFilePath = Join-Path $config.ScriptsDirectory "Config.json"
    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configFilePath
    Write-Log "Configuration saved at $configFilePath"

    return $config
}

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
            $config.PrerequisiteFiles[$csvFile -replace "\.csv", ""] = $destPath
        } else {
            Write-Log "Warning: $csvFile not found in setup directory: $sourcePath"
            # Create an empty file if it doesn't exist
            New-Item -ItemType File -Path $destPath -Force | Out-Null
            Write-Log "Created empty file: $destPath"
        }
    }

    # Save configuration in the Scripts directory
    $configFilePath = Join-Path $config.ScriptsDirectory "Config.json"
    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configFilePath
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
    $siteIdColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
    $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
    Write-Host "Site ID Column: $siteIdColumn, Full Name Column: $fullNameColumn"

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
        $url = "http://emarssu5.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$($selectedRow.$siteIdColumn)&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

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
            
            # Initialize array to hold parsed data
            $parsedData = @()
            
            # Process each row (skip first row which is the header)
            for ($i = 1; $i -lt $rows.length; $i++) {
                try {
                    $row = $rows.item($i)
                    if ($null -eq $row) {
                        Write-Host "Warning: Row $i is null, skipping"
                        continue
                    }
                    
                    # Skip rows that don't have the MAIN class or enough cells
                    $rowClass = try { $row.className } catch { "" }
                    if ($rowClass -ne "MAIN" -and $rowClass -ne "HILITE") {
                        Write-Host "Skipping row $i - not a main data row (class: $rowClass)"
                        continue
                    }
                    
                    # Get all cells in the row
                    $cells = $row.getElementsByTagName("td")
                    
                    if ($null -eq $cells -or $cells.length -lt 6) {
                        Write-Host "Skipping row $i - insufficient cells (found: $(if ($null -eq $cells) { "null" } else { $cells.length }))"
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
                                        Write-Host "Warning: Error processing OEM div $j in row $i : $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Host "Warning: Error processing OEM cell in row $i : $($_.Exception.Message)"
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
                    
                    Write-Host "Added row ${i}: Part(NSN)=$partNSN, Description=$description, QTY=$qty, Location=$location"
                } catch {
                    Write-Host "Error processing row $i : $($_.Exception.Message)"
                    # Continue with next row instead of stopping
                }
            }
            
            Write-Host "Number of parsed data entries: $($parsedData.Count)"
            
            if ($parsedData.Count -eq 0) {
                throw "No data parsed from HTML content"
            }
            
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
        
        # Debug: Check CSV content
        try {
            $csvData = Import-Csv -Path $requiredCsvPath
            Write-Host "CSV file found and contains $($csvData.Count) rows"
            
            # Check for required columns
            $requiredColumns = @('Full Name', 'MS Book No', 'Volume')
            $missingColumns = $requiredColumns | Where-Object { 
                -not ($csvData[0].PSObject.Properties.Name -contains $_) 
            }
            
            if ($missingColumns.Count -gt 0) {
                Write-Host "ERROR: CSV file is missing required columns: $($missingColumns -join ', ')" -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show(
                    "The CSV file is missing required columns: $($missingColumns -join ', ')`n`nPlease ensure the CSV file has the correct format.",
                    "CSV Format Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        catch {
            Write-Host "ERROR: Could not read CSV file: $($_.Exception.Message)" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "Could not read the CSV file: $($_.Exception.Message)`n`nPlease ensure the CSV file has the correct format.",
                "CSV Read Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Save the current directory so we can return to it later
        $previousLocation = Get-Location
        
        try {
            # Change to the Scripts directory first
            Set-Location -Path $config.ScriptsDirectory
            Write-Host "Changed working directory to: $($config.ScriptsDirectory)"
            
            # Create a temporary debug script that adds pause to the end
            $debugScriptPath = Join-Path $config.ScriptsDirectory "Debug-PartsBookCreator.ps1"
            
            # Get the original script content
            $originalScript = Get-Content -Path $partsBookCreatorPath -Raw
            
            # Add debug code at the end
            $debugScript = @"
$originalScript

# Added debug code to keep window open
Write-Host ""
Write-Host "=======================================" -ForegroundColor Yellow
Write-Host "Press any key to close this window..." -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow
`$null = `$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
"@
            
            # Write the debug script
            Set-Content -Path $debugScriptPath -Value $debugScript
            Write-Host "Created debug script at: $debugScriptPath"
            
            # Run the debug script
            Write-Host "Launching parts book creator script with debugging..."
            Start-Process powershell -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$debugScriptPath`"" -Wait
            
            # Clean up the debug script
            if (Test-Path $debugScriptPath) {
                Remove-Item -Path $debugScriptPath -Force
                Write-Host "Removed debug script"
            }
            
            Write-Host "Parts-Books-Creator.ps1 execution completed"
        }
        catch {
            Write-Host "Error running Parts-Books-Creator.ps1: $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            # Change back to the previous directory
            Set-Location -Path $previousLocation
            Write-Host "Restored working directory to: $previousLocation"
        }
    } else {
        Write-Host "Parts-Books-Creator.ps1 not found at $partsBookCreatorPath"
        
        # Show message box for user
        [System.Windows.Forms.MessageBox]::Show(
            "The Parts Books Creator script was not found at:`n$partsBookCreatorPath`n`nPlease check that the script exists in the Scripts directory.",
            "Script Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}


# Main setup logic
function Run-Setup {
    Write-Log "Starting Run-Setup"
    
    # Set up the initial configuration
    $config = Set-InitialConfiguration

    # Copy setup files
    Copy-SetupFiles -config $config

    # Ask for site to download and process
    Set-PartsRoom -config $config

    # Run the Parts Book Creator script
    Run-PartsBookCreator -config $config

    Write-Log "Setup completed successfully!"
}

# Run the setup
Run-Setup
