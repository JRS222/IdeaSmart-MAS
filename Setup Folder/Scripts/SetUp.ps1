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


# Function to parse and export HTML data into CSV format and create Excel
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
        $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
        $htmlFilePath = Join-Path $config.PartsRoomDirectory "$selectedSite.html"
        Write-Host "Saving HTML content to $htmlFilePath..."
        Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show("Downloaded HTML for $selectedSite", "Download Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        Write-Host "Processing the downloaded HTML file for $selectedSite..."

        $logPath = Join-Path $config.PartsRoomDirectory "error_log.txt"

        try {
            Write-Host "Reading HTML content from file $htmlFilePath"
            $htmlContent = Get-Content -Path $htmlFilePath -Raw -ErrorAction Stop

            if ([string]::IsNullOrWhiteSpace($htmlContent)) {
                throw "HTML content is empty or null"
            }

            # Create HTML DOM object
            $html = New-Object -ComObject "HTMLFile"
            $html.write([System.Text.Encoding]::Unicode.GetBytes($htmlContent))

            Write-Host "Parsing content with DOM..."
            $parsedData = @()

            # Find all TR elements with class "MAIN"
            $rows = $html.getElementsByTagName("tr") | Where-Object { 
                $_.className -eq "MAIN" 
            }

            Write-Host "Number of rows found: $($rows.Count)"

            if ($rows.Count -eq 0) {
                throw "No data rows found in HTML content"
            }

            foreach ($row in $rows) {
                try {
                    $cells = @($row.getElementsByTagName("td"))
                    Write-Host "Processing row with $($cells.Count) cells"

                    if ($cells.Count -ge 6) {
                        # Extract cell values using DOM properties
                        $partNSN = $cells[0].innerText.Trim()
                        $description = $cells[1].innerText.Trim()
                        $qty = $cells[2].innerText.Trim() -replace '[^\d]'
                        $usage = $cells[3].innerText.Trim() -replace '[^\d]'
                        $oemData = $cells[4].innerText.Trim()
                        $location = $cells[5].innerText.Trim()

                        # Process OEM information
                        $oems = @($oemData -split 'OEM:' | Select-Object -Skip 1)
                        $oem1 = if ($oems.Count -gt 0) { ($oems[0] -split '\s+', 2)[-1].Trim() } else { "" }
                        $oem2 = if ($oems.Count -gt 1) { ($oems[1] -split '\s+', 2)[-1].Trim() } else { "" }
                        $oem3 = if ($oems.Count -gt 2) { ($oems[2] -split '\s+', 2)[-1].Trim() } else { "" }

                        $parsedData += [PSCustomObject]@{
                            "Part (NSN)"    = $partNSN
                            "Description"   = $description
                            "QTY"          = [int]($qty -as [int])
                            "13 Period Usage" = [int]($usage -as [int])
                            "Location"      = $location
                            "OEM 1"         = $oem1
                            "OEM 2"         = $oem2
                            "OEM 3"         = $oem3
                        }

                        Write-Host "Added row: Part(NSN)=$partNSN"
                    }
                    else {
                        Write-Host "Row has only $($cells.Count) cells, skipping"
                    }
                }
                catch {
                    Write-Host "Error processing row: $($_.Exception.Message)"
                }
            }

            Write-Host "Number of parsed data entries: $($parsedData.Count)"

            if ($parsedData.Count -eq 0) {
                throw "No data parsed from HTML content"
            }

            Write-Host "Exporting parsed data to CSV..."
            $csvFilePath = Join-Path $config.PartsRoomDirectory "$selectedSite.csv"
            $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation

            if (Test-Path $csvFilePath) {
                Write-Host "CSV file created successfully at $csvFilePath"
                [System.Windows.Forms.MessageBox]::Show("CSV file has been created at: $csvFilePath", "CSV Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
        catch {
            Write-Host "Error: $($_.Exception.Message)"
            Write-Host "Stack Trace: $($_.ScriptStackTrace)"
            [System.Windows.Forms.MessageBox]::Show("An error occurred. Please check the error log at $logPath for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
        Write-Host "Running Parts-Books-Creator.ps1"
        & $partsBookCreatorPath
    } else {
        Write-Host "Parts-Books-Creator.ps1 not found at $partsBookCreatorPath"
    }
}

function Test-ConfigurationValidity {
    param($config)
    
    Write-Log "Validating configuration structure..."
    
    # Required root level paths
    $requiredPaths = @(
        'RootDirectory',
        'LaborDirectory',
        'CallLogsDirectory',
        'PartsRoomDirectory',
        'DropdownCsvsDirectory',
        'PartsBooksDirectory',
        'ScriptsDirectory'
    )

    # Required CSV files
    $requiredCsvFiles = @(
        @{Name = "Machines"; Path = "Machines.csv"},
        @{Name = "Causes"; Path = "Causes.csv"},
        @{Name = "Actions"; Path = "Actions.csv"},
        @{Name = "Nouns"; Path = "Nouns.csv"},
        @{Name = "Sites"; Path = "Sites.csv"}
    )

    # Check for required paths
    $missingPaths = @()
    foreach ($path in $requiredPaths) {
        if (-not $config.$path) {
            $missingPaths += $path
        } elseif (-not (Test-Path $config.$path)) {
            Write-Log "Creating missing directory: $($config.$path)"
            New-Item -ItemType Directory -Force -Path $config.$path | Out-Null
        }
    }

    if ($missingPaths.Count -gt 0) {
        throw "Configuration missing required paths: $($missingPaths -join ', ')"
    }

    # Verify Books structure
    if (-not $config.Books -or $config.Books.Count -eq 0) {
        Write-Log "Warning: No books configured in Books section"
    } else {
        foreach ($book in $config.Books.PSObject.Properties) {
            if (-not $book.Value.VolumesToUrlCsvPath -or -not $book.Value.SectionNamesCsvPath) {
                throw "Invalid book configuration for $($book.Name): Missing required paths"
            }
        }
    }

    # Verify SameDayPartsRooms structure
    if (-not $config.SameDayPartsRooms) {
        $config | Add-Member -NotePropertyName SameDayPartsRooms -NotePropertyValue @() -Force
        Write-Log "Added empty SameDayPartsRooms array to configuration"
    }

    # Check required CSV files
    foreach ($csv in $requiredCsvFiles) {
        $csvPath = Join-Path $config.DropdownCsvsDirectory $csv.Path
        if (-not (Test-Path $csvPath)) {
            Write-Log "Creating empty CSV file: $csvPath"
            # Create with headers based on file type
            switch ($csv.Name) {
                "Machines" { "Machine Acronym,Machine Number" | Out-File $csvPath -Encoding utf8 }
                "Sites" { "Site ID,Full Name" | Out-File $csvPath -Encoding utf8 }
                default { "Value" | Out-File $csvPath -Encoding utf8 }
            }
        }
    }

    # Create empty log files if they don't exist
    $laborLogsPath = Join-Path $config.LaborDirectory "LaborLogs.csv"
    $callLogsPath = Join-Path $config.CallLogsDirectory "CallLogs.csv"

    if (-not (Test-Path $laborLogsPath)) {
        "Date,Work Order,Description,Machine,Duration,Notes" | Out-File $laborLogsPath -Encoding utf8
        Write-Log "Created LaborLogs.csv with headers"
    }

    if (-not (Test-Path $callLogsPath)) {
        "Date,Machine,Cause,Action,Noun,Time Down,Time Up,Notes" | Out-File $callLogsPath -Encoding utf8
        Write-Log "Created CallLogs.csv with headers"
    }

    Write-Log "Configuration validation completed successfully"
    return $true
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

    # Ensure config is properly structured
    Test-ConfigurationValidity

    Write-Log "Setup completed successfully!"
}

# Run the setup
Run-Setup
