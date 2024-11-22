# Load required assemblies for the Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# function # Write-Log {
    # param([string]$message)
    # $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Write-Host "$timestamp - $message"
    # Optionally, you can also write to a log file:
    # "$timestamp - $message" | Out-File -Append -FilePath "setup_log.log"
# }

# Function to get site data directly from server
function Get-SiteData {
    param (
        [string]$url = "http://emarssu5.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_criteria"
    )
    
    # Write-Log "Fetching site data from $url"
    
    try {
        $response = Invoke-WebRequest -Uri $url
        $html = $response.ParsedHtml
        
        if (-not $html) {
            throw "Failed to parse HTML response"
        }
        
        $options = $html.getElementById("pSiteID")
        if (-not $options) {
            throw "Could not find site selection element"
        }
        
        $sites = @()
        foreach ($option in $options) {
            if ($option.value -and $option.innerText) {
                $sites += [PSCustomObject]@{
                    'Site ID' = $option.value.Trim()
                    'Full Name' = $option.innerText.Trim()
                }
            }
        }
        
        # Write-Log "Successfully retrieved ${($sites.Count)} sites"
        return $sites
        
    } catch {
        # Write-Log "Error fetching site data: $_"
        throw "Failed to fetch site data: $_"
    }
}

# Function to parse HTML content
function Parse-HTMLContent {
    param(
        [string]$htmlContent,
        [string]$siteName
    )

    # Write-Log "Processing HTML content for $siteName..."

    try {
        if ([string]::IsNullOrWhiteSpace($htmlContent)) {
            throw "HTML content is empty or null"
        }

        # Create HTML document object
        $html = New-Object -ComObject "HTMLFile"
        $html.IHTMLDocument2_write($htmlContent)
        
        # Find all tables and then filter for the one with class="MAIN"
        $tables = $html.getElementsByTagName("TABLE")
        $mainTable = $tables | Where-Object { $_.className -eq "MAIN" }
        
        if (-not $mainTable) {
            throw "Could not find main table element"
        }

        $parsedData = @()
        
        # Get all rows
        $rows = $mainTable.getElementsByTagName("TR")
        $dataRows = $rows | Where-Object { $_.className -eq "MAIN" }
        # Write-Log "Found $($dataRows.Count) data rows"

        foreach ($row in $dataRows) {
            try {
                $cells = $row.getElementsByTagName("TD")
                
                if ($cells.length -ge 6) {
                    # Extract OEM data
                    $oemCell = $cells[4]
                    $oemDivs = $oemCell.getElementsByTagName("DIV")
                    $oems = @("", "", "")

                    foreach ($div in $oemDivs) {
                        if ($div.innerText -match 'OEM:(\d+)\s+(.+)') {
                            $oemNumber = [int]$Matches[1] - 1
                            $oemValue = $Matches[2].Trim()
                            if ($oemNumber -ge 0 -and $oemNumber -lt 3) {
                                $oems[$oemNumber] = $oemValue
                            }
                        }
                    }

                    $parsedData += [PSCustomObject]@{
                        "Part (NSN)" = $cells[0].innerText.Trim()
                        "Description" = $cells[1].innerText.Trim()
                        "QTY" = [int]($cells[2].innerText -replace '[^\d]', '')
                        "13 Period Usage" = [int]($cells[3].innerText -replace '[^\d]', '')
                        "OEM 1" = $oems[0]
                        "OEM 2" = $oems[1]
                        "OEM 3" = $oems[2]
                        "Location" = $cells[5].innerText.Trim()
                    }

                    # Write-Log "Added row: Part(NSN)=$($cells[0].innerText.Trim()), QTY=$($cells[2].innerText), Location=$($cells[5].innerText)"
                }
            }
            catch {
                # Write-Log "Error processing row: $_"
            }
        }

        # Write-Log "Number of parsed data entries: $($parsedData.Count)"

        if ($parsedData.Count -eq 0) {
            throw "No data parsed from HTML content"
        }

        return $parsedData

    }
    catch {
        # Write-Log "Error: $($_.Exception.Message)"
        # Write-Log "Stack Trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("An error occurred while parsing the HTML for $siteName. Please check the log for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
    finally {
        if ($html) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($html) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
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
    # Write-Log "Starting Set-InitialConfiguration"
    
    # Get the current script's directory
    $currentDir = $PSScriptRoot
    if (-not $currentDir) {
        $currentDir = (Get-Location).Path
        # Write-Log "Warning: Unable to determine script directory. Using current directory: $currentDir"
    }
    $setupFolder = Split-Path -Parent $currentDir

    # Write-Log "Current Directory: $currentDir"
    # Write-Log "Setup Folder: $setupFolder"
    
    # Select parent directory
    $parentDir = Show-FolderBrowserDialog -Description "Select parent directory for PartsBookManagerRootDirectory"
    if (-not $parentDir) { 
        # Write-Log "Setup cancelled by user."
        exit 
    }

    # Create root directory
    $rootDir = Join-Path $parentDir "PartsBookManagerRootDirectory"
    New-Item -ItemType Directory -Force -Path $rootDir | Out-Null
    # Write-Log "Created Root Directory at $rootDir"

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
    # Write-Log "Creating subdirectories..."
    $subDirs = @("PartsBooksDirectory", "ScriptsDirectory", "DropdownCsvsDirectory", "PartsRoomDirectory", "LaborDirectory", "CallLogsDirectory")
    foreach ($dir in $subDirs) {
        New-Item -ItemType Directory -Force -Path $config[$dir] | Out-Null
        # Write-Log "Created directory: $($config[$dir])"
    }

    # Create Same Day Parts Room directory
    $sameDayPartsRoomDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    New-Item -ItemType Directory -Force -Path $sameDayPartsRoomDir | Out-Null
    # Write-Log "Created Same Day Parts Room directory: $sameDayPartsRoomDir"

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
            # Write-Log "Created empty file: $($file.Path)"
        }
        $config.PrerequisiteFiles[$file.Name] = $file.Path
    }

    # Save configuration in the Scripts directory
    $configFilePath = Join-Path $config.ScriptsDirectory "Config.json"
    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configFilePath
    # Write-Log "Configuration saved at $configFilePath"

    return $config
}

# Function to copy setup files to the necessary directories
function Copy-SetupFiles {
    param($config)
    # Write-Log "Starting Copy-SetupFiles"
    
    $setupDir = $config.SetupFolder
    # Write-Log "Setup Directory: $setupDir"

    # Copy Scripts
    # Write-Log "Copying Scripts..."
    $scriptsSourceDir = Join-Path $setupDir "Scripts"
    if (Test-Path $scriptsSourceDir) {
        Copy-Item -Path "$scriptsSourceDir\*" -Destination $config.ScriptsDirectory -Recurse -Force
        # Write-Log "Copied Scripts"
    } else {
        # Write-Log "Warning: Scripts directory not found in the SetupFolder: $scriptsSourceDir"
    }

    # Copy Dropdown CSVs - modified to exclude sites.csv
    $csvFiles = @("Parsed-Parts-Volumes.csv", "Causes.csv", "Actions.csv", "Nouns.csv")
    foreach ($csvFile in $csvFiles) {
        $sourcePath = Join-Path $setupDir $csvFile
        $destPath = Join-Path $config.DropdownCsvsDirectory $csvFile
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            # Write-Log "Copied $csvFile to $destPath"
            $config.PrerequisiteFiles[$csvFile -replace "\.csv", ""] = $destPath
        } else {
            # Write-Log "Warning: $csvFile not found in setup directory: $sourcePath"
            New-Item -ItemType File -Path $destPath -Force | Out-Null
            # Write-Log "Created empty file: $destPath"
        }
    }

    $configFilePath = Join-Path $config.ScriptsDirectory "Config.json"
    $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configFilePath
    # Write-Log "Configuration saved at $configFilePath"

    [System.Windows.Forms.MessageBox]::Show(
        "Setup files have been copied successfully.",
        "Files Copied",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Modified Set-PartsRoom function
function Set-PartsRoom {
    param($config)
    
    try {
        # Fetch sites directly from the server
        $sites = Get-SiteData
        
        if (-not $sites -or $sites.Count -eq 0) {
            throw "No sites available"
        }
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select Facility"
        $form.Size = New-Object System.Drawing.Size(400, 200)
        $form.StartPosition = "CenterScreen"
        
        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Location = New-Object System.Drawing.Point(10, 20)
        $dropdown.Size = New-Object System.Drawing.Size(360, 20)
        $dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        
        foreach ($site in $sites) {
            if (![string]::IsNullOrWhiteSpace($site.'Full Name')) {
                $dropdown.Items.Add($site.'Full Name')
            }
        }
        
        # Write-Log "Populated dropdown with ${($dropdown.Items.Count)} sites"
        $form.Controls.Add($dropdown)
        
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(150, 60)
        $button.Size = New-Object System.Drawing.Size(75, 23)
        $button.Text = "Select"
        $button.Add_Click({
            $form.Tag = $dropdown.SelectedItem
            $form.Close()
        })
        $form.Controls.Add($button)
        
        $form.ShowDialog()
        $selectedSite = $form.Tag
        
        if ($selectedSite) {
            # Write-Log "Selected Site: $selectedSite"
            $selectedRow = $sites | Where-Object { $_.'Full Name' -eq $selectedSite }
            $url = "http://emarssu5.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$($selectedRow.'Site ID')&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

            # Download and process HTML content directly
            # Write-Log "Downloading and processing HTML content for $selectedSite..."
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing
            $parsedData = Parse-HTMLContent -htmlContent $response.Content -siteName $selectedSite

            if ($parsedData.Count -gt 0) {
                # Save CSV file
                $csvFilePath = Join-Path $config.PartsRoomDirectory "$selectedSite.csv"
                $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                # Write-Log "Saved CSV file to $csvFilePath"

                # Create Excel file
                Create-ExcelFromCsv -siteName $selectedSite -csvDirectory $config.PartsRoomDirectory -excelDirectory $config.PartsRoomDirectory
                
                [System.Windows.Forms.MessageBox]::Show("Successfully processed data for $selectedSite", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                throw "No data was parsed from the HTML content"
            }
        }
        else {
            # Write-Log "No site was selected"
        }
    }
    catch {
        # Write-Log "Error in Set-PartsRoom: $_"
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

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
    $worksheet = $null
    $progressForm = $null
    
    try {
        # Create progress bar form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Creating Excel File"
        $progressForm.Width = 400
        $progressForm.Height = 150
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.ControlBox = $false

        # Add progress bar
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Width = 360
        $progressBar.Height = 20
        $progressBar.Location = New-Object System.Drawing.Point(10, 40)
        $progressBar.Style = "Continuous"
        $progressBar.Minimum = 0
        $progressBar.Maximum = 100
        
        # Add status label
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Width = 360
        $statusLabel.Height = 20
        $statusLabel.Location = New-Object System.Drawing.Point(10, 10)
        $statusLabel.Text = "Initializing..."

        # Add percentage label
        $percentLabel = New-Object System.Windows.Forms.Label
        $percentLabel.Width = 360
        $percentLabel.Height = 20
        $percentLabel.Location = New-Object System.Drawing.Point(10, 70)
        $percentLabel.Text = "0%"

        $progressForm.Controls.AddRange(@($progressBar, $statusLabel, $percentLabel))
        
        # Show progress form in a non-blocking way
        $progressForm.Show()
        $progressForm.Refresh()

        # Update progress helper function
        function Update-Progress {
            param(
                [int]$value,
                [string]$status
            )
            if ($progressForm -and $progressForm.Visible) {
                $progressBar.Value = [Math]::Max(0, [Math]::Min(100, $value))
                $statusLabel.Text = $status
                $percentLabel.Text = "$value%"
                $progressForm.Refresh()
            }
        }

        Write-Host "Starting to create Excel file from CSV..."
        Update-Progress -value 5 -status "Checking files..."
        
        $csvFilePath = Join-Path $csvDirectory "$siteName.csv"
        $excelFilePath = Join-Path $excelDirectory "$siteName.xlsx"

        # Ensure the CSV file exists
        if (-not (Test-Path $csvFilePath)) {
            throw "CSV file not found at $csvFilePath"
        }

        Update-Progress -value 10 -status "Initializing Excel..."
        # Initialize Excel with error handling
        $excel = New-Object -ComObject Excel.Application
        if (-not $excel) {
            throw "Failed to create Excel application object"
        }
        
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        Update-Progress -value 15 -status "Reading CSV data..."
        # Import CSV data first to verify it's valid
        $csvData = Import-Csv -Path $csvFilePath
        if ($csvData.Count -eq 0) {
            throw "CSV file is empty"
        }

        Update-Progress -value 20 -status "Creating workbook..."
        # Create new workbook with error handling
        $workbook = $excel.Workbooks.Add()
        if (-not $workbook) {
            throw "Failed to create new workbook"
        }

        $worksheet = $workbook.Worksheets.Item(1)
        if (-not $worksheet) {
            throw "Failed to access worksheet"
        }

        $worksheet.Name = "Parts Data"

        Update-Progress -value 25 -status "Writing headers..."
        # Get headers
        $headers = $csvData[0].PSObject.Properties.Name
        
        # Write headers
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
        }

        Update-Progress -value 30 -status "Writing data..."
        
        # Write data row by row
        $totalRows = $csvData.Count
        for ($row = 0; $row -lt $totalRows; $row++) {
            for ($col = 0; $col -lt $headers.Count; $col++) {
                $worksheet.Cells.Item($row + 2, $col + 1) = $csvData[$row].$($headers[$col])
            }
            
            # Update progress every 5 rows to avoid excessive updates
            if ($row % 5 -eq 0) {
                $progressValue = [int](30 + (($row / $totalRows) * 30))
                Update-Progress -value $progressValue -status "Writing row $($row + 1) of $totalRows..."
            }
        }

        Update-Progress -value 65 -status "Creating table format..."
        # Create and format table
        $usedRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($csvData.Count + 1, $headers.Count))
        $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $usedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
        $listObject.Name = $tableName
        $listObject.TableStyle = "TableStyleMedium2"

        Update-Progress -value 75 -status "Applying cell formatting..."
        # Format cells
        $usedRange.Cells.VerticalAlignment = -4108  # xlCenter
        $usedRange.Cells.HorizontalAlignment = -4108  # xlCenter
        $usedRange.Cells.WrapText = $false
        $usedRange.Cells.Font.Name = "Courier New"
        $usedRange.Cells.Font.Size = 12

        Update-Progress -value 85 -status "Formatting Description column..."
        # Special formatting for Description column
        $descriptionColIndex = [Array]::IndexOf($headers, "Description") + 1
        if ($descriptionColIndex -gt 0) {
            $descriptionRange = $worksheet.Range($worksheet.Cells(2, $descriptionColIndex), $worksheet.Cells($csvData.Count + 1, $descriptionColIndex))
            $descriptionRange.HorizontalAlignment = -4131  # xlLeft
        }

        Update-Progress -value 90 -status "Auto-fitting columns..."
        # AutoFit columns
        $worksheet.UsedRange.Columns.AutoFit() | Out-Null

        Update-Progress -value 95 -status "Saving Excel file..."
        # Save with error handling
        Write-Host "Saving Excel file to: $excelFilePath"
        $workbook.SaveAs($excelFilePath)
        
        if (-not (Test-Path $excelFilePath)) {
            throw "Excel file was not created successfully"
        }

        Update-Progress -value 100 -status "Complete!"
        Start-Sleep -Seconds 1  # Give users a chance to see 100%
        
        Write-Host "Excel file created successfully"
        return $true
    }
    catch {
        Write-Host "Error in Create-ExcelFromCsv: $($_.Exception.Message)"
        if ($progressForm -and $progressForm.Visible) {
            $statusLabel.Text = "Error: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            Start-Sleep -Seconds 3
        }
        return $false
    }
    finally {
        if ($workbook) {
            $workbook.Close($true)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        if ($worksheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        }
        if ($progressForm) {
            $progressForm.Close()
            $progressForm.Dispose()
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to run Parts-Book-Creator.ps1
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
    
    # Write-Log "Validating configuration structure..."
    
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

    # Check for required paths
    $missingPaths = @()
    foreach ($path in $requiredPaths) {
        if (-not $config.$path) {
            $missingPaths += $path
        } elseif (-not (Test-Path $config.$path)) {
            # Write-Log "Creating missing directory: $($config.$path)"
            New-Item -ItemType Directory -Force -Path $config.$path | Out-Null
        }
    }

    if ($missingPaths.Count -gt 0) {
        throw "Configuration missing required paths: $($missingPaths -join ', ')"
    }

    # Verify Books structure
    if (-not $config.Books -or $config.Books.Count -eq 0) {
        # Write-Log "Warning: No books configured in Books section"
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
        # Write-Log "Added empty SameDayPartsRooms array to configuration"
    }

    # Create empty log files if they don't exist
    $laborLogsPath = Join-Path $config.LaborDirectory "LaborLogs.csv"
    $callLogsPath = Join-Path $config.CallLogsDirectory "CallLogs.csv"

    if (-not (Test-Path $laborLogsPath)) {
        "Date,Work Order,Description,Machine,Duration,Notes" | Out-File $laborLogsPath -Encoding utf8
        # Write-Log "Created LaborLogs.csv with headers"
    }

    if (-not (Test-Path $callLogsPath)) {
        "Date,Machine,Cause,Action,Noun,Time Down,Time Up,Notes" | Out-File $callLogsPath -Encoding utf8
        # Write-Log "Created CallLogs.csv with headers"
    }

    # Write-Log "Configuration validation completed successfully"
    return $true
}

# Main setup logic
function Run-Setup {
    # Write-Log "Starting Run-Setup"
    
    try {
        # Set up the initial configuration
        $config = Set-InitialConfiguration
        
        # Copy setup files
        Copy-SetupFiles -config $config
        
        # Ask for site to download and process
        Set-PartsRoom -config $config
        
        # Ensure config is properly structured
        Test-ConfigurationValidity -config $config
        
        # Run the Parts Book Creator script
        Run-PartsBookCreator -config $config
        
        # Write-Log "Setup completed successfully!"
        
    } catch {
        # Write-Log "Error during setup: $_"
        [System.Windows.Forms.MessageBox]::Show("Setup failed: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Run the setup
Run-Setup