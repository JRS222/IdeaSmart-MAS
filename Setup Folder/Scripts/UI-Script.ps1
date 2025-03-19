# Load required assemblies for the Windows Forms GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
    $configPath = Join-Path $PSScriptRoot "Config.json"
    if (Test-Path $configPath) {
        $config = Get-Content -Path $configPath | ConvertFrom-Json
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

$config = Initialize-Config -configPath $configPath
$laborLogsFilePath = $config.PrerequisiteFiles.LaborLogs
if (-not $laborLogsFilePath) {
    $laborLogsFilePath = Join-Path $config.LaborDirectory "LaborLogs.csv"
    Write-Log "Labor logs file path was not in config, set to: $laborLogsFilePath"
}

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

# Helper Function to fix the Machines CSV
function Fix-MachinesCSV {
    $csvPath = Join-Path $config.DropdownCsvsDirectory "Machines.csv"
    Write-Log "Fixing Machines.csv file: $csvPath"

    if (-not (Test-Path $csvPath)) {
        Write-Log "Machines.csv not found. Creating new file."
        "Machine Acronym,Machine Number" | Out-File -FilePath $csvPath -Encoding utf8
    } else {
        $content = Get-Content $csvPath
        $data = @()
        $hasHeader = $false

        if ($content[0] -eq "Machine Acronym,Machine Number") {
            $hasHeader = $true
            $data = $content | Select-Object -Skip 1
        } else {
            $data = $content
        }

        $uniqueData = @{}
        foreach ($line in $data) {
            $parts = $line -split ','
            if ($parts.Count -eq 2) {
                $machineAcronym = $parts[0].Trim()
                $machineNumber = $parts[1].Trim()
                if (-not [string]::IsNullOrWhiteSpace($machineAcronym) -and -not [string]::IsNullOrWhiteSpace($machineNumber)) {
                    $key = "$machineAcronym|$machineNumber"
                    $uniqueData[$key] = @{Acronym = $machineAcronym; Number = $machineNumber}
                }
            }
        }

        $validData = @("Machine Acronym,Machine Number")
        $validData += $uniqueData.Values | ForEach-Object { "$($_.Acronym),$($_.Number)" }

        $validData | Out-File -FilePath $csvPath -Encoding utf8
        Write-Log "Machines.csv file has been fixed, reformatted, and duplicate combinations removed."
    }
}

# Helper Function for Machine CSV
function Inspect-MachinesCSV {
    $csvPath = Join-Path $config.DropdownCsvsDirectory "Machines.csv"
    Write-Log "Inspecting Machines.csv file: $csvPath"

    if (-not (Test-Path $csvPath)) {
        Write-Log "Error: Machines.csv not found at $csvPath"
        return
    }

    $content = Get-Content $csvPath
    Write-Log "Machines.csv content:"
    $content | ForEach-Object { Write-Log $_ }

    $data = Import-Csv -Path $csvPath
    Write-Log "Parsed CSV data:"
    $data | ForEach-Object { Write-Log ($_ | Out-String) }

    Write-Log "Total rows in Machines.csv: $($data.Count)"
}

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
        $script:notificationIcon.Text = "●$($script:unacknowledgedEntries.Count)"
    } else {
        $script:notificationIcon.Visible = $false
    }
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
                $config | ConvertTo-Json | Set-Content -Path $configPath
                Write-Log "Supervisor email updated to: $email"
                return $email
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid email format. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return Get-SupervisorEmail  # Recursive call to prompt again
            }
        } else {
            return $null
        }
    } else {
        return $config.SupervisorEmail
    }
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

# Ensure Labor directory exists
if (-not (Test-Path $config.LaborDirectory)) {
    New-Item -ItemType Directory -Path $config.LaborDirectory | Out-Null
    Write-Log "Created Labor directory: $config.LaborDirectory"
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
        $config | ConvertTo-Json -Depth 4 | Set-Content -Path $configPath

        # Create subdirectory for Same Day Parts Room
        $sameDayPartsRoomDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
        if (-not (Test-Path $sameDayPartsRoomDir)) {
            New-Item -Path $sameDayPartsRoomDir -ItemType Directory | Out-Null
            Write-Log "Created directory: $sameDayPartsRoomDir"
        }

        # Process each selected site
        foreach ($site in $selectedSites) {
            $siteID = $site.SiteID
            $siteName = $site.FullName

            # Get site URL from Sites.csv
            $siteInfo = $sites | Where-Object { $_.'Site ID' -eq $siteID }
            if ($siteInfo) {
                $siteUrl = "http://emarssu5.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$($siteID)&p_search_type=DESC&p_search_string=&p_boh_radio=-1"

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
                    Write-Log "URL not found for site $siteID"
                }
            } else {
                Write-Log "Site information not found for site $siteID"
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Selected sites have been added and processed.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })

    $form.ShowDialog()
}

# Helper function to parse HTML content into CSV data
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

# Function to get Excel files (Parts Books)
function Get-ExcelFiles {
    $excelFiles = @()
    if ($config.Books) {
        Write-Log "Processing Books from config:"
        foreach ($book in $config.Books.PSObject.Properties) {
            $bookDir = Join-Path $config.PartsBooksDirectory $book.Name
            $excelFilePath = Get-ChildItem -Path $bookDir -Filter "*.xlsx" | Select-Object -First 1 -ExpandProperty FullName
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

# Function to create and show the main form
function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Parts Management System'
    $form.Size = New-Object System.Drawing.Size(900, 800)
    $form.StartPosition = 'CenterScreen'

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = 'Fill'
    $form.Controls.Add($tabControl)

    # Parts Books Tab
    $partsBookTab = New-Object System.Windows.Forms.TabPage
    $partsBookTab.Text = "Parts Books"
    $tabControl.TabPages.Add($partsBookTab)

    $partsBookPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $partsBookPanel.Dock = 'Fill'
    $partsBookPanel.FlowDirection = 'TopDown'
    $partsBookPanel.WrapContents = $false
    $partsBookPanel.AutoScroll = $true
    $partsBookTab.Controls.Add($partsBookPanel)

    # Call Logs Tab
    $callLogsTab = New-Object System.Windows.Forms.TabPage
    $callLogsTab.Text = "Call Logs"
    $tabControl.TabPages.Add($callLogsTab)

    # Set up Call Logs tab
    Setup-CallLogsTab -parentTab $callLogsTab

    # Labor Log Tab
    $laborLogTab = New-Object System.Windows.Forms.TabPage
    $laborLogTab.Text = "Labor Log"
    $tabControl.TabPages.Add($laborLogTab)

    # Set up Labor Log tab
    Setup-LaborLogTab -parentTab $laborLogTab -tabControl $tabControl

    # Process historical logs
    Process-HistoricalLogs

    # Add form closing event to save logs
    $form.Add_FormClosing({
        $script:listViewCallLogs = $callLogsTab.Controls | Where-Object { $_ -is [System.Windows.Forms.ListView] }
        if ($script:listViewCallLogs) {
            Save-CallLogs -listView $script:listViewCallLogs -filePath $callLogsFilePath
        }

        $listViewLaborLog = $laborLogTab.Controls | Where-Object { $_ -is [System.Windows.Forms.ListView] }
        if ($listViewLaborLog) {
            Save-LaborLogs -listView $listViewLaborLog -filePath $laborLogsFilePath
        }
    })

    # Add Open Parts Room button
    $openPartsRoomButton = New-Button "Open Parts Room" {
        $partsRoomFilePath = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.xlsx" | Select-Object -First 1 -ExpandProperty FullName
        if (Test-Path $partsRoomFilePath) {
            Start-Process $partsRoomFilePath
        } else {
            [System.Windows.Forms.MessageBox]::Show("Parts Room file not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    $partsBookPanel.Controls.Add($openPartsRoomButton)

    # Add Create Parts Book button
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
    }
    $partsBookPanel.Controls.Add($createPartsBookButton)

    # Add Parts Books buttons
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
        } -Tag $file.Path
        $partsBookPanel.Controls.Add($button)
    }

    # Actions Tab
    $actionsTab = New-Object System.Windows.Forms.TabPage
    $actionsTab.Text = "Actions"
    $tabControl.TabPages.Add($actionsTab)

    $actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionsPanel.Dock = 'Fill'
    $actionsPanel.FlowDirection = 'TopDown'
    $actionsPanel.WrapContents = $false
    $actionsPanel.AutoScroll = $true
    $actionsTab.Controls.Add($actionsPanel)

    # Define action buttons
    $actionButtons = @(
        @{Text="Update Parts Books"; Action={ Update-PartsBooks }}
        @{Text="Update Parts Room"; Action={ Update-PartsRoom }}
        @{Text="Take a Part Out"; Action={ Take-PartOut }}
        @{Text="Search for a Part"; Action={ $tabControl.SelectedTab = $searchTab }}
        @{Text="Request a Part to be Ordered"; Action={ Request-PartOrder }}
        @{Text="Request a Work Order"; Action={ Request-WorkOrder }}
        @{Text="Make an MTSC Ticket"; Action={ Make-MTSCTicket }}
        @{Text="Search Knowledge Base"; Action={ Search-KnowledgeBase }}
        @{Text="Add Same Day Parts Room"; Action={ Add-SameDayPartsRoom }}
        @{Text="Add 1-Day Parts Room"; Action={ [System.Windows.Forms.MessageBox]::Show("This feature is not yet implemented.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }}
        @{Text="Add 2-Day Parts Room"; Action={ [System.Windows.Forms.MessageBox]::Show("This feature is not yet implemented.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }}
    )

    foreach ($actionButton in $actionButtons) {
        $button = New-Button $actionButton.Text $actionButton.Action
        $actionsPanel.Controls.Add($button)
    }


    # Search Tab
    $searchTab = New-Object System.Windows.Forms.TabPage
    $searchTab.Text = "Search"
    $tabControl.TabPages.Add($searchTab)

    # Call the function to set up the search interface within the searchTab
    Setup-SearchTab -parentTab $searchTab -config $config

    Write-Log "UI setup completed"
    $form.ShowDialog()
}

# Helper function to create buttons
function New-Button {
    param(
        [string]$text,
        [scriptblock]$action,
        [object]$Tag = $null
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Width = 450
    $button.Height = 30
    $button.Margin = New-Object System.Windows.Forms.Padding(5)
    if ($Tag -ne $null) {
        $button.Tag = $Tag
    }
    $button.Add_Click($action)
    return $button
}

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
    $regexPattern = '^' + [regex]::Escape($Pattern).Replace('\*', '.*') + '$'
    return $InputString -match $regexPattern
}

# Function to set up the Search tab with enhanced debugging
function Setup-SearchTab {
    param($parentTab, $config)

    Write-Log "Setting up Search tab..."

    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock = 'Fill'
    $parentTab.Controls.Add($searchPanel)

    # Create controls with script scope
    $script:textBoxNSN = New-Object System.Windows.Forms.TextBox
    $script:textBoxOEM = New-Object System.Windows.Forms.TextBox
    $script:textBoxDescription = New-Object System.Windows.Forms.TextBox
    $script:listViewAvailability = New-Object System.Windows.Forms.ListView
    $script:listViewSameDayAvailability = New-Object System.Windows.Forms.ListView
    $script:listViewCrossRef = New-Object System.Windows.Forms.ListView

    # Set up NSN controls
    $labelNSN = New-Object System.Windows.Forms.Label
    $labelNSN.Text = "NSN:"
    $labelNSN.Location = New-Object System.Drawing.Point(20, 20)
    $labelNSN.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelNSN)

    $script:textBoxNSN.Location = New-Object System.Drawing.Point(130, 20)
    $script:textBoxNSN.Size = New-Object System.Drawing.Size(200, 20)
    $searchPanel.Controls.Add($script:textBoxNSN)

    # Set up OEM controls
    $labelOEM = New-Object System.Windows.Forms.Label
    $labelOEM.Text = "OEM:"
    $labelOEM.Location = New-Object System.Drawing.Point(20, 60)
    $labelOEM.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelOEM)

    $script:textBoxOEM.Location = New-Object System.Drawing.Point(130, 60)
    $script:textBoxOEM.Size = New-Object System.Drawing.Size(200, 20)
    $searchPanel.Controls.Add($script:textBoxOEM)

    # Set up Description controls
    $labelDescription = New-Object System.Windows.Forms.Label
    $labelDescription.Text = "Description:"
    $labelDescription.Location = New-Object System.Drawing.Point(20, 100)
    $labelDescription.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelDescription)

    $script:textBoxDescription.Location = New-Object System.Drawing.Point(130, 100)
    $script:textBoxDescription.Size = New-Object System.Drawing.Size(200, 20)
    $searchPanel.Controls.Add($script:textBoxDescription)

    # Create a tooltip
    $tooltip = New-Object System.Windows.Forms.ToolTip

    # Set up tooltips for search textboxes
    $tooltip.SetToolTip($script:textBoxNSN, "Use * as a wildcard. E.g., 1234*567")
    $tooltip.SetToolTip($script:textBoxOEM, "Use * as a wildcard. E.g., ABC*123")
    $tooltip.SetToolTip($script:textBoxDescription, "Use * as a wildcard in your description search.")

    # Set up Search button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(350, 60)
    $searchButton.Size = New-Object System.Drawing.Size(75, 30)
    $searchPanel.Controls.Add($searchButton)

    # Set up Availability ListView
    $labelAvailability = New-Object System.Windows.Forms.Label
    $labelAvailability.Text = "Availability"
    $labelAvailability.Location = New-Object System.Drawing.Point(20, 130)
    $labelAvailability.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelAvailability)

    $script:listViewAvailability.Location = New-Object System.Drawing.Point(20, 160)
    $script:listViewAvailability.Size = New-Object System.Drawing.Size(850, 150)
    $script:listViewAvailability.View = [System.Windows.Forms.View]::Details
    $script:listViewAvailability.FullRowSelect = $true
    $script:listViewAvailability.CheckBoxes = $true
    $script:listViewAvailability.Columns.Add("Part (NSN)", 100)
    $script:listViewAvailability.Columns.Add("Description", 200)
    $script:listViewAvailability.Columns.Add("QTY", 50)
    $script:listViewAvailability.Columns.Add("13 Period Usage", 100)
    $script:listViewAvailability.Columns.Add("Location", 100)
    $script:listViewAvailability.Columns.Add("OEM 1", 100)
    $script:listViewAvailability.Columns.Add("OEM 2", 100)
    $script:listViewAvailability.Columns.Add("OEM 3", 100)
    $script:listViewAvailability.Columns.Add("Changed Part (NSN)", 100)
    $searchPanel.Controls.Add($script:listViewAvailability)

    # Set up Same Day Availability ListView
    $labelSameDayAvailability = New-Object System.Windows.Forms.Label
    $labelSameDayAvailability.Text = "Same Day Parts Availability"
    $labelSameDayAvailability.Location = New-Object System.Drawing.Point(20, 320)
    $labelSameDayAvailability.Size = New-Object System.Drawing.Size(200, 20)
    $searchPanel.Controls.Add($labelSameDayAvailability)

    $script:listViewSameDayAvailability.Location = New-Object System.Drawing.Point(20, 350)
    $script:listViewSameDayAvailability.Size = New-Object System.Drawing.Size(850, 150)
    $script:listViewSameDayAvailability.View = [System.Windows.Forms.View]::Details
    $script:listViewSameDayAvailability.FullRowSelect = $true
    $script:listViewSameDayAvailability.CheckBoxes = $true
    $script:listViewSameDayAvailability.Columns.Add("Part (NSN)", 100)
    $script:listViewSameDayAvailability.Columns.Add("Description", 200)
    $script:listViewSameDayAvailability.Columns.Add("QTY", 50)
    $script:listViewSameDayAvailability.Columns.Add("13 Period Usage", 100)
    $script:listViewSameDayAvailability.Columns.Add("Location", 100)
    $script:listViewSameDayAvailability.Columns.Add("OEM 1", 100)
    $script:listViewSameDayAvailability.Columns.Add("OEM 2", 100)
    $script:listViewSameDayAvailability.Columns.Add("OEM 3", 100)
    $script:listViewSameDayAvailability.Columns.Add("Site Name", 100)
    $searchPanel.Controls.Add($script:listViewSameDayAvailability)

    # Set up Cross Reference ListView
    $labelCrossRef = New-Object System.Windows.Forms.Label
    $labelCrossRef.Text = "Cross Reference"
    $labelCrossRef.Location = New-Object System.Drawing.Point(20, 510)
    $labelCrossRef.Size = New-Object System.Drawing.Size(100, 20)
    $searchPanel.Controls.Add($labelCrossRef)

    $script:listViewCrossRef.Location = New-Object System.Drawing.Point(20, 540)
    $script:listViewCrossRef.Size = New-Object System.Drawing.Size(850, 150)
    $script:listViewCrossRef.View = [System.Windows.Forms.View]::Details
    $script:listViewCrossRef.FullRowSelect = $true
    $script:listViewCrossRef.CheckBoxes = $true
    $script:listViewCrossRef.Columns.Add("Handbook", 100)
    $script:listViewCrossRef.Columns.Add("Section Name", 150)
    $script:listViewCrossRef.Columns.Add("NO.", 50)
    $script:listViewCrossRef.Columns.Add("PART DESCRIPTION", 200)
    $script:listViewCrossRef.Columns.Add("REF.", 100)
    $script:listViewCrossRef.Columns.Add("STOCK NO.", 100)
    $script:listViewCrossRef.Columns.Add("PART NO.", 100)
    $script:listViewCrossRef.Columns.Add("CAGE", 50)
    $script:listViewCrossRef.Columns.Add("Location", 100)
    $script:listViewCrossRef.Columns.Add("QTY", 50)
    $searchPanel.Controls.Add($script:listViewCrossRef)

    # Set up "Open Figures" button
    $script:openFiguresButton = New-Object System.Windows.Forms.Button
    $script:openFiguresButton.Text = "Open Figures"
    $script:openFiguresButton.Location = New-Object System.Drawing.Point(20, 700)
    $script:openFiguresButton.Size = New-Object System.Drawing.Size(120, 30)
    $searchPanel.Controls.Add($script:openFiguresButton)

    # Set up "Take Part(s) Out" button
    $script:takePartOutButton = New-Object System.Windows.Forms.Button
    $script:takePartOutButton.Text = "Take Part(s) Out"
    $script:takePartOutButton.Location = New-Object System.Drawing.Point(150, 700)
    $script:takePartOutButton.Size = New-Object System.Drawing.Size(120, 30)
    $script:takePartOutButton.Enabled = $false
    $searchPanel.Controls.Add($script:takePartOutButton)

    # Event handlers
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
                $ref = $item.SubItems[4].Text  # REF. is the 5th column (index 4)

                if ($handbook -and $ref) {
                    $bookDir = Join-Path $config.PartsBooksDirectory $handbook
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
        $takePartOutButton.Enabled = ($script:listViewAvailability.CheckedItems.Count -gt 0 -or $script:listViewSameDayAvailability.CheckedItems.Count -gt 0)
    })

    $script:listViewSameDayAvailability.Add_ItemChecked({
        $takePartOutButton.Enabled = ($script:listViewAvailability.CheckedItems.Count -gt 0 -or $script:listViewSameDayAvailability.CheckedItems.Count -gt 0)
    })

    $script:listViewCrossRef.Add_ItemChecked({
        $openFiguresButton.Enabled = ($script:listViewCrossRef.CheckedItems.Count -gt 0)
    })

    $searchButton.Add_Click({
        Write-Log "Performing part search..."

        $nsnSearch = $script:textBoxNSN.Text.Trim()
        $oemSearch = $script:textBoxOEM.Text.Trim()
        $descriptionSearch = $script:textBoxDescription.Text.Trim()

        Write-Log "Search criteria - NSN: $nsnSearch, OEM: $oemSearch, Description: $descriptionSearch"

        #### Availability Search ####
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
            $data = Import-Csv -Path $csvFilePath
            Write-Log "CSV file loaded successfully. Row count: $($data.Count)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to read CSV file. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Failed to read CSV file. Error: $_"
            return
        }

        $filteredData = $data | Where-Object {
            ($nsnSearch -eq '' -or $_.'Part (NSN)' -like "*$nsnSearch*" -or $_.'Changed Part (NSN)' -like "*$nsnSearch*") -and
            ($oemSearch -eq '' -or ($_.'OEM 1' -like "*$oemSearch*" -or $_.'OEM 2' -like "*$oemSearch*" -or $_.'OEM 3' -like "*$oemSearch*")) -and
            ($descriptionSearch -eq '' -or $_.Description -like "*$descriptionSearch*") -and
    ([int]$_.QTY -gt 0)
        }

        Write-Log "Found $($filteredData.Count) matching records in Availability."

        $script:listViewAvailability.Items.Clear()

        if ($filteredData.Count -gt 0) {
            foreach ($row in $filteredData) {
                $item = New-Object System.Windows.Forms.ListViewItem($row.'Part (NSN)')
                $item.SubItems.Add($row.Description)
                $item.SubItems.Add($row.QTY)
                $item.SubItems.Add($row.'13 Period Usage')
                $item.SubItems.Add($row.Location)
                $item.SubItems.Add($row.'OEM 1')
                $item.SubItems.Add($row.'OEM 2')
                $item.SubItems.Add($row.'OEM 3')
                if ($row.PSObject.Properties.Name -contains 'Changed Part (NSN)') {
                    $item.SubItems.Add($row.'Changed Part (NSN)')
                } else {
                    $item.SubItems.Add("")
                }
                $script:listViewAvailability.Items.Add($item)
            }
            Write-Log "Added $($filteredData.Count) items to Availability ListView."
        } else {
            [System.Windows.Forms.MessageBox]::Show("No matching records found in Availability.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Log "No matching records found in Availability."
        }

        #### Same Day Parts Availability Search ####
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
                    Write-Log "Processed $($csvData.Count) rows from $($csvFile.Name)"
                } catch {
                    Write-Log "Failed to read CSV file $($csvFile.FullName). Error: $_"
                }
            }

            $filteredSameDayData = $sameDayData | Where-Object {
                ($nsnSearch -eq '' -or $_.'Part (NSN)' -like "*$nsnSearch*") -and
                ($oemSearch -eq '' -or ($_.'OEM 1' -like "*$oemSearch*" -or $_.'OEM 2' -like "*$oemSearch*" -or $_.'OEM 3' -like "*$oemSearch*")) -and
                ($descriptionSearch -eq '' -or $_.Description -like "*$descriptionSearch*") -and
    ([int]$_.QTY -gt 0)
            }

            Write-Log "Found $($filteredSameDayData.Count) matching records in Same Day Parts Availability."

            $script:listViewSameDayAvailability.Items.Clear()

            if ($filteredSameDayData.Count -gt 0) {
                foreach ($row in $filteredSameDayData) {
                    $item = New-Object System.Windows.Forms.ListViewItem($row.'Part (NSN)')
                    $item.SubItems.Add($row.Description)
                    $item.SubItems.Add($row.QTY)
                    $item.SubItems.Add($row.'13 Period Usage')
                    $item.SubItems.Add($row.Location)
                    $item.SubItems.Add($row.'OEM 1')
                    $item.SubItems.Add($row.'OEM 2')
                    $item.SubItems.Add($row.'OEM 3')
                    $item.SubItems.Add($row.SiteName)
                    $script:listViewSameDayAvailability.Items.Add($item)
                }
                Write-Log "Added $($filteredSameDayData.Count) items to Same Day Availability ListView."
            } else {
                [System.Windows.Forms.MessageBox]::Show("No matching records found in Same Day Parts Availability.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                Write-Log "No matching records found in Same Day Parts Availability."
            }
        } else {
            Write-Log "Same Day Parts Room directory not found at $sameDayPartsDir"
        }

        #### Cross Reference Search ####
        $crossRefResults = @()

        if ($config.Books) {
            foreach ($book in $config.Books.PSObject.Properties) {
                $bookName = $book.Name
                $bookDir = Join-Path $config.PartsBooksDirectory $bookName
                $combinedSectionsDir = Join-Path $bookDir "CombinedSections"
                $sectionNamesFile = Join-Path $bookDir "SectionNames.txt"

                # Create a mapping of section numbers to full section names
                $sectionNameMapping = @{}
                if (Test-Path $sectionNamesFile) {
                    $sectionNames = Get-Content -Path $sectionNamesFile
                    foreach ($line in $sectionNames) {
                        if ($line -match '^Section\s+(\d+)\s*(.*)$') {
                            $sectionNumber = $Matches[1]
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

                    # Extract section number and get full section name
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
                
                    $filteredSectionData = $sectionData | Where-Object {
                        ($nsnSearch -eq '' -or $_.'STOCK NO.' -like "*$nsnSearch*") -and
                        ($oemSearch -eq '' -or $_.'PART NO.' -like "*$oemSearch*") -and
                        ($descriptionSearch -eq '' -or $_.'PART DESCRIPTION' -like "*$descriptionSearch*")
                    }
                
                    foreach ($item in $filteredSectionData) {

                    }
                }
            }

            Write-Log "Found $($crossRefResults.Count) matching records in Cross Reference."

            $script:listViewCrossRef.Items.Clear()

            if ($crossRefResults.Count -gt 0) {
                $columns = @('Handbook', 'Section Name', 'NO.', 'PART DESCRIPTION', 'REF.', 'STOCK NO.', 'PART NO.', 'CAGE', 'Location', 'QTY')

                foreach ($row in $crossRefResults) {
                    $item = New-Object System.Windows.Forms.ListViewItem($row.Handbook)
                    foreach ($column in $columns[1..($columns.Count-1)]) {
                        $value = if ($row.PSObject.Properties.Name -contains $column) { $row.$column } else { "" }
                        if ($column -eq 'REF.' -and $value -is [string]) {
                            $value = $value -replace '\.csv$', ''
                        }
                        $item.SubItems.Add($value)
                    }
                    $script:listViewCrossRef.Items.Add($item)
                }

                $script:listViewCrossRef.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
                Write-Log "Added $($crossRefResults.Count) items to Cross Reference ListView."
            } else {
                [System.Windows.Forms.MessageBox]::Show("No matching records found in Cross Reference.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                Write-Log "No matching records found in Cross Reference."
            }
        } else {
            Write-Log "No books defined in configuration."
        }
    })

    Write-Log "Search tab setup completed."
}

# Function to update Parts Books
function Update-PartsBooks {
    Write-Log "Updating Parts Books..."
    [System.Windows.Forms.MessageBox]::Show("Parts Books update not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to update Parts Room
function Update-PartsRoom {
    Write-Log "Updating Parts Room..."
    [System.Windows.Forms.MessageBox]::Show("Parts Room update not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to take a part out
function Take-PartOut {
    Write-Log "Taking a part out..."
    [System.Windows.Forms.MessageBox]::Show("Take Part Out process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

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

# Function to save logs to a CSV file
function Save-CallLogs {
    param($listView, $filePath)
    
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

# Function to set up the Call Logs tab
function Setup-CallLogsTab {
    param($parentTab)

    Write-Log "Setting up Call Logs tab..."

    $callLogsPanel = New-Object System.Windows.Forms.Panel
    $callLogsPanel.Dock = 'Fill'
    $parentTab.Controls.Add($callLogsPanel)

    # Call Log ListView
    $script:listViewCallLogs = New-Object System.Windows.Forms.ListView
    $script:listViewCallLogs.Location = New-Object System.Drawing.Point(10, 10)
    $script:listViewCallLogs.Size = New-Object System.Drawing.Size(850, 420)
    $script:listViewCallLogs.View = [System.Windows.Forms.View]::Details
    $script:listViewCallLogs.FullRowSelect = $true
    $script:listViewCallLogs.Columns.Clear()
    $script:listViewCallLogs.Columns.Add("Date", 100) | Out-Null
    $script:listViewCallLogs.Columns.Add("Machine ID", 120) | Out-Null
    $script:listViewCallLogs.Columns.Add("Cause", 100) | Out-Null
    $script:listViewCallLogs.Columns.Add("Action", 100) | Out-Null
    $script:listViewCallLogs.Columns.Add("Noun", 100) | Out-Null
    $script:listViewCallLogs.Columns.Add("Time Down", 80) | Out-Null
    $script:listViewCallLogs.Columns.Add("Time Up", 80) | Out-Null
    $script:listViewCallLogs.Columns.Add("Notes", 150) | Out-Null
    $callLogsPanel.Controls.Add($script:listViewCallLogs)

    # Load existing call logs
    Load-Logs -listView $script:listViewCallLogs -filePath $callLogsFilePath

    # Ensure Labor Logs CSV exists and process historical logs
    Process-HistoricalLogs

    # Add New Call Log Button
    $addCallLogButton = New-Object System.Windows.Forms.Button
    $addCallLogButton.Location = New-Object System.Drawing.Point(10, 440)
    $addCallLogButton.Size = New-Object System.Drawing.Size(150, 30)
    $addCallLogButton.Text = "Add New Call Log"
    $addCallLogButton.Add_Click({
        $addCallLogForm = New-Object System.Windows.Forms.Form
        $addCallLogForm.Text = "Add New Call Log"
        $addCallLogForm.Size = New-Object System.Drawing.Size(400, 500)
        $addCallLogForm.StartPosition = 'CenterScreen'


        # Date
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

        # Machine ID
        $labelMachineId = New-Object System.Windows.Forms.Label
        $labelMachineId.Location = New-Object System.Drawing.Point(10, 50)
        $labelMachineId.Size = New-Object System.Drawing.Size(100, 20)
        $labelMachineId.Text = "Machine ID:"
        $addCallLogForm.Controls.Add($labelMachineId)

        $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox
        $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 50)
        $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20)
        $addCallLogForm.Controls.Add($comboBoxMachineId)

        # Load Machine IDs
        Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"

        # Cause
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

        # Action
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

        # Noun
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

        # Time Down
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

        # Time Up
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

        # Notes
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

        # Add button
        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(150, 330)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add"
        $addButton.Add_Click({
            $item = New-Object System.Windows.Forms.ListViewItem($textBoxDate.Text)
            $item.SubItems.Add($comboBoxMachineId.SelectedItem)
            $item.SubItems.Add($comboBoxCause.SelectedItem)
            $item.SubItems.Add($comboBoxAction.SelectedItem)
            $item.SubItems.Add($comboBoxNoun.SelectedItem)
            $timeDown = "$($comboBoxTimeDownHour.SelectedItem):$($comboBoxTimeDownMinute.SelectedItem)"
            $timeUp = "$($comboBoxTimeUpHour.SelectedItem):$($comboBoxTimeUpMinute.SelectedItem)"
            $item.SubItems.Add($timeDown)
            $item.SubItems.Add($timeUp)
            $item.SubItems.Add($textBoxNotes.Text)

            $script:listViewCallLogs.Items.Add($item)

            # Save call logs after adding a new entry
            Save-Logs -listView $script:listViewCallLogs -filePath $callLogsFilePath

            $addCallLogForm.Close()
        })
        $addCallLogForm.Controls.Add($addButton)

        $addCallLogForm.ShowDialog()
    })
    $callLogsPanel.Controls.Add($addCallLogButton)

    # Add Machine Button
    $addMachineButton = New-Object System.Windows.Forms.Button
    $addMachineButton.Location = New-Object System.Drawing.Point(170, 440)
    $addMachineButton.Size = New-Object System.Drawing.Size(150, 30)
    $addMachineButton.Text = "Add Machine"
    $addMachineButton.Add_Click({
        $addMachineForm = New-Object System.Windows.Forms.Form
        $addMachineForm.Text = "Add New Machine"
        $addMachineForm.Size = New-Object System.Drawing.Size(300, 200)
        $addMachineForm.StartPosition = 'CenterScreen'

        $labelAcronym = New-Object System.Windows.Forms.Label
        $labelAcronym.Location = New-Object System.Drawing.Point(10, 20)
        $labelAcronym.Size = New-Object System.Drawing.Size(100, 20)
        $labelAcronym.Text = "Machine Acronym:"
        $addMachineForm.Controls.Add($labelAcronym)

        $textBoxAcronym = New-Object System.Windows.Forms.TextBox
        $textBoxAcronym.Location = New-Object System.Drawing.Point(120, 20)
        $textBoxAcronym.Size = New-Object System.Drawing.Size(150, 20)
        $addMachineForm.Controls.Add($textBoxAcronym)

        $labelEquipmentNumber = New-Object System.Windows.Forms.Label
        $labelEquipmentNumber.Location = New-Object System.Drawing.Point(10, 50)
        $labelEquipmentNumber.Size = New-Object System.Drawing.Size(100, 20)
        $labelEquipmentNumber.Text = "Equipment Number:"
        $addMachineForm.Controls.Add($labelEquipmentNumber)

        $textBoxEquipmentNumber = New-Object System.Windows.Forms.TextBox
        $textBoxEquipmentNumber.Location = New-Object System.Drawing.Point(120, 50)
        $textBoxEquipmentNumber.Size = New-Object System.Drawing.Size(150, 20)
        $addMachineForm.Controls.Add($textBoxEquipmentNumber)

        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(100, 90)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add Machine"
        $addButton.Add_Click({
            $machineAcronym = $textBoxAcronym.Text.Trim()
            $equipmentNumber = $textBoxEquipmentNumber.Text.Trim()
    
            if ([string]::IsNullOrWhiteSpace($machineAcronym) -or [string]::IsNullOrWhiteSpace($equipmentNumber)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter both Machine Acronym and Equipment Number.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
    
            $machinesCsvPath = Join-Path $config.DropdownCsvsDirectory "Machines.csv"
    
            # Create the CSV file if it doesn't exist
            if (-not (Test-Path $machinesCsvPath)) {
                "Machine Acronym,Machine Number" | Out-File -FilePath $machinesCsvPath -Encoding utf8
                Write-Log "Created new Machines.csv file at $machinesCsvPath"
            }
    
            # Read existing data
            $existingData = Import-Csv -Path $machinesCsvPath
    
            # Check if the entry already exists
            if ($existingData | Where-Object { $_.'Machine Acronym' -eq $machineAcronym -and $_.'Machine Number' -eq $equipmentNumber }) {
                [System.Windows.Forms.MessageBox]::Show("This machine and equipment number combination already exists.", "Duplicate Entry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
    
            # Append the new machine to the CSV
            "$machineAcronym,$equipmentNumber" | Out-File -FilePath $machinesCsvPath -Append -Encoding utf8
            Write-Log "Added new machine: $machineAcronym with equipment number: $equipmentNumber to $machinesCsvPath"
    
            # Set a flag to indicate that the Machines list has been updated
            $script:machinesUpdated = $true
    
            [System.Windows.Forms.MessageBox]::Show("Machine added successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $addMachineForm.Close()
        })
        $addMachineForm.Controls.Add($addButton)
    
        $addMachineForm.ShowDialog()
    })
    $callLogsPanel.Controls.Add($addMachineButton)

    # Add Send Logs Button
    $sendLogsButton = New-Object System.Windows.Forms.Button
    $sendLogsButton.Location = New-Object System.Drawing.Point(330, 440)
    $sendLogsButton.Size = New-Object System.Drawing.Size(150, 30)
    $sendLogsButton.Text = "Send Logs"
    $sendLogsButton.Add_Click({
        $sendLogsForm = New-Object System.Windows.Forms.Form
        $sendLogsForm.Text = "Send Logs"
        $sendLogsForm.Size = New-Object System.Drawing.Size(300, 200)
        $sendLogsForm.StartPosition = 'CenterScreen'

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
                }
                else {
                    Write-Log "Log file save cancelled by user"
                }
            }
            catch {
                Write-Log "Error saving log file: $_"
                [System.Windows.Forms.MessageBox]::Show("Error saving log file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }

            $sendLogsForm.Close()
        })
        $sendLogsForm.Controls.Add($sendButton)

        $sendLogsForm.ShowDialog()
    })
    $callLogsPanel.Controls.Add($sendLogsButton)

    Write-Log "Call Logs tab setup completed."
}

function Ensure-LaborLogsCsvExists {
    param (
        [string]$filePath
    )
   
    if (-not $filePath) {
        Write-Log "Error: Labor logs file path is not set"
        return
    }
    $headers = "Date,Work Order,Description,Machine,Duration,Notes"
    if (-not (Test-Path $filePath)) {
        $headers | Out-File -FilePath $filePath -Encoding UTF8
        Write-Log "Created new Labor Logs CSV file with headers at $filePath"
    } else {
        $content = Get-Content -Path $filePath
        if ($content.Count -eq 0) {
            $headers | Out-File -FilePath $filePath -Encoding UTF8
            Write-Log "Added headers to empty Labor Logs CSV file at $filePath"
        } elseif ($content[0] -ne $headers) {
            $headers | Set-Content -Path $filePath -Encoding UTF8
            $content | Select-Object -Skip 1 | Add-Content -Path $filePath -Encoding UTF8
            Write-Log "Updated headers in existing Labor Logs CSV file at $filePath"
        } else {
            Write-Log "Labor Logs CSV file exists and has correct headers at $filePath"
        }
    }
}

# Function to set up the Labor Log tab
function Setup-LaborLogTab {
    param($parentTab, $tabControl)

    Write-Log "Setting up Labor Log tab..."

    Ensure-LaborLogsCsvExists -filePath $laborLogsFilePath

    $laborLogPanel = New-Object System.Windows.Forms.Panel
    $laborLogPanel.Dock = 'Fill'
    $parentTab.Controls.Add($laborLogPanel)

    # Labor Log ListView
    $script:listViewLaborLog = New-Object System.Windows.Forms.ListView
    $script:listViewLaborLog.Location = New-Object System.Drawing.Point(10, 10)
    $script:listViewLaborLog.Size = New-Object System.Drawing.Size(850, 500)
    $script:listViewLaborLog.View = [System.Windows.Forms.View]::Details
    $script:listViewLaborLog.FullRowSelect = $true
    $script:listViewLaborLog.Scrollable = $true
    $script:listViewLaborLog.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::None)
    $script:listViewLaborLog.Columns.Add("Date", 100) | Out-Null            # Index 0
    $script:listViewLaborLog.Columns.Add("Work Order", 150) | Out-Null      # Index 1
    $script:listViewLaborLog.Columns.Add("Description", 300) | Out-Null     # Index 2
    $script:listViewLaborLog.Columns.Add("Machine", 100) | Out-Null         # Index 3
    $script:listViewLaborLog.Columns.Add("Duration", 100) | Out-Null        # Index 4
    $script:listViewLaborLog.Columns.Add("Parts", 300) | Out-Null           # Index 5
    $script:listViewLaborLog.Columns.Add("Notes", 150) | Out-Null           # Index 6
    
    # Tooltip for hovering
    $script:listViewLaborLog.MouseMove += {
        param($sender, $e)
        $item = $script:listViewLaborLog.GetItemAt($e.X, $e.Y)
        if ($item -ne $null) {
            $toolTipText = $item.SubItems[5].Text  # Assuming "Parts" is at index 5
            $script:listViewLaborLog.ToolTipText = $toolTipText
        } else {
            $script:listViewLaborLog.ToolTipText = ""
        }
    }

    # Double-Click for details
    $script:listViewLaborLog.DoubleClick += {
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $item = $selectedItems[0]
            $workOrderNumber = $item.SubItems[1].Text
            if ($script:workOrderParts.ContainsKey($workOrderNumber)) {
                $parts = $script:workOrderParts[$workOrderNumber]
                
                $detailsForm = New-Object System.Windows.Forms.Form
                $detailsForm.Text = "Parts Details for Work Order #$workOrderNumber"
                $detailsForm.Size = New-Object System.Drawing.Size(600, 400)
                
                $detailsListView = New-Object System.Windows.Forms.ListView
                $detailsListView.View = [System.Windows.Forms.View]::Details
                $detailsListView.Dock = 'Fill'
                $detailsListView.Columns.Add("Part Number", 100)
                $detailsListView.Columns.Add("OEM", 100)
                $detailsListView.Columns.Add("Quantity", 100)
                $detailsListView.Columns.Add("Location", 100)
                $detailsListView.Columns.Add("Source", 100)
                
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
    }
    
    # Initialize notification icon
    $script:notificationIcon = New-Object System.Windows.Forms.Label
    $script:notificationIcon.Text = "●"
    $script:notificationIcon.ForeColor = [System.Drawing.Color]::Red
    $script:notificationIcon.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $script:notificationIcon.Size = New-Object System.Drawing.Size(20, 20)
    $script:notificationIcon.Location = New-Object System.Drawing.Point(($tabControl.Width - 25), 5)
    $script:notificationIcon.Visible = $false
    $parentTab.Controls.Add($script:notificationIcon)

    # Load existing labor logs
    Load-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath

    # Add New Labor Log Entry Button
    $addLaborLogButton = New-Object System.Windows.Forms.Button
    $addLaborLogButton.Location = New-Object System.Drawing.Point(10, 520)
    $addLaborLogButton.Size = New-Object System.Drawing.Size(150, 30)
    $addLaborLogButton.Text = "Add Labor Log Entry"
    $addLaborLogButton.Add_Click({
        $addLaborLogForm = New-Object System.Windows.Forms.Form
        $addLaborLogForm.Text = "Add Labor Log Entry"
        $addLaborLogForm.Size = New-Object System.Drawing.Size(400, 400)
        $addLaborLogForm.StartPosition = 'CenterScreen'

        $labelDate = New-Object System.Windows.Forms.Label
        $labelDate.Location = New-Object System.Drawing.Point(10, 20)
        $labelDate.Size = New-Object System.Drawing.Size(100, 20)
        $labelDate.Text = "Date:"
        $addLaborLogForm.Controls.Add($labelDate)

        $textBoxDate = New-Object System.Windows.Forms.TextBox
        $textBoxDate.Location = New-Object System.Drawing.Point(120, 20)
        $textBoxDate.Size = New-Object System.Drawing.Size(250, 20)
        $textBoxDate.Text = Get-Date -Format "yyyy-MM-dd"
        $addLaborLogForm.Controls.Add($textBoxDate)

        $labelWorkOrder = New-Object System.Windows.Forms.Label
        $labelWorkOrder.Location = New-Object System.Drawing.Point(10, 50)
        $labelWorkOrder.Size = New-Object System.Drawing.Size(100, 20)
        $labelWorkOrder.Text = "Work Order #:"
        $addLaborLogForm.Controls.Add($labelWorkOrder)

        $textBoxWorkOrder = New-Object System.Windows.Forms.TextBox
        $textBoxWorkOrder.Location = New-Object System.Drawing.Point(120, 50)
        $textBoxWorkOrder.Size = New-Object System.Drawing.Size(250, 20)
        $addLaborLogForm.Controls.Add($textBoxWorkOrder)

        $labelTask = New-Object System.Windows.Forms.Label
        $labelTask.Location = New-Object System.Drawing.Point(10, 80)
        $labelTask.Size = New-Object System.Drawing.Size(100, 20)
        $labelTask.Text = "Task:"
        $addLaborLogForm.Controls.Add($labelTask)

        $textBoxTask = New-Object System.Windows.Forms.TextBox
        $textBoxTask.Location = New-Object System.Drawing.Point(120, 80)
        $textBoxTask.Size = New-Object System.Drawing.Size(250, 60)
        $textBoxTask.Multiline = $true
        $addLaborLogForm.Controls.Add($textBoxTask)

        $labelMachineId = New-Object System.Windows.Forms.Label
        $labelMachineId.Location = New-Object System.Drawing.Point(10, 150)
        $labelMachineId.Size = New-Object System.Drawing.Size(100, 20)
        $labelMachineId.Text = "Machine ID:"
        $addLaborLogForm.Controls.Add($labelMachineId)

        $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox
        $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 150)
        $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20)
        $addLaborLogForm.Controls.Add($comboBoxMachineId)
        Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"

        $labelDuration = New-Object System.Windows.Forms.Label
        $labelDuration.Location = New-Object System.Drawing.Point(10, 180)
        $labelDuration.Size = New-Object System.Drawing.Size(100, 20)
        $labelDuration.Text = "Duration:"
        $addLaborLogForm.Controls.Add($labelDuration)

        $textBoxDuration = New-Object System.Windows.Forms.TextBox
        $textBoxDuration.Location = New-Object System.Drawing.Point(120, 180)
        $textBoxDuration.Size = New-Object System.Drawing.Size(250, 20)
        $addLaborLogForm.Controls.Add($textBoxDuration)

        $labelNotes = New-Object System.Windows.Forms.Label
        $labelNotes.Location = New-Object System.Drawing.Point(10, 210)
        $labelNotes.Size = New-Object System.Drawing.Size(100, 20)
        $labelNotes.Text = "Notes:"
        $addLaborLogForm.Controls.Add($labelNotes)

        $textBoxNotes = New-Object System.Windows.Forms.TextBox
        $textBoxNotes.Location = New-Object System.Drawing.Point(120, 210)
        $textBoxNotes.Size = New-Object System.Drawing.Size(250, 60)
        $textBoxNotes.Multiline = $true
        $addLaborLogForm.Controls.Add($textBoxNotes)

        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Location = New-Object System.Drawing.Point(150, 300)
        $addButton.Size = New-Object System.Drawing.Size(100, 30)
        $addButton.Text = "Add"
        $addButton.Add_Click({
            $workOrderNumber = if ([string]::IsNullOrWhiteSpace($textBoxWorkOrder.Text)) { "Need W/O #" } else { $textBoxWorkOrder.Text }
            $item = New-Object System.Windows.Forms.ListViewItem($textBoxDate.Text)
            $item.SubItems.Add($workOrderNumber)
            $item.SubItems.Add($textBoxTask.Text)
            $item.SubItems.Add($comboBoxMachineId.SelectedItem)
            $item.SubItems.Add($textBoxDuration.Text)
            $item.SubItems.Add("")  # Empty parts column
            $item.SubItems.Add($textBoxNotes.Text)
            $script:listViewLaborLog.Items.Add($item)
            
            if ($workOrderNumber -eq "Need W/O #") {
                $key = "$($textBoxDate.Text)_$($comboBoxMachineId.SelectedItem)_$($textBoxTask.Text)"
                $script:unacknowledgedEntries[$key] = $true
                Update-NotificationIcon
            }
            
            # Save labor logs after adding a new entry
            Save-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath
            
            $addLaborLogForm.Close()
        })
        $addLaborLogForm.Controls.Add($addButton)

        $addLaborLogForm.ShowDialog()
    })
    $laborLogPanel.Controls.Add($addLaborLogButton)

    # Edit Labor Log Entry button
    $editLaborLogButton = New-Object System.Windows.Forms.Button
    $editLaborLogButton.Location = New-Object System.Drawing.Point(170, 520)
    $editLaborLogButton.Size = New-Object System.Drawing.Size(150, 30)
    $editLaborLogButton.Text = "Edit Labor Log Entry"
    $editLaborLogButton.Add_Click({
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $item = $selectedItems[0]
            $editLaborLogForm = New-Object System.Windows.Forms.Form
            $editLaborLogForm.Text = "Edit Labor Log Entry"
            $editLaborLogForm.Size = New-Object System.Drawing.Size(400, 400)
            $editLaborLogForm.StartPosition = 'CenterScreen'

            Write-Log "Editing Labor Log Entry. Default values:"
            Write-Log "Date: $($item.SubItems[0].Text)"
            Write-Log "Work Order Number: $($item.SubItems[1].Text)"
            Write-Log "Task: $($item.SubItems[2].Text)"
            Write-Log "Machine ID: $($item.SubItems[3].Text)"
            Write-Log "Duration: $($item.SubItems[4].Text)"
            Write-Log "Notes: $($item.SubItems[5].Text)"

            # Date
            $labelDate = New-Object System.Windows.Forms.Label
            $labelDate.Location = New-Object System.Drawing.Point(10, 20)
            $labelDate.Size = New-Object System.Drawing.Size(100, 20)
            $labelDate.Text = "Date:"
            $editLaborLogForm.Controls.Add($labelDate)

            $textBoxDate = New-Object System.Windows.Forms.TextBox
            $textBoxDate.Location = New-Object System.Drawing.Point(120, 20)
            $textBoxDate.Size = New-Object System.Drawing.Size(250, 20)
            $textBoxDate.Text = $item.SubItems[0].Text
            $editLaborLogForm.Controls.Add($textBoxDate)

            # Work Order Number
            $labelWorkOrder = New-Object System.Windows.Forms.Label
            $labelWorkOrder.Location = New-Object System.Drawing.Point(10, 50)
            $labelWorkOrder.Size = New-Object System.Drawing.Size(100, 20)
            $labelWorkOrder.Text = "Work Order #:"
            $editLaborLogForm.Controls.Add($labelWorkOrder)

            $textBoxWorkOrder = New-Object System.Windows.Forms.TextBox
            $textBoxWorkOrder.Location = New-Object System.Drawing.Point(120, 50)
            $textBoxWorkOrder.Size = New-Object System.Drawing.Size(250, 20)
            $textBoxWorkOrder.Text = $item.SubItems[1].Text
            $editLaborLogForm.Controls.Add($textBoxWorkOrder)

            # Task
            $labelTask = New-Object System.Windows.Forms.Label
            $labelTask.Location = New-Object System.Drawing.Point(10, 80)
            $labelTask.Size = New-Object System.Drawing.Size(100, 20)
            $labelTask.Text = "Task:"
            $editLaborLogForm.Controls.Add($labelTask)

            $textBoxTask = New-Object System.Windows.Forms.TextBox
            $textBoxTask.Location = New-Object System.Drawing.Point(120, 80)
            $textBoxTask.Size = New-Object System.Drawing.Size(250, 60)
            $textBoxTask.Multiline = $true
            $textBoxTask.Text = $item.SubItems[2].Text
            $editLaborLogForm.Controls.Add($textBoxTask)

            # Machine ID
            $labelMachineId = New-Object System.Windows.Forms.Label
            $labelMachineId.Location = New-Object System.Drawing.Point(10, 150)
            $labelMachineId.Size = New-Object System.Drawing.Size(100, 20)
            $labelMachineId.Text = "Machine ID:"
            $editLaborLogForm.Controls.Add($labelMachineId)

            $comboBoxMachineId = New-Object System.Windows.Forms.ComboBox
            $comboBoxMachineId.Location = New-Object System.Drawing.Point(120, 150)
            $comboBoxMachineId.Size = New-Object System.Drawing.Size(250, 20)
            $editLaborLogForm.Controls.Add($comboBoxMachineId)
            Load-ComboBoxData -comboBox $comboBoxMachineId -csvName "Machines"
            $comboBoxMachineId.Text = $item.SubItems[3].Text

            # Duration
            $labelDuration = New-Object System.Windows.Forms.Label
            $labelDuration.Location = New-Object System.Drawing.Point(10, 180)
            $labelDuration.Size = New-Object System.Drawing.Size(100, 20)
            $labelDuration.Text = "Duration:"
            $editLaborLogForm.Controls.Add($labelDuration)

            $textBoxDuration = New-Object System.Windows.Forms.TextBox
            $textBoxDuration.Location = New-Object System.Drawing.Point(120, 180)
            $textBoxDuration.Size = New-Object System.Drawing.Size(250, 20)
            $textBoxDuration.Text = $item.SubItems[4].Text
            $editLaborLogForm.Controls.Add($textBoxDuration)

            # Notes
            $labelNotes = New-Object System.Windows.Forms.Label
            $labelNotes.Location = New-Object System.Drawing.Point(10, 210)
            $labelNotes.Size = New-Object System.Drawing.Size(100, 20)
            $labelNotes.Text = "Notes:"
            $editLaborLogForm.Controls.Add($labelNotes)

            $textBoxNotes = New-Object System.Windows.Forms.TextBox
            $textBoxNotes.Location = New-Object System.Drawing.Point(120, 210)
            $textBoxNotes.Size = New-Object System.Drawing.Size(250, 60)
            $textBoxNotes.Multiline = $true
            $textBoxNotes.Text = $item.SubItems[5].Text
            $editLaborLogForm.Controls.Add($textBoxNotes)

            $saveButton = New-Object System.Windows.Forms.Button
            $saveButton.Location = New-Object System.Drawing.Point(150, 300)
            $saveButton.Size = New-Object System.Drawing.Size(100, 30)
            $saveButton.Text = "Save"
            $saveButton.Add_Click({
                # Update item with new values
                $item.SubItems[0].Text = $textBoxDate.Text
                $item.SubItems[1].Text = $textBoxWorkOrder.Text
                $item.SubItems[2].Text = $textBoxTask.Text
                $item.SubItems[3].Text = $comboBoxMachineId.Text
                $item.SubItems[4].Text = $textBoxDuration.Text
                $item.SubItems[5].Text = $textBoxNotes.Text

                # Handle acknowledgment and save logs
                $key = "$($textBoxDate.Text)_$($comboBoxMachineId.Text)_$($textBoxTask.Text)"
                if ($textBoxWorkOrder.Text -eq "Need W/O #") {
                    $script:unacknowledgedEntries[$key] = $true
                } else {
                    $script:unacknowledgedEntries.Remove($key)
                }
                Update-NotificationIcon

                # Save labor logs after editing an entry
                Save-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath

                $editLaborLogForm.Close()
            })
            $editLaborLogForm.Controls.Add($saveButton)

            $editLaborLogForm.ShowDialog()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select an entry to edit.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "Attempted to edit Labor Log Entry without selection"
        }
    })
    $laborLogPanel.Controls.Add($editLaborLogButton)

    # Add Parts to Work Order Button
    $addPartsButton = New-Object System.Windows.Forms.Button
    $addPartsButton.Location = New-Object System.Drawing.Point(490, 520)
    $addPartsButton.Size = New-Object System.Drawing.Size(150, 30)
    $addPartsButton.Text = "Add Parts to Work Order"
    $addPartsButton.Add_Click({
        $selectedItems = $script:listViewLaborLog.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $workOrderItem = $selectedItems[0]
            $workOrderNumber = $workOrderItem.SubItems[1].Text

            # Call the function to add parts to the selected work order
            Add-PartsToWorkOrder -workOrderNumber $workOrderNumber
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a work order to add parts.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $laborLogPanel.Controls.Add($addPartsButton)

    # Add a Refresh button to manually reload the labor logs
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(330, 520)
    $refreshButton.Size = New-Object System.Drawing.Size(150, 30)
    $refreshButton.Text = "Refresh Labor Logs"
    $refreshButton.Add_Click({
        $script:listViewLaborLog.Items.Clear()
        Load-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath
        Process-HistoricalLogs
        Write-Log "Labor Logs manually refreshed"
    })
    $laborLogPanel.Controls.Add($refreshButton)

    

    Write-Log "Labor Log tab setup completed."
}

# A hashtable to store processed call logs to avoid duplication
$script:processedCallLogs = @{}
function Process-HistoricalLogs {
    Write-Log "Processing historical Call Logs to create Labor Logs if necessary..."
   
    if ($null -eq $script:listViewLaborLog) {
        Write-Log "Error: Labor Log ListView is not initialized. Cannot process historical logs."
        return
    }

    Ensure-LaborLogsCsvExists -filePath $global:laborLogsFilePath
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

# Function to load Labor Logs from CSV
function Load-LaborLogs {
    param($listView, $filePath)
    
    Write-Log "=== START Load-LaborLogs ==="
    Write-Log "Loading labor logs from: $filePath"
    
    if (-not (Test-Path $filePath)) {
        Write-Log "ERROR: Labor logs file not found at: $filePath"
        return
    }
    
    try {
        $laborLogs = Import-Csv -Path $filePath
        Write-Log "Successfully loaded labor logs. Entry count: $($laborLogs.Count)"
        
        # Initialize workOrderParts dictionary if needed
        if ($null -eq $script:workOrderParts) {
            Write-Log "Initializing workOrderParts dictionary"
            $script:workOrderParts = @{}
        }
        
        Write-Log "Processing labor log entries..."
        foreach ($log in $laborLogs) {
            $workOrderNumber = $log.'Work Order'
            Write-Log "Processing work order: $workOrderNumber"
            
            $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
            $item.SubItems.Add($workOrderNumber)
            $item.SubItems.Add($log.Description)
            $item.SubItems.Add($log.Machine)
            $item.SubItems.Add($log.Duration)
            
            # First add empty Parts column
            $partsDisplay = ""
            
            # Handle the Parts column
            if ($log.PSObject.Properties.Name -contains 'Parts' -and -not [string]::IsNullOrWhiteSpace($log.Parts)) {
                Write-Log "Work order has parts data: $($log.Parts)"
                try {
                    $parts = $log.Parts | ConvertFrom-Json
                    Write-Log "Successfully parsed JSON parts data. Part count: $($parts.Count)"
                    
                    # Store parts in workOrderParts dictionary
                    $script:workOrderParts[$workOrderNumber] = $parts
                    
                    # Create a formatted string for display
                    $partsDisplay = ($parts | ForEach-Object { 
                        "$($_.PartNumber) - $($_.PartNo) - Qty:$($_.Quantity)" 
                    }) -join ", "
                    
                    Write-Log "Parts display string: $partsDisplay"
                } catch {
                    Write-Log "ERROR parsing Parts JSON for work order ${workOrderNumber}: $($_.Exception.Message)"
                    $partsDisplay = $log.Parts
                }
            }
            
            # Now add Parts and Notes in the correct order
            $item.SubItems.Add($partsDisplay)  # Add parts at index 5
            $item.SubItems.Add($log.Notes)     # Add notes at index 6
            
            $listView.Items.Add($item)
            Write-Log "Added item to list view for work order: $workOrderNumber"
        }
        
        Write-Log "Finished loading labor logs. List view now has $($listView.Items.Count) items"
        Write-Log "=== END Load-LaborLogs ==="
    }
    catch {
        Write-Log "ERROR in Load-LaborLogs: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
    }
}

# Moving Calls from Calls to Labor Log
function Add-LaborLogEntryFromCallLog {
    param ($log)
    try {
        if ($null -eq $script:listViewLaborLog) {
            Write-Log "Error: Labor Log ListView is not initialized. Cannot add entry."
            return
        }

        $duration = Get-TimeDifference -startTime $log.'Time Down' -endTime $log.'Time Up'
        $durationHours = [Math]::Round($duration / 60, 2)
       
        $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
        $item.SubItems.Add("Need W/O #")
        $item.SubItems.Add("$($log.Cause) / $($log.Action) / $($log.Noun)")
        $item.SubItems.Add($log.Machine)
        $item.SubItems.Add($durationHours.ToString("F2"))
        $item.SubItems.Add("")  # Add empty parts column
        $item.SubItems.Add($log.Notes)  # Now notes are at index 6
        
        $script:listViewLaborLog.Items.Add($item)
        Write-Log "Added Labor Log entry: Date=$($log.Date), Machine=$($log.Machine), Duration=$durationHours hours"
        
        # Save labor logs after adding a new entry
        Save-LaborLogs -listView $script:listViewLaborLog -filePath $global:laborLogsFilePath
    } catch {
        Write-Log "Error adding Labor Log entry: $_"
    }
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
            $item.SubItems.Add($part.Description)
            $item.SubItems.Add($part.Quantity)
            $item.SubItems.Add($part.Location)
            $item.SubItems.Add($part.OEMNumber)
            $item.SubItems.Add($part.Source)
            $item.Tag = $part  # Store the full part object for later use
            $resultsListView.Items.Add($item)
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
                $newItem.SubItems.Add($partObj.Description)
                $newItem.SubItems.Add($qty)
                $newItem.SubItems.Add($partObj.Source)
                $newItem.SubItems.Add($partObj.Location)
                $newItem.Tag = [PSCustomObject]@{
                    PartNumber = $partObj.PartNumber
                    Description = $partObj.Description
                    Quantity = $qty
                    Source = $partObj.Source
                    Location = $partObj.Location
                    OEMNumber = $partObj.OEMNumber
                }
                $selectedListView.Items.Add($newItem)
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

# Simple helper function to search parts
function Search-Parts {
    param(
        [string]$NSN,
        [string]$Description
    )
    
    $results = @()
    
    # 1. Search Local Parts Room
    $partsRoomCsv = Join-Path $config.PartsRoomDirectory "*.csv" 
    $localParts = Get-ChildItem -Path $partsRoomCsv | ForEach-Object {
        $parts = Import-Csv -Path $_.FullName
        
        # Filter based on search criteria
        $parts | Where-Object {
            ($NSN -eq "" -or $_.'Part (NSN)' -like "*$NSN*") -and
            ($Description -eq "" -or $_.Description -like "*$Description*") -and
            ([int]$_.QTY -gt 0)  # Only show parts with stock
        } | ForEach-Object {
            [PSCustomObject]@{
                PartNumber = $_.'Part (NSN)'
                Description = $_.Description
                Quantity = [int]$_.QTY
                Location = $_.Location
                OEMNumber = $_.'OEM 1'
                Source = "Local Parts Room"
            }
        }
    }
    
    $results += $localParts
    
    # 2. Search Same Day Parts Room (similar pattern to above)
    $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
    if (Test-Path $sameDayDir) {
        $sameDayParts = Get-ChildItem -Path "$sameDayDir\*.csv" | ForEach-Object {
            $siteName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
            $parts = Import-Csv -Path $_.FullName
            
            # Filter based on search criteria
            $parts | Where-Object {
                ($NSN -eq "" -or $_.'Part (NSN)' -like "*$NSN*") -and
                ($Description -eq "" -or $_.Description -like "*$Description*") -and
                ([int]$_.QTY -gt 0)  # Only show parts with stock
            } | ForEach-Object {
                [PSCustomObject]@{
                    PartNumber = $_.'Part (NSN)'
                    Description = $_.Description
                    Quantity = [int]$_.QTY
                    Location = $_.Location
                    OEMNumber = $_.'OEM 1'
                    Source = "$siteName (Same Day)"
                }
            }
        }
        
        $results += $sameDayParts
    }
    
    return $results
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
        $newWorkOrderNumber = "Need W/O #-$($targetIndex + 1)"
        Write-Log "Updating work order from '$WorkOrderNumber' to '$newWorkOrderNumber'"
        $targetItem.SubItems[1].Text = $newWorkOrderNumber
        $WorkOrderNumber = $newWorkOrderNumber  # Update for dictionary storage
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
        $form.Close()  # Close the form after saving
        return $true
    }
    catch {
        Write-Log "ERROR in Save-PartsToWorkOrder: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("Error saving parts to work order: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# data retrieval logic
function Search-CrossReferenceData {
    param ($NSN, $OEM, $Description)
    
    $crossRefResults = @()

    foreach ($book in $config.Books.PSObject.Properties) {
        $bookName = $book.Name
        $bookDir = Join-Path $config.PartsBooksDirectory $bookName
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
                    Source = $sourceFileName  # This will be the CSV filename without extension
                }
                $crossRefResults += $resultItem
            }
        }
    }

    return $crossRefResults
}

# Main execution
Show-MainForm
Write-Log "Application closed"

Read-Host
