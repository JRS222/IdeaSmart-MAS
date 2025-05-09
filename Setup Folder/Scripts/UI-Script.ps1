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

################################################################################
#                           Global Variables                                   #
################################################################################

<# 
# Configuration variables
$global:config = $null                    # Configuration object
$global:configPath = $null                # Path to configuration file

# File paths
$script:laborLogsFilePath = $null         # Path to labor logs file
$script:callLogsFilePath = $null          # Path to call logs file
$script:logPath = "UI.log"                # Path to UI log file

# State tracking
$script:unacknowledgedEntries = @{}       # Unacknowledged entry tracking
$script:processedCallLogs = @{}           # Processed call logs tracking
$script:workOrderParts = @{}              # Parts for work orders
$script:machinesUpdated = $false          # Flag for machines list update

# Labor and Call Logs UI elements
$script:listViewLaborLog = New-Object System.Windows.Forms.ListView  # Labor logs
$script:listViewCallLogs = $null          # Call logs list view
$script:notificationIcon = $null          # Notification icon

# Search UI elements
$script:textBoxNSN = $null                # NSN search text box
$script:textBoxOEM = $null                # OEM search text box
$script:textBoxDescription = $null        # Description search text box
$script:listViewAvailability = $null      # Availability results
$script:listViewSameDayAvailability = $null  # Same-day availability
$script:listViewCrossRef = $null          # Cross-reference results
$script:openFiguresButton = $null         # Open figures button
$script:takePartOutButton = $null         # Take part out button

# Tab controls
$script:tabControl = $null                # Main tab control
$script:partsBookTab = $null              # Parts book tab
$script:callLogsTab = $null               # Call logs tab
$script:laborLogTab = $null               # Labor log tab
$script:searchTab = $null                 # Search tab
$script:actionsTab = $null                # Actions tab 
#>

################################################################################
#                            Core Utilities                                    #
################################################################################

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

# Data retrieval logic
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
        $rowCount = $csvData.Count
        $colCount = $headers.Count
        
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
                $percent = 30 + ($rowIdx / $rowCount * 20)  # Scale from 30% to 50%
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
        $descriptionColumn = $listObject.ListColumns | Where-Object { $_.Name -eq "Description" }
        if ($descriptionColumn) {
            $descriptionColumn.Range.Offset(1, 0).HorizontalAlignment = -4131 # xlLeft
        }

        # Remove any columns named "Importing data..."
        for ($col = $headers.Count; $col -ge 1; $col--) {
            $columnHeader = $worksheet.Cells.Item(1, $col).Value2
            if ($columnHeader -eq "Importing data...") {
                $column = $worksheet.Columns.Item($col)
                $column.Delete()
                Write-Log "Removed 'Importing data...' column"
            }
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
        if ($excel) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
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

################################################################################
#                     Labor and Call Logs Management                           #
################################################################################

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

# Enhanced Load-LaborLogs function with debugging
function Load-LaborLogs {
    param(
        [System.Windows.Forms.ListView]$listView,
        [string]$filePath
    )
    
    Write-Log "=== START Load-LaborLogs ==="
    Write-Log "Loading labor logs from: $filePath"
    
    if (-not (Test-Path $filePath)) {
        Write-Log "ERROR: Labor logs file not found at: $filePath"
        return
    }
    
    try {
        # Import the labor logs CSV
        $laborLogs = Import-Csv -Path $filePath
        Write-Log "Successfully loaded labor logs. Entry count: $($laborLogs.Count)"
        
        # Initialize the workOrderParts dictionary if it doesn't exist
        if ($null -eq $script:workOrderParts) {
            Write-Log "Initializing workOrderParts dictionary"
            $script:workOrderParts = @{}
        }
        
        # Clear the ListView before loading new data
        $listView.Items.Clear()
        
        # Process each labor log entry
        foreach ($log in $laborLogs) {
            $workOrderNumber = $log.'Work Order'
            Write-Log "Processing work order: $workOrderNumber"
            
            # Create a new ListView item for the log entry
            $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
            $item.SubItems.Add($workOrderNumber)
            $item.SubItems.Add($log.Description)
            $item.SubItems.Add($log.Machine)
            $item.SubItems.Add($log.Duration)
            $item.SubItems.Add($log.Notes)
            
            # Handle the Parts column
            if ($log.PSObject.Properties.Name -contains 'Parts' -and -not [string]::IsNullOrWhiteSpace($log.Parts)) {
                Write-Log "Work order has parts data: $($log.Parts)"
                
                try {
                    # Deserialize the JSON parts data
                    $parts = $log.Parts | ConvertFrom-Json
                    Write-Log "Successfully parsed JSON parts data. Part count: $($parts.Count)"
                    
                    # Store the parts in the workOrderParts dictionary
                    $script:workOrderParts[$workOrderNumber] = $parts
                    
                    # Create a formatted string for display in the ListView
                    $partsDisplay = ($parts | ForEach-Object { 
                        "$($_.PartNumber) - $($_.PartNo) - Qty:$($_.Quantity)" 
                    }) -join ", "
                    
                    Write-Log "Parts display string: $partsDisplay"
                    $item.SubItems.Add($partsDisplay)
                } catch {
                    Write-Log "ERROR parsing Parts JSON for work order ${workOrderNumber}: $($_.Exception.Message)"
                    $item.SubItems.Add("Invalid Parts Data")  # Add a placeholder for invalid data
                }
            } else {
                Write-Log "Work order has no parts data"
                $item.SubItems.Add("")  # Add an empty column if no parts data exists
            }
            
            # Add the item to the ListView
            $listView.Items.Add($item)
            Write-Log "Added item to list view for work order: $workOrderNumber"
        }
        
        Write-Log "Finished loading labor logs. List view now has $($listView.Items.Count) items"
        Write-Log "=== END Load-LaborLogs ==="
    }
    catch {
        Write-Log "ERROR in Load-LaborLogs: $($_.Exception.Message)"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("Error loading labor logs: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# A hashtable to store processed call logs to avoid duplication
$script:processedCallLogs = @{}
function Process-HistoricalLogs {
    Write-Log "Processing historical Call Logs to create Labor Logs if necessary..."
   
    if ($null -eq $script:listViewLaborLog) {
        Write-Log "Error: Labor Log ListView is not initialized. Cannot process historical logs."
        return
    }

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
       
        $workOrderNumber = "Need W/O #-$(New-Guid)"
        $item = New-Object System.Windows.Forms.ListViewItem($log.Date)
        $item.SubItems.Add($workOrderNumber)
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

################################################################################
#                          Parts Management                                    #
################################################################################

# Function to take a part out
function Take-PartOut {
    Write-Log "Taking a part out..."
    [System.Windows.Forms.MessageBox]::Show("Take Part Out process not implemented yet.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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
        $newWorkOrderNumber = "Need W/O #-$(New-Guid)"
        $targetItem.SubItems[1].Text = $newWorkOrderNumber
        $WorkOrderNumber = $newWorkOrderNumber
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

# Function to set up the Labor Log tab
function Setup-LaborLogTab {
    param($parentTab, $tabControl)

    Write-Log "Setting up Labor Log tab..."

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
    $script:listViewLaborLog.Columns.Add("Date", 100) | Out-Null
    $script:listViewLaborLog.Columns.Add("Work Order", 150) | Out-Null
    $script:listViewLaborLog.Columns.Add("Description", 300) | Out-Null
    $script:listViewLaborLog.Columns.Add("Machine", 100) | Out-Null
    $script:listViewLaborLog.Columns.Add("Duration", 100) | Out-Null
    $script:listViewLaborLog.Columns.Add("Parts", 300) | Out-Null  # Increase width to 300 pixels
    $script:listViewLaborLog.Columns.Add("Notes", 150) | Out-Null
    $laborLogPanel.Controls.Add($script:listViewLaborLog)
    
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
    $script:listViewAvailability.Size = New-Object System.Drawing.Size(1100, 150)
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
    $script:listViewSameDayAvailability.Size = New-Object System.Drawing.Size(1100, 150)
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
    $script:listViewCrossRef.Size = New-Object System.Drawing.Size(1100, 150)
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
            $matchesNSN = if ($nsnSearch -eq '') {
                $true
            } else {
                # Remove hyphens from both the search term and the NSN for comparison
                $cleanNSN = $_.'Part (NSN)' -replace '-', ''
                $cleanSearch = $nsnSearch -replace '-', ''
                # Use contains rather than exact match
                $cleanNSN -like "*$cleanSearch*"
            }
        
            $matchesOEM = ($oemSearch -eq '') -or 
                          ($_.'OEM 1' -like "*$oemSearch*") -or 
                          ($_.'OEM 2' -like "*$oemSearch*") -or 
                          ($_.'OEM 3' -like "*$oemSearch*")
            
            $matchesDesc = ($descriptionSearch -eq '') -or 
                           ($_.Description -like "*$descriptionSearch*")
            
            $matchesNSN -and $matchesOEM -and $matchesDesc
        }

        Write-Log "Found $($filteredData.Count) matching records in Availability."

        $script:listViewAvailability.Items.Clear()
        
        if ($filteredData.Count -gt 0) {
            foreach ($row in $filteredData) {
                Write-Log "Adding row to listview: $($row.'Part (NSN)')"  # Add debug logging
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
                    # Write-Log "Processed $($csvData.Count) rows from $($csvFile.Name)"
                } catch {
                    Write-Log "Failed to read CSV file $($csvFile.FullName). Error: $_"
                }
            }

            $filteredSameDayData = $sameDayData | Where-Object {
                ($nsnSearch -eq '' -or $_.'Part (NSN)' -like "*$nsnSearch*") -and
                ($oemSearch -eq '' -or ($_.'OEM 1' -like "*$oemSearch*" -or $_.'OEM 2' -like "*$oemSearch*" -or $_.'OEM 3' -like "*$oemSearch*")) -and
                ($descriptionSearch -eq '' -or $_.Description -like "*$descriptionSearch*")
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
                        $item | Add-Member -NotePropertyName 'Handbook' -NotePropertyValue $bookName -Force
                        $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $sectionName -Force
                        $crossRefResults += $item
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
    
    # Get all the parts books from the configuration
    if (-not $config.Books -or $config.Books.PSObject.Properties.Count -eq 0) {
        Write-Log "No parts books found in configuration"
        [System.Windows.Forms.MessageBox]::Show("No parts books found in configuration", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
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

################################################################################
#                     Parts Room and Books Management                          #
################################################################################

# Function to update Parts Books with the latest inventory data
function Update-PartsBooks {
    param(
        [string]$sourceCSVPath
    )
    
    Write-Log "Starting Parts Books update process..."
    
    # Show progress form
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Updating Parts Books"
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = 'CenterScreen'
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(370, 20)
    $progressLabel.Text = "Loading source data..."
    $progressForm.Controls.Add($progressLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 50)
    $progressBar.Size = New-Object System.Drawing.Size(370, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    $progressForm.Controls.Add($progressBar)
    
    $bookLabel = New-Object System.Windows.Forms.Label
    $bookLabel.Location = New-Object System.Drawing.Point(10, 80)
    $bookLabel.Size = New-Object System.Drawing.Size(370, 20)
    $bookLabel.Text = ""
    $progressForm.Controls.Add($bookLabel)
    
    # Show the progress form
    $progressForm.Show()
    $progressForm.Refresh()
    
    if (-not (Test-Path $sourceCSVPath)) {
        Write-Log "Error: Source CSV file not found at $sourceCSVPath"
        [System.Windows.Forms.MessageBox]::Show("Source CSV file not found at $sourceCSVPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }
    
    # Load the source data
    $sourceData = Import-Csv -Path $sourceCSVPath
    
    # Create dictionaries for faster lookups
    $nsnDict = @{}
    $oemDict = @{}
    
    foreach ($part in $sourceData) {
        # Store by NSN
        if (-not [string]::IsNullOrEmpty($part.'Part (NSN)')) {
            $nsnDict[$part.'Part (NSN)'] = $part
        }
        
        # Store by OEM numbers
        if (-not [string]::IsNullOrEmpty($part.'OEM 1')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 1'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) {
                $oemDict[$normalizedOEM] = $part
            }
        }
        if (-not [string]::IsNullOrEmpty($part.'OEM 2')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 2'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) {
                $oemDict[$normalizedOEM] = $part
            }
        }
        if (-not [string]::IsNullOrEmpty($part.'OEM 3')) {
            $normalizedOEM = Normalize-OEM -oem $part.'OEM 3'
            if (-not [string]::IsNullOrEmpty($normalizedOEM)) {
                $oemDict[$normalizedOEM] = $part
            }
        }
    }
    
    $progressLabel.Text = "Loaded $($sourceData.Count) parts from source CSV"
    $progressForm.Refresh()
    Write-Log "Loaded $($sourceData.Count) parts, $($nsnDict.Count) unique NSNs, $($oemDict.Count) unique OEMs"
    
    # Get all the parts books from the configuration
    if (-not $config.Books -or $config.Books.PSObject.Properties.Count -eq 0) {
        Write-Log "No parts books found in configuration"
        [System.Windows.Forms.MessageBox]::Show("No parts books found in configuration", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }
    
    # Collect books to process
    $booksToProcess = @()
    foreach ($bookProp in $config.Books.PSObject.Properties) {
        $bookName = $bookProp.Name
        $bookDir = Join-Path $config.PartsBooksDirectory $bookName
        $combinedSectionsDir = Join-Path $bookDir "CombinedSections"
        
        if (Test-Path $combinedSectionsDir) {
            $booksToProcess += @{
                Name = $bookName
                Directory = $bookDir
                CombinedSectionsDir = $combinedSectionsDir
                ExcelPath = Join-Path $bookDir "$bookName.xlsx"
            }
        } else {
            Write-Log "Warning: CombinedSections directory not found for book: $bookName"
        }
    }
    
    $totalBooks = $booksToProcess.Count
    if ($totalBooks -eq 0) {
        Write-Log "No parts books with combined sections found to update"
        [System.Windows.Forms.MessageBox]::Show("No parts books with combined sections found to update", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressForm.Close()
        return $false
    }
    
    # Set up counters and progress
    $progressBar.Maximum = $totalBooks
    $progressBar.Value = 0
    $totalUpdatedCSVs = 0
    $totalUpdatedParts = 0
    
    # Process each book
    for ($bookIndex = 0; $bookIndex -lt $totalBooks; $bookIndex++) {
        $book = $booksToProcess[$bookIndex]
        $progressBar.Value = $bookIndex
        $bookLabel.Text = "Processing book: $($book.Name)"
        $progressLabel.Text = "Scanning sections..."
        $progressForm.Refresh()
        
        Write-Log "Processing book: $($book.Name)"
        
        # Get all section CSV files in this book
        $sectionFiles = Get-ChildItem -Path $book.CombinedSectionsDir -Filter "Section *.csv"
        if ($sectionFiles.Count -eq 0) {
            Write-Log "No section CSV files found for book: $($book.Name)"
            continue
        }
        
        $sectionsWithChanges = @()
        $bookUpdatedParts = 0
        
        # Process each section CSV file
        foreach ($sectionFile in $sectionFiles) {
            $sectionName = $sectionFile.BaseName
            $progressLabel.Text = "Processing section: $sectionName"
            $progressForm.Refresh()
            
            try {
                # Load the section CSV
                $sectionData = Import-Csv -Path $sectionFile.FullName
                $sectionUpdated = $false
                $sectionUpdatedParts = 0
                
                # Process each part in the section
                foreach ($part in $sectionData) {
                    $stockNo = $part.'STOCK NO.'
                    $partNo = $part.'PART NO.'
                    
                    # Try to match by NSN first
                    if (-not [string]::IsNullOrEmpty($stockNo) -and $stockNo -ne "NSL") {
                        if ($nsnDict.ContainsKey($stockNo)) {
                            $sourcePart = $nsnDict[$stockNo]
                            
                            # Update QTY and Location if different
                            if ($part.QTY -ne $sourcePart.QTY -or $part.Location -ne $sourcePart.Location) {
                                $part.QTY = $sourcePart.QTY
                                $part.Location = $sourcePart.Location
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        } else {
                            # NSN not found in current inventory
                            if ($part.QTY -ne "0" -or $part.Location -ne "Not in current inventory") {
                                $part.QTY = "0"
                                $part.Location = "Not in current inventory"
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        }
                    }
                    # If NSN didn't work, try to match by OEM number
                    elseif (-not [string]::IsNullOrEmpty($partNo)) {
                        $normalizedPartNo = Normalize-OEM -oem $partNo
                        if (-not [string]::IsNullOrEmpty($normalizedPartNo) -and $oemDict.ContainsKey($normalizedPartNo)) {
                            $sourcePart = $oemDict[$normalizedPartNo]
                            
                            # Update QTY and Location if different
                            if ($part.QTY -ne $sourcePart.QTY -or $part.Location -ne $sourcePart.Location) {
                                $part.QTY = $sourcePart.QTY
                                $part.Location = $sourcePart.Location
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        } else {
                            # OEM not found in current inventory
                            if ($part.QTY -ne "0" -or $part.Location -ne "Not in current inventory") {
                                $part.QTY = "0"
                                $part.Location = "Not in current inventory"
                                $sectionUpdated = $true
                                $sectionUpdatedParts++
                            }
                        }
                    }
                }
                
                # Save the updated section if changes were made
                if ($sectionUpdated) {
                    $sectionData | Export-Csv -Path $sectionFile.FullName -NoTypeInformation
                    $sectionsWithChanges += $sectionName
                    $bookUpdatedParts += $sectionUpdatedParts
                    $totalUpdatedCSVs++
                    $totalUpdatedParts += $sectionUpdatedParts
                    Write-Log "Updated section $sectionName with $sectionUpdatedParts changes"
                }
            } catch {
                Write-Log "Error processing section $sectionName : $($_.Exception.Message)"
            }
        }
        
        # Update the Excel workbook if it exists and if sections were updated
        if ($sectionsWithChanges.Count -gt 0 -and (Test-Path $book.ExcelPath)) {
            $progressLabel.Text = "Updating Excel workbook..."
            $bookLabel.Text = "Processing book: $($book.Name) - Excel update"
            $progressForm.Refresh()
            
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                
                $workbook = $excel.Workbooks.Open($book.ExcelPath)
                
                # Calculate base progress and progress weight for this phase
                $baseProgress = $bookIndex / $totalBooks * 100
                $progressWeight = 100 / $totalBooks / 2  # Half of book's progress weight for Excel
                
                # Track total operations to perform
                $totalOperations = $sectionsWithChanges.Count * 10  # Rough estimate of operations per section
                $currentOperation = 0
                
                # Update each worksheet that corresponds to a section with changes
                foreach ($sectionName in $sectionsWithChanges) {
                    try {
                        # Update progress for section start
                        $currentOperation += 5  # Increment for starting section
                        $sectionProgress = $currentOperation / $totalOperations * $progressWeight
                        $totalProgress = $baseProgress + $sectionProgress
                        $progressBar.Value = [Math]::Min([int]$totalProgress, 100)
                        $progressLabel.Text = "Processing section: $sectionName"
                        $progressForm.Refresh()
                        
                        # Create a mapping of possible truncated names to full section names
                        $possibleNames = @()
                        $possibleNames += $sectionName  # Original name
                        
                        # Add truncated version (Excel limits worksheet names to 31 chars)
                        $truncatedName = $sectionName.Substring(0, [Math]::Min(31, $sectionName.Length)) -replace '[:\\/?*\[\]]', ''
                        $possibleNames += $truncatedName
                        
                        # Also look for section number only (e.g., "Section 1")
                        if ($sectionName -match '^(Section \d+)') {
                            $possibleNames += $matches[1]
                        }
                        
                        # Try to find the worksheet with any of the possible names
                        $worksheet = $null
                        foreach ($nameVariant in $possibleNames) {
                            try {
                                $worksheet = $workbook.Worksheets.Item($nameVariant)
                                Write-Log "Found worksheet using name variant: $nameVariant"
                                break  # Exit loop if worksheet found
                            } catch {
                                # Continue to next name variant
                                continue
                            }
                        }
                        
                        # If still not found, try fuzzy matching
                        if ($worksheet -eq $null) {
                            Write-Log "Could not find exact worksheet match for $sectionName, trying fuzzy matching..."
                            foreach ($ws in $workbook.Worksheets) {
                                # Check if worksheet name starts with the section number
                                if ($sectionName -match '^(Section \d+)' -and 
                                    $ws.Name -match "^$($matches[1])") {
                                    $worksheet = $ws
                                    Write-Log "Found worksheet using fuzzy match: $($ws.Name)"
                                    break
                                }
                            }
                        }
                        
                        # Update progress for finding the worksheet
                        $currentOperation += 5
                        $progressBar.Value = [Math]::Min([int]($baseProgress + $currentOperation / $totalOperations * $progressWeight), 100)
                        $progressForm.Refresh()
                        
                        if ($worksheet -ne $null) {
                            # Find the QTY and Location columns
                            $qtyCol = $null
                            $locationCol = $null
                            
                            # Get column indices
                            $lastCol = 20  # Reasonable limit
                            for ($col = 1; $col -le $lastCol; $col++) {
                                $colName = $worksheet.Cells.Item(1, $col).Text
                                if ($colName -eq "QTY") {
                                    $qtyCol = $col
                                } elseif ($colName -eq "LOCATION") {
                                    $locationCol = $col
                                }
                                
                                # Once we found both columns, we can break
                                if ($qtyCol -and $locationCol) {
                                    break
                                }
                            }
                            
                            # Find the STOCK NO. and PART NO. columns
                            $stockNoCol = $null
                            $partNoCol = $null
                            
                            for ($col = 1; $col -le $lastCol; $col++) {
                                $colName = $worksheet.Cells.Item(1, $col).Text
                                if ($colName -eq "STOCK NO.") {
                                    $stockNoCol = $col
                                } elseif ($colName -eq "PART NO.") {
                                    $partNoCol = $col
                                }
                                
                                # Once we found both columns, we can break
                                if ($stockNoCol -and $partNoCol) {
                                    break
                                }
                            }
                            
                            # We need at least stock or part number columns
                            if ($stockNoCol -or $partNoCol) {
                                # We need QTY and Location columns to update
                                if ($qtyCol -and $locationCol) {
                                    # Load the section data for reference
                                    $sectionFile = Get-ChildItem -Path $book.CombinedSectionsDir -Filter "$sectionName.csv" | Select-Object -First 1
                                    if ($sectionFile) {
                                        $sectionData = Import-Csv -Path $sectionFile.FullName
                                        
                                        # Create a lookup dictionary by stock number and part number
                                        $sectionDict = @{}
                                        foreach ($part in $sectionData) {
                                            $stockNo = $part.'STOCK NO.'
                                            if (-not [string]::IsNullOrEmpty($stockNo)) {
                                                $sectionDict[$stockNo] = $part
                                            }
                                            
                                            $partNo = $part.'PART NO.'
                                            if (-not [string]::IsNullOrEmpty($partNo)) {
                                                $sectionDict[$partNo] = $part
                                            }
                                        }
                                        
                                        # Now update the cells in the worksheet
                                        $lastRow = $worksheet.UsedRange.Rows.Count
                                        for ($row = 2; $row -le $lastRow; $row++) {
                                            # Update progress every 10 rows for performance
                                            if (($row % 10) -eq 0 -or $row -eq $lastRow) {
                                                $rowProgress = ($row - 2) / ($lastRow - 2) * 10  # Scale to progress units
                                                $currentOperation += $rowProgress
                                                $cellProgress = $currentOperation / $totalOperations * $progressWeight
                                                $totalProgress = $baseProgress + $cellProgress
                                                
                                                $progressBar.Value = [Math]::Min([int]$totalProgress, 100)
                                                $progressLabel.Text = "Processing section: $sectionName - Row $row of $lastRow"
                                                $progressForm.Refresh()
                                                [System.Windows.Forms.Application]::DoEvents()
                                            }
                                            
                                            $key = $null
                                            
                                            # Try to get the key from STOCK NO. first
                                            if ($stockNoCol) {
                                                $stockNo = $worksheet.Cells.Item($row, $stockNoCol).Text
                                                if (-not [string]::IsNullOrEmpty($stockNo)) {
                                                    $key = $stockNo
                                                }
                                            }
                                            
                                            # If no stock number, try PART NO.
                                            if (-not $key -and $partNoCol) {
                                                $partNo = $worksheet.Cells.Item($row, $partNoCol).Text
                                                if (-not [string]::IsNullOrEmpty($partNo)) {
                                                    $key = $partNo
                                                }
                                            }
                                            
                                            # If we have a key and it's in our dictionary
                                            if ($key -and $sectionDict.ContainsKey($key)) {
                                                $part = $sectionDict[$key]
                                                
                                                # Update QTY
                                                $worksheet.Cells.Item($row, $qtyCol).Value2 = $part.QTY
                                                
                                                # Update Location
                                                $worksheet.Cells.Item($row, $locationCol).Value2 = $part.Location
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Log "Error updating worksheet ${sectionName}: $($_.Exception.Message)"
                    }
                }
                
                # Save the workbook
                $progressBar.Value = [Math]::Min([int]($baseProgress + $progressWeight), 100)
                $progressLabel.Text = "Saving Excel workbook..."
                $progressForm.Refresh()
                $workbook.Save()
                
                # Close the workbook
                $workbook.Close($false)
                $excel.Quit()
            } catch {
                Write-Log "Error updating Excel workbook: $($_.Exception.Message)"
            } finally {
                if ($excel) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
        
        Write-Log "Completed updating book $($book.Name) - Updated $bookUpdatedParts parts"
    }
    
    # Close the progress form
    $progressForm.Close()
    
    $message = "Parts Books update completed:`n"
    $message += "- Updated $totalUpdatedParts parts`n"
    $message += "- Updated $totalUpdatedCSVs section CSV files`n"
    $message += "- Processed $totalBooks books"
    
    Write-Log $message
    [System.Windows.Forms.MessageBox]::Show($message, "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    return $true
}

function Update-PartsRoom {
    Write-Log "Starting Parts Room update process..."
    
    # Get the configuration
    if (-not $config) {
        Write-Log "Error: Configuration not available"
        [System.Windows.Forms.MessageBox]::Show("Configuration not available. Cannot update Parts Room.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Check if a previously selected site exists in the config
    $selectedSiteName = $null
    $selectedSiteId = $null
    
    # Check parts room directory for existing CSV files
    $existingCsvFile = Get-ChildItem -Path $config.PartsRoomDirectory -Filter "*.csv" | Select-Object -First 1
    if ($existingCsvFile) {
        $selectedSiteName = [System.IO.Path]::GetFileNameWithoutExtension($existingCsvFile.Name)
        Write-Log "Found existing Parts Room data for site: $selectedSiteName"
        
        # Get the site ID from Sites.csv
        $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
        if (Test-Path $sitesPath) {
            $sites = Import-Csv -Path $sitesPath
            $siteIdColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
            $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
            
            $matchingSite = $sites | Where-Object { $_.$fullNameColumn -eq $selectedSiteName }
            if ($matchingSite) {
                $selectedSiteId = $matchingSite.$siteIdColumn
                Write-Log "Found Site ID for ${selectedSiteName}: ${selectedSiteId}"
            }
        }
    }
    
    # Ask user if they want to update the existing site or select a new one
    $dialogResult = [System.Windows.Forms.DialogResult]::Yes
    if ($selectedSiteName) {
        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to update the existing Parts Room data for $selectedSiteName?`n`nClick 'Yes' to update the existing site.`nClick 'No' to select a different site.",
            "Update Parts Room",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)
    }
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Log "Parts Room update cancelled by user"
        return
    }
    
    # If user wants to select a new site, or no existing site was found
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::No -or -not $selectedSiteName) {
        # Load the sites from the CSV
        $sitesPath = Join-Path $config.DropdownCsvsDirectory "Sites.csv"
        Write-Log "Loading Sites from CSV at $sitesPath..."
        
        if (-not (Test-Path $sitesPath)) {
            Write-Log "Error: Sites.csv not found at $sitesPath"
            [System.Windows.Forms.MessageBox]::Show("Sites CSV file not found. Cannot update Parts Room.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $sites = Import-Csv -Path $sitesPath
        
        # Check if the CSV was loaded successfully
        if ($sites.Count -eq 0) {
            Write-Log "Error: No data found in the Sites.csv file"
            [System.Windows.Forms.MessageBox]::Show("No data found in the Sites.csv file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Determine the correct column names
        $siteIdColumn = if ($sites[0].PSObject.Properties.Name -contains "Site ID") { "Site ID" } else { $sites[0].PSObject.Properties.Name[0] }
        $fullNameColumn = if ($sites[0].PSObject.Properties.Name -contains "Full Name") { "Full Name" } else { $sites[0].PSObject.Properties.Name[1] }
        Write-Log "Site ID Column: $siteIdColumn, Full Name Column: $fullNameColumn"
        
        # Create the form for site selection
        Write-Log "Creating form for site selection..."
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
        
        Write-Log "Dropdown populated with site names."
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
        Write-Log "Showing form for site selection..."
        $form.ShowDialog()
        $selectedSiteName = $form.Tag
        
        if (-not $selectedSiteName) {
            Write-Log "No site selected, cancelling operation."
            return
        }
        
        Write-Log "Selected Site: $selectedSiteName"
        $selectedRow = $sites | Where-Object { $_.$fullNameColumn -eq $selectedSiteName }
        $selectedSiteId = $selectedRow.$siteIdColumn
    }
    
    # Now we have a site name and ID, proceed with downloading and processing the data
    if (-not $selectedSiteId) {
        Write-Log "Error: Cannot determine site ID for $selectedSiteName"
        [System.Windows.Forms.MessageBox]::Show("Cannot determine site ID for $selectedSiteName. Update cancelled.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Construct the URL for the selected site
    $url = "http://emarssu5.eng.usps.gov/pemarsnp/nm_national_stock.stockroom_by_site?p_site_id=$selectedSiteId&p_search_type=DESC&p_search_string=&p_boh_radio=-1"
    Write-Log "URL for site ${selectedSiteName}: $url"
    
    # Download HTML content
    Write-Log "Downloading HTML content for $selectedSiteName..."
    try {
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Updating Parts Room"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = 'CenterScreen'
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
        $progressLabel.Size = New-Object System.Drawing.Size(370, 20)
        $progressLabel.Text = "Downloading data for $selectedSiteName..."
        $progressForm.Controls.Add($progressLabel)
        
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Location = New-Object System.Drawing.Point(10, 50)
        $progressBar.Size = New-Object System.Drawing.Size(370, 20)
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressForm.Controls.Add($progressBar)
        
        # Show the progress form
        $progressForm.Show()
        $progressForm.Refresh()
        
        # Download the HTML content
        $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing
        $htmlFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.html"
        Write-Log "Saving HTML content to $htmlFilePath..."
        Set-Content -Path $htmlFilePath -Value $htmlContent.Content -Encoding UTF8
        
        # Update the progress label
        $progressLabel.Text = "Processing HTML data..."
        $progressForm.Refresh()
        
        # Process the HTML content
        Write-Log "Processing the downloaded HTML file for $selectedSiteName using DOM..."
        
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
            
            Write-Log "HTML document loaded into DOM successfully."
            
            # Find the main table with the parts data
            $tables = $htmlDoc.getElementsByTagName("table")
            $mainTable = $null
            
            Write-Log "Found $($tables.length) tables in the document."
            
            # Try multiple methods to find the main table
            foreach ($table in $tables) {
                # Try with className
                if ($table.className -eq "MAIN") {
                    $mainTable = $table
                    Write-Log "Found main table using className property."
                    break
                }
                
                # Try with getAttribute
                try {
                    if ($table.getAttribute("class") -eq "MAIN") {
                        $mainTable = $table
                        Write-Log "Found main table using getAttribute method."
                        break
                    }
                } catch {
                    # Ignore errors with getAttribute
                }
                
                # Check if it's a wide table with borders and multiple columns
                try {
                    if ($table.border -eq "1" -and $table.summary -match "stock") {
                        $mainTable = $table
                        Write-Log "Found main table by border and summary attributes."
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
                            Write-Log "Found table with $($headerRow.cells.length) columns, using as main table."
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
            Write-Log "Number of rows found: $($rows.length)"
            
            # Initialize array to hold parsed data
            $parsedData = @()
            
            # Update the progress bar
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            $progressBar.Maximum = $rows.length
            $progressBar.Value = 0
            
            # Process each row (skip first row which is the header)
            for ($i = 1; $i -lt $rows.length; $i++) {
                $progressBar.Value = $i
                $progressLabel.Text = "Processing row $i of $($rows.length)..."
                $progressForm.Refresh()
                
                try {
                    $row = $rows.item($i)
                    if ($null -eq $row) {
                        Write-Log "Warning: Row $i is null, skipping"
                        continue
                    }
                    
                    # Skip rows that don't have the MAIN class or enough cells
                    $rowClass = try { $row.className } catch { "" }
                    if ($rowClass -ne "MAIN" -and $rowClass -ne "HILITE") {
                        Write-Log "Skipping row $i - not a main data row (class: $rowClass)"
                        continue
                    }
                    
                    # Get all cells in the row
                    $cells = $row.getElementsByTagName("td")
                    
                    if ($null -eq $cells -or $cells.length -lt 6) {
                        Write-Log "Skipping row $i - insufficient cells (found: $(if ($null -eq $cells) { "null" } else { $cells.length }))"
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
                                        Write-Log "Warning: Error processing OEM div $j in row $i : $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Log "Warning: Error processing OEM cell in row $i : $($_.Exception.Message)"
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
                    
                    Write-Log "Added row ${i}: Part(NSN)=$partNSN, Description=$description, QTY=$qty, Location=$location"
                } catch {
                    Write-Log "Error processing row $i : $($_.Exception.Message)"
                    # Continue with next row instead of stopping
                }
            }
            
            Write-Log "Number of parsed data entries: $($parsedData.Count)"
            
            if ($parsedData.Count -eq 0) {
                throw "No data parsed from HTML content"
            }
            
            # Check if there's an existing CSV file with parts book data that we need to preserve
            $tempCsvPath = Join-Path $config.PartsRoomDirectory "temp_$selectedSiteName.csv"
            $csvFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.csv"
            $existingData = $null
            
            if (Test-Path $csvFilePath) {
                Write-Log "Found existing CSV file. Will merge with new data..."
                $progressLabel.Text = "Merging with existing data..."
                $progressForm.Refresh()
                
                # Load the existing data
                $existingData = Import-Csv -Path $csvFilePath
                
                # Identify column names from book references (any column not in the base set)
                $baseColumns = @("Part (NSN)", "Description", "QTY", "13 Period Usage", "Location", "OEM 1", "OEM 2", "OEM 3", "Changed Part (NSN)")
                $bookColumns = $existingData[0].PSObject.Properties.Name | Where-Object { $baseColumns -notcontains $_ }
                Write-Log "Found book reference columns: $($bookColumns -join ', ')"
                
                # First, export the new data to a temporary file
                $parsedData | Export-Csv -Path $tempCsvPath -NoTypeInformation
                
                # Now read it back to ensure formatting is consistent
                $newData = Import-Csv -Path $tempCsvPath
                
                # Add any missing columns from the existing data to the new data
                foreach ($column in $bookColumns) {
                    if ($newData[0].PSObject.Properties.Name -notcontains $column) {
                        $newData | ForEach-Object { $_ | Add-Member -NotePropertyName $column -NotePropertyValue "" }
                    }
                }
                
                # Add 'Changed Part (NSN)' column if it doesn't exist
                if ($newData[0].PSObject.Properties.Name -notcontains 'Changed Part (NSN)') {
                    $newData | ForEach-Object { $_ | Add-Member -NotePropertyName 'Changed Part (NSN)' -NotePropertyValue "" }
                }
                
                # Update progress for the merge operation
                $progressBar.Value = 0
                $progressBar.Maximum = $newData.Count
                
                # Create a dictionary for faster lookups of existing data
                $existingDict = @{}
                foreach ($item in $existingData) {
                    if (-not [string]::IsNullOrEmpty($item.'Part (NSN)')) {
                        $existingDict[$item.'Part (NSN)'] = $item
                    }
                }
                
                # Counters for stats
                $updatedCount = 0
                $newCount = 0
                
                # Process each item in the new data
                for ($i = 0; $i -lt $newData.Count; $i++) {
                    $progressBar.Value = $i
                    if ($i % 10 -eq 0) {  # Update the label less frequently for performance
                        $progressLabel.Text = "Merging item $i of $($newData.Count)..."
                        $progressForm.Refresh()
                    }
                    
                    $item = $newData[$i]
                    $partNSN = $item.'Part (NSN)'
                    
                    # Skip items with no Part (NSN)
                    if ([string]::IsNullOrEmpty($partNSN)) {
                        continue
                    }
                    
                    # Check if this part exists in the existing data
                    if ($existingDict.ContainsKey($partNSN)) {
                        # Update QTY and Location from new data
                        $existingItem = $existingDict[$partNSN]
                        
                        # Check if QTY or Location has changed
                        if ($existingItem.QTY -ne $item.QTY -or $existingItem.Location -ne $item.Location) {
                            $existingItem.QTY = $item.QTY
                            $existingItem.Location = $item.Location
                            $existingItem.'13 Period Usage' = $item.'13 Period Usage'
                            $updatedCount++
                            Write-Log "Updated part ${partNSN}: QTY=$($item.QTY), Location=$($item.Location)"
                        }
                        
                        # Update OEM information if it's more complete in the new data
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 1') -and -not [string]::IsNullOrEmpty($item.'OEM 1')) {
                            $existingItem.'OEM 1' = $item.'OEM 1'
                        }
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 2') -and -not [string]::IsNullOrEmpty($item.'OEM 2')) {
                            $existingItem.'OEM 2' = $item.'OEM 2'
                        }
                        if ([string]::IsNullOrEmpty($existingItem.'OEM 3') -and -not [string]::IsNullOrEmpty($item.'OEM 3')) {
                            $existingItem.'OEM 3' = $item.'OEM 3'
                        }
                    } else {
                        # This is a new part, add it to the existing data with empty book references
                        foreach ($column in $bookColumns) {
                            if ($item.PSObject.Properties.Name -notcontains $column) {
                                $item | Add-Member -NotePropertyName $column -NotePropertyValue ""
                            }
                        }
                        
                        # Add to the dictionary for future reference
                        $existingDict[$partNSN] = $item
                        
                        # Add to the existing data list
                        $existingData += $item
                        $newCount++
                        Write-Log "Added new part: $partNSN"
                    }
                }
                
                # Now check for parts that exist in the existing data but not in the new data
                # They might have been removed from inventory
                $removedParts = @()
                foreach ($existingItem in $existingData) {
                    $partNSN = $existingItem.'Part (NSN)'
                    if (-not [string]::IsNullOrEmpty($partNSN)) {
                        $found = $false
                        foreach ($newItem in $newData) {
                            if ($newItem.'Part (NSN)' -eq $partNSN) {
                                $found = $true
                                break
                            }
                        }
                        
                        if (-not $found) {
                            # Mark this part as potentially removed from inventory
                            $existingItem.QTY = "0"
                            $existingItem.Location = "Not in current inventory"
                            $removedParts += $partNSN
                        }
                    }
                }
                
                if ($removedParts.Count -gt 0) {
                    Write-Log "Marked $($removedParts.Count) parts as not in current inventory"
                }
                
                # Export the updated data back to the main CSV file
                $progressLabel.Text = "Saving updated data..."
                $progressForm.Refresh()
                
                $existingData | Export-Csv -Path $csvFilePath -NoTypeInformation
                
                # Clean up the temporary file
                if (Test-Path $tempCsvPath) {
                    Remove-Item -Path $tempCsvPath -Force
                }
                
                Write-Log "Merged data successfully: Updated $updatedCount parts, added $newCount new parts, marked $($removedParts.Count) parts as not in current inventory"
                [System.Windows.Forms.MessageBox]::Show("Parts Room data updated successfully:`n- Updated $updatedCount existing parts`n- Added $newCount new parts`n- Marked $($removedParts.Count) parts as not in current inventory", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                # No existing file, just save the new data
                $progressLabel.Text = "Saving data to CSV..."
                $progressForm.Refresh()
                
                # Export parsed data to CSV
                $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-Log "Created new CSV file at $csvFilePath with $($parsedData.Count) parts"
                [System.Windows.Forms.MessageBox]::Show("Parts Room data has been created successfully with $($parsedData.Count) parts.", "Update Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            
            # Now check if we need to update the Excel file
            $excelFilePath = Join-Path $config.PartsRoomDirectory "$selectedSiteName.xlsx"
            if (Test-Path $excelFilePath) {
                $updateExcel = [System.Windows.Forms.MessageBox]::Show(
                    "Do you want to update the Excel file with the new data?`n`nNote: This will create a new Excel file. Your existing links to figures in parts books will be preserved in the CSV file.",
                    "Update Excel",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($updateExcel -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $progressLabel.Text = "Updating Excel file..."
                    $progressForm.Refresh()
                    Update-ExcelFile -siteName $selectedSiteName -csvPath $csvFilePath -excelPath $excelFilePath
                }
            } else {
                $createExcel = [System.Windows.Forms.MessageBox]::Show(
                    "Do you want to create an Excel file with the new data?",
                    "Create Excel",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($createExcel -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $progressLabel.Text = "Creating Excel file..."
                    $progressForm.Refresh()
                    Update-ExcelFile -siteName $selectedSiteName -csvPath $csvFilePath -excelPath $excelFilePath
                }
            }
            
            # Ask if the user wants to update parts books with the latest QTY and Location data
            $updatePartsBooks = [System.Windows.Forms.MessageBox]::Show(
                "Do you want to update Parts Books with the latest quantity and location information?",
                "Update Parts Books",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
                
            if ($updatePartsBooks -eq [System.Windows.Forms.DialogResult]::Yes) {
                $progressLabel.Text = "Updating Parts Books..."
                $progressForm.Refresh()
                Update-PartsBooks -sourceCSVPath $csvFilePath
            }
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            Write-Log "Stack Trace: $($_.ScriptStackTrace)"
            $errorMessage = "Error: $($_.Exception.Message)`r`nStack Trace: $($_.ScriptStackTrace)"
            $errorMessage | Out-File -FilePath $logPath -Append
            [System.Windows.Forms.MessageBox]::Show("An error occurred. Please check the error log at $logPath for details.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            # Clean up COM objects
            if ($null -ne $htmlDoc) {
                try {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
                } catch {
                    Write-Log "Warning: Failed to release COM object: $($_.Exception.Message)"
                }
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            # Close the progress form
            $progressForm.Close()
        }
    } catch {
        Write-Log "Error downloading HTML content: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to download HTML content: $($_.Exception.Message)", "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

################################################################################
#                           Needs Category                                     #
################################################################################

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

################################################################################
#                           Main Execution                                     #
################################################################################

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

        $form.Add_FormClosing({
            Save-LaborLogs -listView $script:listViewLaborLog -filePath $laborLogsFilePath
        })
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


# Main execution
Show-MainForm
Write-Log "Application closed"