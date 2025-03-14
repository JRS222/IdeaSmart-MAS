# Test script to parse HTML file into CSV using DOM approach
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define paths
$htmlFilePath = "C:\Users\JR\Documents\Work\PartsBookManagerRootDirectory\Parts Room\BROCKTON PDC, MA, 02301-9731.html"
$outputDirectory = Split-Path -Parent $htmlFilePath
$siteName = [System.IO.Path]::GetFileNameWithoutExtension($htmlFilePath)
$csvFilePath = Join-Path $outputDirectory "$siteName.csv"
$logPath = Join-Path $outputDirectory "parsing_error_log.txt"

Write-Host "Testing HTML to CSV conversion using DOM parsing" -ForegroundColor Cyan
Write-Host "HTML File: $htmlFilePath" -ForegroundColor Cyan
Write-Host "Output CSV: $csvFilePath" -ForegroundColor Cyan

try {
    # Verify the HTML file exists
    if (-not (Test-Path $htmlFilePath)) {
        throw "HTML file not found at $htmlFilePath"
    }

    # Read the HTML content
    Write-Host "Reading HTML file..." -ForegroundColor Yellow
    $htmlContent = Get-Content -Path $htmlFilePath -Raw -Encoding UTF8
    
    if ([string]::IsNullOrWhiteSpace($htmlContent)) {
        throw "HTML content is empty or null"
    }

    Write-Host "HTML file read successfully. Creating HTML Document object..." -ForegroundColor Yellow
    
    # Create HTML Document object
    $htmlDoc = New-Object -ComObject "HTMLFile"
    
    # Try to load HTML content
    try {
        # For newer PowerShell versions
        $htmlDoc.IHTMLDocument2_write($htmlContent)
    } catch {
        # Fallback method for older PowerShell
        Write-Host "Using fallback method to load HTML..." -ForegroundColor Yellow
        $src = [System.Text.Encoding]::Unicode.GetBytes($htmlContent)
        $htmlDoc.write($src)
    }

    Write-Host "HTML content loaded into DOM. Searching for main table..." -ForegroundColor Yellow
    
    # Find the main table
    $tables = $htmlDoc.getElementsByTagName("table")
    $mainTable = $null
    
    Write-Host "Found $($tables.length) tables in the HTML document" -ForegroundColor Yellow
    
    # First try with className
    foreach ($table in $tables) {
        if ($table.className -eq "MAIN") {
            $mainTable = $table
            Write-Host "Found main table using className property" -ForegroundColor Green
            break
        }
    }
    
    # Try with getAttribute if className didn't work
    if ($mainTable -eq $null) {
        Write-Host "Trying alternative table search method..." -ForegroundColor Yellow
        foreach ($table in $tables) {
            if ($table.getAttribute("class") -eq "MAIN") {
                $mainTable = $table
                Write-Host "Found main table using getAttribute method" -ForegroundColor Green
                break
            }
        }
    }
    
    # Last resort: check for a table with 'border' attribute and specific structure
    if ($mainTable -eq $null) {
        Write-Host "Trying to find table by attributes..." -ForegroundColor Yellow
        foreach ($table in $tables) {
            if ($table.border -eq "1" -and $table.summary -match "stock") {
                $mainTable = $table
                Write-Host "Found main table by border and summary attributes" -ForegroundColor Green
                break
            }
        }
    }
    
    if ($mainTable -eq $null) {
        throw "Could not find the main table in the HTML content"
    }
    
    # Get the table rows
    $rows = $mainTable.getElementsByTagName("tr")
    Write-Host "Found $($rows.length) rows in the main table" -ForegroundColor Yellow
    
    # Initialize array for parsed data
    $parsedData = @()
    $rowCount = 0
    $parsedCount = 0
    
    # Process each row (skip the header row)
    for ($i = 1; $i -lt $rows.length; $i++) {
        $row = $rows.item($i)
        $rowCount++
        
        # Get all cells in the row
        $cells = $row.getElementsByTagName("td")
        
        # Only process rows with enough cells
        if ($cells.length -ge 6) {
            try {
                # Extract basic part information
                $partNSN = $cells.item(0).innerText.Trim()
                $description = $cells.item(1).innerText.Trim()
                $qty = [int]($cells.item(2).innerText -replace '[^\d]', '')
                $usage = [int]($cells.item(3).innerText -replace '[^\d]', '')
                $location = $cells.item(5).innerText.Trim()
                
                # Extract OEM information
                $oemCell = $cells.item(4)
                $oemDivs = $oemCell.getElementsByTagName("div")
                
                $oem1 = ""
                $oem2 = ""
                $oem3 = ""
                
                # Process each div for OEM info
                for ($j = 0; $j -lt $oemDivs.length; $j++) {
                    $oemDiv = $oemDivs.item($j)
                    $oemText = $oemDiv.innerText.Trim()
                    
                    if ($oemText -match "OEM:1\s+(.+)") {
                        $oem1 = $matches[1]
                    } elseif ($oemText -match "OEM:2\s+(.+)") {
                        $oem2 = $matches[1]
                    } elseif ($oemText -match "OEM:3\s+(.+)") {
                        $oem3 = $matches[1]
                    }
                }
                
                # Create object with part data
                $partObject = [PSCustomObject]@{
                    "Part (NSN)" = $partNSN
                    "Description" = $description
                    "QTY" = $qty
                    "13 Period Usage" = $usage
                    "Location" = $location
                    "OEM 1" = $oem1
                    "OEM 2" = $oem2
                    "OEM 3" = $oem3
                }
                
                $parsedData += $partObject
                $parsedCount++
                
                # Print some samples for verification
                if ($parsedCount -le 3 -or $parsedCount % 50 -eq 0) {
                    Write-Host "Parsed part: $partNSN - $description" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "Error processing row $i : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Skipping row $i - insufficient cells (found $($cells.length))" -ForegroundColor Yellow
        }
    }
    
    Write-Host "Parsing complete. Total rows processed: $rowCount, Parts extracted: $($parsedData.Count)" -ForegroundColor Cyan
    
    if ($parsedData.Count -eq 0) {
        throw "No data was parsed from the HTML"
    }
    
    # Export data to CSV
    Write-Host "Exporting data to CSV file: $csvFilePath" -ForegroundColor Yellow
    $parsedData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    
    if (Test-Path $csvFilePath) {
        $fileInfo = Get-Item $csvFilePath
        Write-Host "CSV file created successfully: $csvFilePath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Green
        Write-Host "First few rows of CSV:" -ForegroundColor Cyan
        
        # Display first few rows as preview
        $previewRows = Import-Csv -Path $csvFilePath | Select-Object -First 3
        $previewRows | Format-Table -AutoSize
        
        [System.Windows.Forms.MessageBox]::Show("CSV file has been created successfully at:`n$csvFilePath", 
            "CSV Creation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        throw "Failed to create CSV file"
    }
    
    # Clean up COM objects
    if ($htmlDoc -ne $null) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
} catch {
    $errorMessage = "Error: $($_.Exception.Message)`r`nStack Trace: $($_.ScriptStackTrace)"
    Write-Host $errorMessage -ForegroundColor Red
    $errorMessage | Out-File -FilePath $logPath -Append
    
    [System.Windows.Forms.MessageBox]::Show("An error occurred during processing.`n`n$($_.Exception.Message)`n`nPlease check the error log at:`n$logPath", 
        "Error", [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

Write-Host "Script execution complete. Press any key to exit..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")