function Process-HTMLContent($htmlContent, $selectedRow, $directoryPath) {
    # Initialize logging
    $logPath = Join-Path $directoryPath "dom_processing_debug.log"
    function Write-Log {
        param([string]$Message, [string]$Level = "Info")
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp][$Level] $Message"
        $color = @{Error="Red";Warning="Yellow";DOM="Cyan";Info="White"}[$Level]
        Write-Host $logEntry -ForegroundColor $color
        Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
    }

    Write-Log "Starting HTML processing..." -Level Info
    Write-Log "Input HTML length: $($htmlContent.Length) characters" -Level Info

    # Initialize outputs
    $figureCsvLines = "Figure No.,Name,Section No.,MS Book No,Volume,URL"
    $sectionNames = @{}
    $totalFigures = 0
    $processingErrors = @()

    # DOM PARSING SECTION ======================================================
    Write-Log "Attempting DOM parsing..." -Level DOM
    
    try {
        # Create COM object with proper encoding handling
        $htmlDoc = New-Object -ComObject "HTMLFile"
        $htmlDoc.designMode = "on"  # Critical for proper parsing
        
        # Load HTML content
        try {
            $htmlDoc.IHTMLDocument2_write($htmlContent)
        } catch {
            # Fallback encoding method
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($htmlContent)
            $htmlDoc.write($bytes)
        }

        Write-Log "DOM loaded successfully" -Level DOM

        # DOM Inspection Report
        Write-Log "=== DOM STRUCTURE REPORT ===" -Level DOM
        Write-Log "Body exists: $($null -ne $htmlDoc.body)" -Level DOM
        Write-Log "phbk_tree element found: $($null -ne $htmlDoc.getElementById('phbk_tree'))" -Level DOM

        $treeElement = $htmlDoc.getElementById("phbk_tree")
        if ($treeElement) {
            Write-Log "phbk_tree children: $($treeElement.children.length)" -Level DOM
            Write-Log "phbk_tree innerHTML length: $($treeElement.innerHTML.Length)" -Level DOM
        } else {
            Write-Log "phbk_tree not found by ID, scanning document..." -Level DOM
            $allULs = $htmlDoc.getElementsByTagName("ul")
            Write-Log "Found $($allULs.length) UL elements" -Level DOM
        }

        # Recursive processing function with detailed logging
        function Process-TreeNode {
            param($node, $depth=0)
            
            $indent = '  ' * $depth
            if ($node.tagName) {
                Write-Log "${indent}Processing $($node.tagName)" -Level DOM
                
                # Section detection
                if ($node.tagName -eq 'li') {
                    $sno = $node.getAttribute('sno')
                    $specialChar = $node.getAttribute('special_char')
                    Write-Log "${indent}LI attributes: sno=$sno special_char=$specialChar" -Level DOM

                    if ($sno -and $specialChar) {
                        $span = $node.getElementsByTagName('span') | Select-Object -First 1
                        if ($span) {
                            $sectionTitle = $span.innerText.Trim()
                            if ($sectionTitle -match 'Section\s+(\d+)') {
                                $sectionNum = $matches[1]
                                $sectionNames["Section $sectionNum"] = $sectionTitle
                                Write-Log "${indent}Found section: $sectionTitle" -Level Info
                            }
                        }
                    }
                }

                # Figure detection
                if ($node.tagName -eq 'li' -and $node.hasAttribute('figno')) {
                    $figNo = $node.getAttribute('figno')
                    $span = $node.getElementsByTagName('span') | Where-Object { 
                        $_.className -eq 'go_fig' 
                    } | Select-Object -First 1
                    
                    if ($span) {
                        $figName = $span.innerText.Trim()
                        $sectionNo = ($figNo -split '-')[0]
                        $msBookNo = $selectedRow.'MS Book No'
                        $volume = $selectedRow.Volume
                        $figureUrl = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figNo&viewerflag=d&layout=L11"
                        
                        $figureCsvLines += "`n$figNo,""$figName"",$sectionNo,$msBookNo,$volume,$figureUrl"
                        $totalFigures++
                        Write-Log "${indent}Found figure: $figNo - $figName" -Level Info
                    }
                }

                # Recursive processing
                if ($node.hasChildNodes()) {
                    foreach ($child in $node.childNodes) {
                        Process-TreeNode $child ($depth + 1)
                    }
                }
            }
        }

        # Start processing from phbk_tree or body
        if ($treeElement) {
            Write-Log "Starting tree processing..." -Level DOM
            Process-TreeNode $treeElement
        } else {
            Write-Log "Starting full document processing..." -Level DOM
            Process-TreeNode $htmlDoc.body
        }

    } catch {
        $errMsg = "DOM parsing failed: $_"
        Write-Log $errMsg -Level Error
        $processingErrors += $errMsg
    } finally {
        if ($htmlDoc) {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($htmlDoc) | Out-Null
                Write-Log "Released COM object" -Level DOM
            } catch {
                Write-Log "Error releasing COM object: $_" -Level Warning
            }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # FALLBACK TO REGEX =======================================================
    if ($totalFigures -eq 0 -or $sectionNames.Count -eq 0) {
        Write-Log "Initiating regex fallback..." -Level Warning
        
        # Regex section parsing
        $sectionMatches = [regex]::Matches($htmlContent, '<li[^>]+sno="(\d+)"[^>]+special_char="[^"]*"[^>]*>\s*<span[^>]*>(Section\s+\d+.*?)</span>')
        Write-Log "Found $($sectionMatches.Count) sections via regex" -Level Info
        foreach ($match in $sectionMatches) {
            $sectionTitle = $match.Groups[2].Value.Trim()
            $sectionNum = $match.Groups[1].Value
            $sectionNames["Section $sectionNum"] = $sectionTitle
        }

        # Regex figure parsing
        $figureMatches = [regex]::Matches($htmlContent, '<li[^>]+figno="([^"]+)"[^>]*>\s*<span[^>]+class="go_fig"[^>]*>(.*?)</span>')
        Write-Log "Found $($figureMatches.Count) figures via regex" -Level Info
        foreach ($match in $figureMatches) {
            $figNo = $match.Groups[1].Value.Trim()
            $figName = $match.Groups[2].Value.Trim()
            $sectionNo = ($figNo -split '-')[0]
            $msBookNo = $selectedRow.'MS Book No'
            $volume = $selectedRow.Volume
            $figureUrl = "https://www1.mtsc.usps.gov/apps/phbk/content/printfigandtable.php?msbookno=$msBookNo&volno=$volume&secno=$sectionNo&figno=$figNo&viewerflag=d&layout=L11"
            
            $figureCsvLines += "`n$figNo,""$figName"",$sectionNo,$msBookNo,$volume,$figureUrl"
            $totalFigures++
        }
    }

    # OUTPUT GENERATION =======================================================
    try {
        $volumesToUrlPath = Join-Path $directoryPath "Volumes-to-URL.csv"
        $figureCsvLines | Out-File -FilePath $volumesToUrlPath -Encoding UTF8 -Force
        Write-Log "Generated CSV file with $totalFigures figures" -Level Info

        $sectionNamesPath = Join-Path $directoryPath "SectionNames.txt"
        $sectionNames.Values | Out-File -FilePath $sectionNamesPath -Encoding UTF8 -Force
        Write-Log "Generated section names with $($sectionNames.Count) entries" -Level Info

    } catch {
        $errMsg = "File save error: $_"
        Write-Log $errMsg -Level Error
        $processingErrors += $errMsg
    }

    # FINAL REPORT ============================================================
    Write-Log "Processing complete. Results:`n" +
              "Sections found: $($sectionNames.Count)`n" +
              "Figures found: $totalFigures`n" +
              "Errors encountered: $($processingErrors.Count)" -Level Info

    return @{
        VolumesToUrlPath = $volumesToUrlPath
        SectionNamesPath = $sectionNamesPath
        ProcessingErrors = $processingErrors
        SectionCount = $sectionNames.Count
        FigureCount = $totalFigures
    }
}