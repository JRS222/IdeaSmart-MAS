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

# Test the function
$htmlFilePath = "C:\Users\JR\Documents\Work\Programming Projects\IdeaSmart-MAS\Sample Data\MS-174 Identification Code Sort - ICS (Vol. C)\HTML and CSV Files\Figure 3-5.html"
$htmlContent = Get-Content -Path $htmlFilePath -Raw
Process-HTMLToCSV -htmlContent $htmlContent -htmlFilePath $htmlFilePath

# Open the resulting CSV file
$csvFilePath = [System.IO.Path]::ChangeExtension($htmlFilePath, '.csv')
if (Test-Path $csvFilePath) {
    Invoke-Item $csvFilePath
} else {
    Write-Host "CSV file was not created. Check for errors."
}