Here are all three changes. Since you haven't applied anything yet, this is a complete drop-in set — apply them in order.

---

Change 1 — Excel hyperlink points to the HTML figure, not the CSV

The bug: Merge-PartsBookSections writes the raw CSV filename (Figure 2-1.csv) into the REF. column, so $figName.html became Figure 2-1.csv.html — which doesn't exist, so no hyperlink was ever added.

Find this block in New-PartsBookExcel:

```powershell
            $refIdx = -1
            for ($c = 0; $c -lt $headers.Count; $c++) { if ($headers[$c] -eq 'REF.') { $refIdx = $c + 1; break } }
            if ($refIdx -gt 0 -and $rowCount -gt 0) {
                for ($r = 2; $r -le $rowCount + 1; $r++) {
                    $v = $ws.Cells.Item($r, $refIdx).Value2
                    if ($v -and "$v" -match 'Figure \d+-\d+') {
                        $figName = "$v"
                        $htmlPath = Join-Path $SourceDir "HTML and CSV Files\$figName.html"
                        if (Test-Path $htmlPath) {
                            $cell = $ws.Cells.Item($r, $refIdx)
                            $ws.Hyperlinks.Add($cell, $htmlPath, "", "", $figName) | Out-Null
                        }
                    }
                }
            }
```

Replace with:

```powershell
            $refIdx = -1
            for ($c = 0; $c -lt $headers.Count; $c++) { if ($headers[$c] -eq 'REF.') { $refIdx = $c + 1; break } }
            if ($refIdx -gt 0 -and $rowCount -gt 0) {
                for ($r = 2; $r -le $rowCount + 1; $r++) {
                    $v = $ws.Cells.Item($r, $refIdx).Value2
                    if ($v -and "$v" -match 'Figure \d+-\d+') {
                        # REF. stores the CSV filename (e.g. "Figure 2-1.csv").
                        # Strip the extension so we look for the HTML that
                        # Export-PartsBookFigures wrote alongside it.
                        $figName  = ("$v" -replace '\.csv$', '').Trim()
                        $htmlPath = Join-Path $SourceDir "HTML and CSV Files\$figName.html"
                        if (Test-Path $htmlPath) {
                            # file:/// URI so paths with spaces
                            # ("HTML and CSV Files") survive intact.
                            try {
                                $uri = (New-Object System.Uri -ArgumentList $htmlPath).AbsoluteUri
                            } catch {
                                $uri = $htmlPath
                            }
                            $cell = $ws.Cells.Item($r, $refIdx)
                            $ws.Hyperlinks.Add($cell, $uri, "", "", $figName) | Out-Null
                        }
                    }
                }
            }
```

---

Change 2 — Highlight the source <tr> inside the opened HTML figure

When you check a cross-reference row (e.g., item 4 on Figure 2-1) and click Open Figures, the figure HTML opens with the row for item 4 highlighted yellow. Since the downloaded HTML is static and should stay pristine, we open a patched copy from %TEMP%.

Find the existing $openFiguresButton.Add_Click({ ... }) block in Setup-SearchTab (it starts with $checkedItems = $script:listViewCrossRef.CheckedItems) and replace the whole block with:

```powershell
    $openFiguresButton.Add_Click({
        $checkedItems = $script:listViewCrossRef.CheckedItems
        if ($checkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one item to open figures.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        # Group checked rows by figure so a figure referenced from multiple
        # rows opens once, with every matching "NO." highlighted.
        # CrossRef ListView columns:
        #   [0] Handbook  [1] Section Name  [2] NO.  [3] PART DESCRIPTION
        #   [4] REF.  [5] STOCK NO.  [6] PART NO.  [7] CAGE  [8] Location  [9] QTY
        $byFigure = @{}
        foreach ($item in $checkedItems) {
            $handbook = $item.SubItems[0].Text
            $no       = $item.SubItems[2].Text     # NO. — identifies the row inside the figure
            $ref      = $item.SubItems[4].Text     # REF. — ".csv" already stripped by search

            if ([string]::IsNullOrWhiteSpace($handbook) -or [string]::IsNullOrWhiteSpace($ref)) { continue }

            $key = "$handbook|$ref"
            if (-not $byFigure.ContainsKey($key)) {
                $byFigure[$key] = [PSCustomObject]@{
                    Handbook = $handbook
                    Figure   = $ref
                    Items    = New-Object System.Collections.ArrayList
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($no)) {
                [void]$byFigure[$key].Items.Add($no)
            }
        }

        if ($byFigure.Count -eq 0) { return }

        # Temp files live outside the data tree so originals stay untouched.
        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) 'PartsMgmt_FigViews'
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        foreach ($entry in $byFigure.Values) {
            $bookProp = $config.Books.PSObject.Properties[$entry.Handbook]
            if ($bookProp -and $bookProp.Value.VolumesToUrlCsvPath) {
                $bookDir = Split-Path -Path $bookProp.Value.VolumesToUrlCsvPath -Parent
            } else {
                $bookDir = Join-Path $config.PartsBooksDirectory $entry.Handbook
            }
            $htmlFilePath = Join-Path $bookDir "HTML and CSV Files\$($entry.Figure).html"

            if (-not (Test-Path $htmlFilePath)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Figure file not found:`r`n$htmlFilePath",
                    "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Log "Figure file not found: $htmlFilePath"
                continue
            }

            try {
                $raw = [System.IO.File]::ReadAllText($htmlFilePath)

                # JSON array of item numbers.  Cheap quote-escaping keeps
                # this safe even if NO. ever contains quotes.
                $itemJson = (@($entry.Items) | ForEach-Object {
                    '"' + ([string]$_).Replace('"', '\"') + '"'
                }) -join ','

                $inject = @"
<style>
tr.pm-highlight-row, tr.pm-highlight-row td {
    background-color: #ffff99 !important;
}
</style>
<script>
(function () {
    var targets = [$itemJson];
    if (!targets.length) { return; }
    var norm = function (s) {
        return (s || '').replace(/\u00a0/g, ' ').replace(/^\s+|\s+$/g, '').toUpperCase();
    };
    var wanted = targets.map(norm);
    var rows = document.querySelectorAll('tr');
    for (var i = 0; i < rows.length; i++) {
        var cell = rows[i].cells && rows[i].cells[0];
        if (!cell) { continue; }
        var txt = norm(cell.innerText || cell.textContent || '');
        for (var j = 0; j < wanted.length; j++) {
            if (txt === wanted[j]) {
                rows[i].className = (rows[i].className || '') + ' pm-highlight-row';
                break;
            }
        }
    }
})();
</script>
"@

                if ($raw -match '(?i)</body>') {
                    $patched = [regex]::Replace($raw, '(?i)</body>', "$inject`r`n</body>", 1)
                } else {
                    $patched = $raw + "`r`n" + $inject
                }

                $safeFig = ($entry.Figure -replace '[^\w\-]', '_')
                $tmpName = "$safeFig-" + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html'
                $tmpPath = Join-Path $tempDir $tmpName
                [System.IO.File]::WriteAllText(
                    $tmpPath, $patched, (New-Object System.Text.UTF8Encoding($false)))

                Start-Process $tmpPath
                Write-Log "Opened figure with highlight: $htmlFilePath (items: $($entry.Items -join ','))"
            } catch {
                Write-Log "Failed to open figure with highlight: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to open figure:`r`n$($_.Exception.Message)",
                    "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
```

---

Change 3 — Repair existing Excel files when Update Files runs

Two edits: (a) add a new function, (b) call it from Update-PartsBooks so both Update Files and Update Parts Books trigger the repair on every book.

3a — Add the repair function

Insert this new function between New-PartsBookExcel and Rename-PartsBookWorksheets (search for function Rename-PartsBookWorksheets and paste immediately before it):

```powershell
function Repair-PartsBookFigureHyperlinks {
    # Walks every worksheet in the workbook, finds the REF. column, and
    # normalizes any figure hyperlink to point at the real ".html" file
    # sitting next to its ".csv".  Handles all four historical states:
    #   (a) bare text "Figure 2-1.csv", no hyperlink      -> add hyperlink
    #   (b) hyperlink -> "Figure 2-1.csv"                 -> retarget to .html
    #   (c) hyperlink -> "Figure 2-1.html", display dirty -> clean display
    #   (d) hyperlink -> "Figure 2-1.html", display clean -> skip
    # Returns the number of cells repaired.  Idempotent.
    param(
        [string]$ExcelPath,
        [string]$SourceDir
    )

    if (-not (Test-Path $ExcelPath)) { return 0 }
    $htmlCsvDir = Join-Path $SourceDir "HTML and CSV Files"
    if (-not (Test-Path $htmlCsvDir)) { return 0 }

    $repairs = 0
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($ExcelPath)

        foreach ($ws in @($wb.Worksheets)) {
            # Locate the REF. column (case-insensitive, tolerant of trailing space)
            $refCol = 0
            $lastCol = [Math]::Max(20, [int]$ws.UsedRange.Columns.Count)
            for ($c = 1; $c -le $lastCol; $c++) {
                $hdr = "$($ws.Cells.Item(1, $c).Value2)".Trim()
                if ($hdr -ieq 'REF.') { $refCol = $c; break }
            }
            if ($refCol -eq 0) { continue }

            $lastRow = [int]$ws.UsedRange.Rows.Count
            if ($lastRow -lt 2) { continue }

            for ($r = 2; $r -le $lastRow; $r++) {
                if (($r % 100) -eq 0) {
                    try { [System.Windows.Forms.Application]::DoEvents() } catch { }
                }

                $cell = $ws.Cells.Item($r, $refCol)
                $val  = "$($cell.Value2)"
                if ($val -notmatch 'Figure\s+\d+-\d+') { continue }

                # Canonical display text: "Figure 2-1" (no extension)
                $figName = ($val -replace '\.csv$', '' -replace '\.html$', '').Trim()
                $htmlPath = Join-Path $htmlCsvDir "$figName.html"
                if (-not (Test-Path $htmlPath)) { continue }   # nothing to link to

                try {
                    $uri = (New-Object System.Uri -ArgumentList $htmlPath).AbsoluteUri
                } catch {
                    $uri = $htmlPath
                }

                # Inspect the current state of this cell
                $existing = $null
                try {
                    if ($cell.Hyperlinks.Count -gt 0) { $existing = $cell.Hyperlinks.Item(1) }
                } catch { }

                $needsFix = $false
                if ($null -eq $existing) {
                    $needsFix = $true
                } else {
                    $target = "$($existing.Address)"
                    $targetIsHtml = ($target -match '\.html($|\?)')
                    $displayClean = ($val -eq $figName)
                    if (-not $targetIsHtml -or -not $displayClean) { $needsFix = $true }
                }

                if (-not $needsFix) { continue }

                # Delete any existing hyperlink(s) on the cell, then add a
                # clean one.  Excel won't let us mutate the address in place.
                try {
                    if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() }
                } catch { }

                $ws.Hyperlinks.Add($cell, $uri, "", "", $figName) | Out-Null
                $cell.Value2 = $figName        # belt & suspenders
                $repairs++
            }
        }

        if ($repairs -gt 0) { $wb.Save() }
    } catch {
        Write-Log "Repair-PartsBookFigureHyperlinks error: $($_.Exception.Message)"
    } finally {
        if ($wb) {
            try { $wb.Close($false) } catch { }    # we already Save()'d above if needed
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    return $repairs
}
```

3b — Call the repair from Update-PartsBooks

Find this line near the end of the outer for ($bookIndex …) loop in Update-PartsBooks:

```powershell
        Write-Log "Completed updating book $($book.Name) - Updated $bookUpdatedParts parts"
    }
```

Insert the repair pass immediately before that Write-Log line:

```powershell
        # ---- Repair figure hyperlinks in the workbook ----
        # Runs unconditionally (not just when CSV data changed) so that
        # older workbooks with ".csv" references get normalized to the
        # real ".html" files.  Idempotent: already-correct links are
        # left alone.
        if ($book.ExcelPath -and (Test-Path $book.ExcelPath)) {
            $progressLabel.Text = "Repairing figure hyperlinks..."
            $bookLabel.Text = "Processing book: $($book.Name) - link repair"
            $progressForm.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $fixed = Repair-PartsBookFigureHyperlinks `
                    -ExcelPath $book.ExcelPath `
                    -SourceDir $book.Directory
                if ($fixed -gt 0) {
                    Write-Log "Repaired $fixed figure hyperlink(s) in $($book.Name)"
                }
            } catch {
                Write-Log "Hyperlink repair failed for $($book.Name): $($_.Exception.Message)"
            }
        }

        Write-Log "Completed updating book $($book.Name) - Updated $bookUpdatedParts parts"
```

Because Update-AllFiles calls Update-PartsBooks in Phase 3, clicking Update Files will:

1. Refresh every parts-room CSV,
2. Refresh every CombinedSections CSV,
3. Refresh QTY/Location in every workbook's worksheets, and
4. Repair every worksheet's REF. column so cells like Figure 2-1.csv become clickable links to Figure 2-1.html.

The repair is idempotent — running Update Files a second time finds nothing to fix and doesn't touch the workbook.

---

What the user sees after applying all three changes

Flow Behaviour
Build a new parts book → open its .xlsx REF. cells show Figure 2-1 and open the matching .html in the browser.
Existing workbook → click Update Files Every sheet's REF. column is inspected; any .csv reference or missing hyperlink is silently fixed to the .html file. Log records Repaired N figure hyperlink(s) in <book>.
Search → check a cross-ref row → Open Figures The figure HTML opens in the browser and the row whose NO. matches your selection is highlighted yellow. Original files on disk stay unmodified.