function Get-WorksheetPmRows {
    # PM rows of the Weekly Worksheet for a date: RowType A/B with a
    # numeric equipment number.  Returns an array of row objects
    # (preserving the array even when empty).
    param([DateTime]$Date)

    $ws = Get-WeeklyWorksheet -Date $Date
    if ($null -eq $ws) { return ,@() }

    $rows = @()
    foreach ($row in @($ws.Rows)) {
        if ("$($row.RowType)" -notin @('A', 'B')) { continue }
        $mno = "$($row.Equipment)".Trim()
        $mnoInt = 0
        if (-not [int]::TryParse($mno, [ref]$mnoInt)) { continue }
        $rows += $row
    }
    return ,$rows
}

function Show-PmFetchEmarsDialog {
    # Worksheet-driven picker.  Reads the Weekly Worksheet for the PM
    # tab's selected date, matches its PM rows (RowType A/B) against
    # the eMARS machine list, and fetches those checklists (read-only).
    # The checklist number comes from the worksheet's "Reg-WO No"
    # column (stored as WorkOrderNo); when that number does not
    # resolve, the machine's open checklist is used instead.
    # Returns the Import-PmChecklistFromEmars summary (single fetch),
    # an ARRAY of summaries (bulk fetch), or $null on cancel.
    if ([string]::IsNullOrWhiteSpace((Get-EmarsSetting -Name 'EmarsSite'))) {
        [System.Windows.Forms.MessageBox]::Show(
            "eMARS Site ID is not configured.`r`n`r`nSet it under Settings -> Work Tracking -> eMARS (PM Checklist Fetch).",
            "Fetch from eMARS", "OK", "Information")
        return $null
    }

    # Only machines listed on the Weekly Worksheet for the date the
    # PM Checklist tab is currently showing are gathered.
    $filterDate = (Get-Date).Date
    if ($script:pmDatePicker) { $filterDate = $script:pmDatePicker.Value.Date }
    $filterDateStr = $filterDate.ToString('yyyy-MM-dd')

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Fetch PM Checklists from eMARS (Weekly Worksheet)"
    $form.Size = New-Object System.Drawing.Size(880, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 460)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'Sizable'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- top bar -------------------------------------------------------
    $topBar = New-Object System.Windows.Forms.Panel
    $topBar.Dock = 'Top'
    $topBar.Height = 104
    $topBar.BackColor = [System.Drawing.Color]::White
    $topBar.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 0)
    $form.Controls.Add($topBar)

    $loadBtn = New-Object System.Windows.Forms.Button
    $loadBtn.Text = "Load Worksheet PMs"
    $loadBtn.Location = New-Object System.Drawing.Point(12, 10)
    $loadBtn.Size = New-Object System.Drawing.Size(180, 30)
    $loadBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $loadBtn.FlatAppearance.BorderSize = 0
    $loadBtn.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $loadBtn.ForeColor = [System.Drawing.Color]::White
    $loadBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $loadBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $topBar.Controls.Add($loadBtn)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "PM rows of the Weekly Worksheet for $filterDateStr, matched against eMARS. Fetch Selected fetches one machine; Fetch All Listed gathers every machine below."
    $hint.Location = New-Object System.Drawing.Point(12, 46)
    $hint.Size = New-Object System.Drawing.Size(840, 22)
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(120,130,145)
    $topBar.Controls.Add($hint)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = ""
    $status.Location = New-Object System.Drawing.Point(12, 70)
    $status.Size = New-Object System.Drawing.Size(840, 24)
    $status.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $status.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $topBar.Controls.Add($status)

    # --- worksheet PM machine list ---------------------------------------
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    Set-ListViewStyle -ListView $lv
    $lv.Columns.Add("Equip", 60)      | Out-Null
    $lv.Columns.Add("Acronym", 90)    | Out-Null
    $lv.Columns.Add("Class", 60)      | Out-Null
    $lv.Columns.Add("Reg-WO No", 120) | Out-Null
    $lv.Columns.Add("Machine ID", 110)| Out-Null
    $lv.Columns.Add("Site", 300)      | Out-Null
    $lv.FullRowSelect = $true
    $form.Controls.Add($lv)

    # --- bottom bar ------------------------------------------------------
    $btnBar = New-Object System.Windows.Forms.Panel
    $btnBar.Dock = 'Bottom'
    $btnBar.Height = 52
    $form.Controls.Add($btnBar)

    $fetchBtn = New-Object System.Windows.Forms.Button
    $fetchBtn.Text = "Fetch Selected"
    $fetchBtn.Size = New-Object System.Drawing.Size(140, 32)
    $fetchBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $fetchBtn.FlatAppearance.BorderSize = 0
    $fetchBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $fetchBtn.ForeColor = [System.Drawing.Color]::White
    $fetchBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $fetchBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $fetchBtn.Location = New-Object System.Drawing.Point(12, 10)
    $fetchBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $fetchBtn.Enabled = $false
    $btnBar.Controls.Add($fetchBtn)

    $fetchAllBtn = New-Object System.Windows.Forms.Button
    $fetchAllBtn.Text = "Fetch All Listed"
    $fetchAllBtn.Size = New-Object System.Drawing.Size(160, 32)
    $fetchAllBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $fetchAllBtn.FlatAppearance.BorderSize = 0
    $fetchAllBtn.BackColor = [System.Drawing.Color]::FromArgb(39,174,96)
    $fetchAllBtn.ForeColor = [System.Drawing.Color]::White
    $fetchAllBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $fetchAllBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $fetchAllBtn.Location = New-Object System.Drawing.Point(160, 10)
    $fetchAllBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $fetchAllBtn.Enabled = $false
    $btnBar.Controls.Add($fetchAllBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Cancel"
    $cancelBtn.Size = New-Object System.Drawing.Size(90, 32)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.BackColor = [System.Drawing.Color]::FromArgb(189,195,199)
    $cancelBtn.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $cancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $cancelBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelBtn.Location = New-Object System.Drawing.Point(760, 10)
    $cancelBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $cancelBtn.Add_Click({ $form.Tag = $null; $form.Close() })
    $btnBar.Controls.Add($cancelBtn)

    # --- wiring ----------------------------------------------------------
    $doFetchOne = {
        $sel = $lv.SelectedItems
        if ($sel.Count -eq 0) { return }
        $entry = $sel[0].Tag          # @{ Row = <worksheet row>; Machine = <eMARS machine> }
        if ($null -eq $entry) { return }
        $m = $entry.Machine
        $status.Text = "Fetching checklist for $($m.Acronym) $($m.No) ..."
        $form.Refresh()
        try {
            $summary = Import-PmChecklistFromEmars -Machine $m `
                -TrackNo "$($entry.Row.WorkOrderNo)" `
                -WorksheetChecklistNo "$($entry.Row.ChecklistNo)"
            $form.Tag = $summary
            $form.Close()
        } catch {
            Write-Log "Show-PmFetchEmarsDialog: fetch failed — $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Fetch failed:`r`n$($_.Exception.Message)",
                "Fetch from eMARS", "OK", "Error")
            $status.Text = "Fetch failed — see message above."
        }
    }

    $doFetchAll = {
        $entries = @()
        foreach ($item in $lv.Items) { $entries += $item.Tag }
        if ($entries.Count -eq 0) {
            $status.Text = "No machines listed. Click Load Worksheet PMs first."
            return
        }

        $imported = @()
        $skipped  = @()
        $failed   = @()
        foreach ($entry in $entries) {
            $m = $entry.Machine
            $label = "$($m.Acronym) $($m.No)"
            $status.Text = "Fetching checklist for $label ..."
            $form.Refresh()
            try {
                $imported += (Import-PmChecklistFromEmars -Machine $m `
                    -TrackNo "$($entry.Row.WorkOrderNo)" `
                    -WorksheetChecklistNo "$($entry.Row.ChecklistNo)")
            } catch {
                $msg = $_.Exception.Message
                Write-Log "Show-PmFetchEmarsDialog: $label failed — $msg"
                if ($msg -match 'Checklist No\. 000') {
                    $skipped += "$label (not created yet)"
                } elseif ($msg -match 'No open checklist') {
                    $skipped += "$label (none open)"
                } else {
                    $failed += "$label — $msg"
                }
            }
        }

        $lines = @()
        if ($imported.Count -gt 0) { $lines += "Imported: $($imported.Count)" }
        if ($skipped.Count  -gt 0) { $lines += "Skipped: $($skipped -join ', ')" }
        if ($failed.Count   -gt 0) { $lines += "Failed: $($failed -join '; ')" }
        [System.Windows.Forms.MessageBox]::Show(($lines -join "`r`n"), "Fetch All Complete", "OK", "Information")
        $status.Text = ($lines -join '  |  ')

        if ($imported.Count -gt 0) {
            $form.Tag = ,@($imported)
            $form.Close()
        }
    }

    $loadBtn.Add_Click({
        $status.ForeColor = [System.Drawing.Color]::FromArgb(44,62,80)
        $status.Text = "Reading Weekly Worksheet for $filterDateStr ..."
        $form.Refresh()

        $pmRows = Get-WorksheetPmRows -Date $filterDate
        if ($pmRows.Count -eq 0) {
            $status.Text = "No PM rows on the Weekly Worksheet for $filterDateStr.`r`nFetch the worksheet first (Work Tracking -> Weekly Worksheet -> Fetch from eDAC), then load again."
            $status.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
            $lv.Items.Clear()
            $fetchBtn.Enabled    = $false
            $fetchAllBtn.Enabled = $false
            return
        }

        # Machine list per acronym (cached; one GET per acronym)
        $machineCache = @{}
        $listed   = 0
        $unmatched = 0
        $lv.Items.Clear()
        $fetchBtn.Enabled    = $false
        $fetchAllBtn.Enabled = $false

        foreach ($row in $pmRows) {
            $acr = "$($row.Acronym)".Trim().ToUpper()
            if ([string]::IsNullOrWhiteSpace($acr)) { $unmatched++; continue }
            $eq = 0
            if (-not [int]::TryParse("$($row.Equipment)".Trim(), [ref]$eq)) { $unmatched++; continue }

            if (-not $machineCache.ContainsKey($acr)) {
                try {
                    $status.Text = "Loading eMARS machine list for $acr ..."
                    $form.Refresh()
                    $machineCache[$acr] = @(Get-EmarsMachineList -Acronym $acr)
                } catch {
                    Write-Log "Show-PmFetchEmarsDialog: $acr machine list failed — $($_.Exception.Message)"
                    $machineCache[$acr] = @()
                }
            }

            $machine = $null
            foreach ($m in $machineCache[$acr]) {
                $mNo = 0
                if (-not [int]::TryParse("$($m.No)", [ref]$mNo)) { continue }
                if ($mNo -eq $eq) { $machine = $m; break }
            }
            if ($null -eq $machine) { $unmatched++; continue }

            $it = New-Object System.Windows.Forms.ListViewItem("$eq")
            $it.SubItems.Add("$acr")                   | Out-Null
            $it.SubItems.Add("$($machine.ClassCode)")  | Out-Null
            $it.SubItems.Add("$($row.WorkOrderNo)")    | Out-Null
            $it.SubItems.Add("$($machine.Ptr)")        | Out-Null
            $it.SubItems.Add("$($machine.SiteName)")   | Out-Null
            $it.Tag = @{ Row = $row; Machine = $machine }
            $lv.Items.Add($it) | Out-Null
            $listed++
        }

        $fetchAllBtn.Enabled = ($listed -gt 0)

        if ($listed -eq 0) {
            $status.Text = "No worksheet PM machines matched the eMARS machine list ($unmatched unmatched)."
            $status.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
        } else {
            $status.Text = "$listed worksheet PM machine(s) listed"
            if ($unmatched -gt 0) { $status.Text += "  ($unmatched could not be matched to eMARS)" }
            $status.Text += ". Select one and Fetch Selected, or Fetch All Listed."
        }
    })

    $lv.Add_ItemSelectionChanged({
        $fetchBtn.Enabled = ($lv.SelectedItems.Count -gt 0)
    })

    $lv.Add_DoubleClick($doFetchOne)
    $fetchBtn.Add_Click($doFetchOne)
    $fetchAllBtn.Add_Click($doFetchAll)
    $form.AcceptButton = $fetchBtn
    $form.CancelButton = $cancelBtn

    $form.ShowDialog() | Out-Null
    return $form.Tag
}
