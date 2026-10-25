Combined Change Checklist

Grouped into A. Bugs still present from last thread (confirmed in your log), B. New issues. Each item: find-line → old → new.

---

Part A — Bugs still present

A1. Search button Enabled set before it exists

Find: $updateSearchEnabled = { (the first occurrence — inside Setup-SearchTab, around line 14012).

Cause: The block references $searchButton but sits above the line $searchButton = New-Object ... (~line 14040). .GetNewClosure() captures $searchButton = $null, and the immediate $searchButton.Enabled = ... hits the same $null. Both errors at 14015 / 14028 come from here.

Old — delete this whole block from its current position:

```powershell
		$updateSearchEnabled = {
		$searchButton.Enabled = (
			$script:comboBook.SelectedIndex    -ge 0 -and
			$script:comboSection.SelectedIndex -ge 0
		)
	}.GetNewClosure()

	$script:comboBook.Add_SelectedIndexChanged($updateSearchEnabled)
	$script:comboSection.Add_SelectedIndexChanged($updateSearchEnabled)

	# Pre-select first book so the tab is usable as soon as it opens.
	if ($script:comboBook.Items.Count -gt 0) {
		$script:comboBook.SelectedIndex = 0
	}
	$searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -and $script:comboSection.SelectedIndex -ge 0)
```

New — re-add it after $topPanel.Controls.Add($searchButton) (i.e. right after that line):

```powershell
    $topPanel.Controls.Add($searchButton)

    $updateSearchEnabled = {
        $searchButton.Enabled = (
            $script:comboBook.SelectedIndex    -ge 0 -and
            $script:comboSection.SelectedIndex -ge 0
        )
    }.GetNewClosure()

    $script:comboBook.Add_SelectedIndexChanged($updateSearchEnabled)
    $script:comboSection.Add_SelectedIndexChanged($updateSearchEnabled)
```

B2 (below) replaces this enable logic entirely — you can skip A1 if you apply B2, since B2 removes the "must pick something" requirement.

---

A2. Set-ThemeWindowChrome not found inside the closure

Find: $Root.Add_HandleCreated({ inside Apply-UiTheme (around line 910).

Cause: .GetNewClosure() puts the block in a fresh module whose $script: is empty. Function-name resolution then fails when the handle is finally created during ShowDialog().

Old:

```powershell
    if ($Root -is [System.Windows.Forms.Form]) {
        try {
            if (-not $script:themeBase[$Root].ChromeWired) {
                $Root.Add_HandleCreated({
                    Set-ThemeWindowChrome -Form $Root
                }.GetNewClosure())
                $script:themeBase[$Root] | Add-Member -NotePropertyName 'ChromeWired' -NotePropertyValue $true -Force
            }
            Set-ThemeWindowChrome -Form $Root
        } catch { }
    }
```

New:

```powershell
    if ($Root -is [System.Windows.Forms.Form]) {
        try {
            # Capture the function as a CommandInfo local BEFORE the closure.
            # Inside .GetNewClosure() the $script: scope is empty, so a bare
            # call to Set-ThemeWindowChrome would not resolve at fire time.
            $chromeFn = Get-Command -Name Set-ThemeWindowChrome -CommandType Function -ErrorAction SilentlyContinue

            if (-not $script:themeBase[$Root].ChromeWired) {
                $Root.Add_HandleCreated({
                    param($sender, $e)
                    try {
                        if ($null -ne $chromeFn) { & $chromeFn -Form $sender }
                    } catch { }
                }.GetNewClosure())
                $script:themeBase[$Root] | Add-Member -NotePropertyName 'ChromeWired' -NotePropertyValue $true -Force
            }
            Set-ThemeWindowChrome -Form $Root
        } catch { }
    }
```

---

A3. Get-JobById re-reads Jobs.json on every timer tick

Find: function Get-JobById {.

Your log shows ~30 Get-Jobs: read 45 job(s) lines in ~3 seconds — that's the job-timer refresh loop hitting disk each tick.

Old:

```powershell
function Get-JobById {
    param([string]$JobId)
    if ([string]::IsNullOrWhiteSpace($JobId)) { return $null }
    foreach ($j in @(Get-Jobs)) { if ("$($j.id)" -eq $JobId) { return $j } }
    return $null
}
```

New:

```powershell
function Get-JobById {
    param([string]$JobId)
    if ([string]::IsNullOrWhiteSpace($JobId)) { return $null }
    if ($script:wtJobs) {
        foreach ($j in @($script:wtJobs)) { if ("$($j.id)" -eq $JobId) { return $j } }
    }
    foreach ($j in @(Get-Jobs)) { if ("$($j.id)" -eq $JobId) { return $j } }
    return $null
}
```

$script:wtJobs is already maintained by Refresh-JobLogList (set to the freshly-loaded list). The fallback only fires for jobs added since the last refresh — rare in practice.

---

A4. PM Start Timer button calls the wrong function

Find: $startBtn.Add_Click({ inside Show-PmTaskTimerForm.

Cause: The handler you added calls Start-JobTimer -JobId $JobId — but this is the PM task timer form; $JobId and Start-JobTimer are for the Job Log timer. $JobId is undefined here (silent null).

Old:

```powershell
	$startBtn.Add_Click({
		try {
			Start-JobTimer -JobId $JobId
			Refresh-JobLogList
			Refresh-WorkBudget
			& $refreshDisplay
		} catch { Write-Log "Start Timer error: $($_.Exception.Message)" }
	}.GetNewClosure())
```

New:

```powershell
	$startBtn.Add_Click({
		try {
			$curEntry = $Machine.State[$taskId]
			if ((Get-PmTaskTimerState -Entry $curEntry) -eq 'idle') {
				Start-PmTaskTimer -Machine $Machine -Task $Task
				Refresh-PmTaskRow -Machine $Machine -Task $Task
			}
			& $refreshDisplay
		} catch { Write-Log "PM Start Timer error: $($_.Exception.Message)" }
	}.GetNewClosure())
```

$Machine, $Task, $taskId, $refreshDisplay are all already captured correctly by that form. Only the three lines inside try were wrong.

---

Part B — New issues

B1. Sections dropdown stays empty in Part Picker

Find: function Get-SectionsForBook { and function Get-EmarsSetting {.

Root cause: Both handlers that populate sections/refresh the user list are .GetNewClosure() blocks. Inside a closure, $script: refers to an empty closure scope, not the main script scope. So:

· Get-SectionsForBook reads $script:config → $null → returns ,@() → section dropdown never populates.
· Get-EmarsSetting reads $script:config.WorkTracking → $null → Get-EmarsUserList throws "site not set" → caught silently → user list stays empty.

Fix — mirror config into a global that closures can see, and update the two functions to read it.

Step B1a — Set the global. Find:

```powershell
$config = Initialize-Config
```

New (add the second line below it):

```powershell
$config = Initialize-Config
$global:PMConfig = $config
```

Step B1b — Get-SectionsForBook. Change the initial guard and all $script:config references inside the function:

Old:

```powershell
function Get-SectionsForBook {
    param([string]$BookName)

    $out = @()
    if ([string]::IsNullOrWhiteSpace($BookName)) { return ,$out }
    if (-not $script:config -or -not $script:config.Books) { return ,$out }

    $bookProp = $script:config.Books.PSObject.Properties[$BookName]
    if (-not $bookProp) { return ,$out }

    $volumesCsvPath = $bookProp.Value.VolumesToUrlCsvPath
    $bookDir = if ($volumesCsvPath) { Split-Path -Path $volumesCsvPath -Parent }
               else { Join-Path $script:config.PartsBooksDirectory $BookName }

    $combinedDir = Join-Path $bookDir "CombinedSections"
    if (-not (Test-Path $combinedDir)) { return ,$out }
    ...
```

New:

```powershell
function Get-SectionsForBook {
    param([string]$BookName)

    $out = @()
    if ([string]::IsNullOrWhiteSpace($BookName)) { return ,$out }
    # Read from the global so calls inside .GetNewClosure() see it.
    $cfg = if ($global:PMConfig) { $global:PMConfig } else { $script:config }
    if (-not $cfg -or -not $cfg.Books) { return ,$out }

    $bookProp = $cfg.Books.PSObject.Properties[$BookName]
    if (-not $bookProp) { return ,$out }

    $volumesCsvPath = $bookProp.Value.VolumesToUrlCsvPath
    $bookDir = if ($volumesCsvPath) { Split-Path -Path $volumesCsvPath -Parent }
               else { Join-Path $cfg.PartsBooksDirectory $BookName }

    $combinedDir = Join-Path $bookDir "CombinedSections"
    if (-not (Test-Path $combinedDir)) { return ,$out }
    ...
```

Leave the rest of the function body unchanged.

Step B1c — Get-EmarsSetting.

Old:

```powershell
function Get-EmarsSetting {
    # Reads one WorkTracking eMARS setting, '' when missing.
    param([string]$Name)
    $wt = $script:config.WorkTracking
    if ($null -eq $wt) { return '' }
    $p = $wt.PSObject.Properties[$Name]
    if ($null -eq $p) { return '' }
    return "$($p.Value)".Trim()
}
```

New:

```powershell
function Get-EmarsSetting {
    # Reads one WorkTracking eMARS setting, '' when missing.
    param([string]$Name)
    $cfg = if ($global:PMConfig) { $global:PMConfig } else { $script:config }
    if ($null -eq $cfg) { return '' }
    $wt = $cfg.WorkTracking
    if ($null -eq $wt) { return '' }
    $p = $wt.PSObject.Properties[$Name]
    if ($null -eq $p) { return '' }
    return "$($p.Value)".Trim()
}
```

Step B1d — keep $global:PMConfig in sync whenever $script:config is reassigned. Search the file for these lines (there are ~4); add $global:PMConfig = $script:config immediately after each one:

1. In $emarsTestBtn.Add_Click — $script:config = $cfg
2. In $saveBtn.Add_Click (Settings save) — $script:config = $cfg
3. In $reloadBtn.Add_Click (Settings reload) — $script:config = $fresh
4. In the installation-wizard callback inside $btnSetupWizard.Add_Click — $script:config = Get-Content ...
5. In Show-MainForm — $script:config = Get-Content -Path $script:configPath -Raw | ConvertFrom-Json

Pattern for each:

```powershell
$script:config = $cfg
$global:PMConfig = $script:config       # keep closures in sync
```

Other mutations (e.g. $script:config.Books.Foo = X, Add-Member on $script:config) mutate in place — the object identity is preserved, so $global:PMConfig automatically sees them. Only reassignments need the extra line.

Step B1e — also allow Search-Parts to work from the closure. Find:

```powershell
    # 1. Local Parts Room
    $localFiles = Get-ChildItem -Path (Join-Path $config.PartsRoomDirectory "*.csv") -File
```

New:

```powershell
    # 1. Local Parts Room
    $cfg = if ($global:PMConfig) { $global:PMConfig } else { $config }
    $localFiles = Get-ChildItem -Path (Join-Path $cfg.PartsRoomDirectory "*.csv") -File
```

And in the Same Day section right below it:

Old:

```powershell
    $sameDayDir = Join-Path $config.PartsRoomDirectory "Same Day Parts Room"
```

New:

```powershell
    $sameDayDir = Join-Path $cfg.PartsRoomDirectory "Same Day Parts Room"
```

---

B2. Search tab: default should be "no book/section limit"

Find: the block near the top of Setup-SearchTab that creates $script:comboBook and $script:comboSection.

Goal:

· Book dropdown's first option is (All Books)
· Section dropdown's first option is (All Sections)
· Both default to (All ...) — no scope narrowing
· Search button is enabled by default
· User picks a specific book/section to narrow; picking a book repopulates sections with (All Sections) at index 0

Old — the combo creation block:

```powershell
	$script:comboBook = New-Object System.Windows.Forms.ComboBox
	$script:comboBook.Location = New-Object System.Drawing.Point(590, 15)
	$script:comboBook.Size = New-Object System.Drawing.Size(190, 25)
	$script:comboBook.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$script:comboBook.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	foreach ($b in (Get-SearchableBooks)) { [void]$script:comboBook.Items.Add($b) }
	$topPanel.Controls.Add($script:comboBook)
	...
	$script:comboSection = New-Object System.Windows.Forms.ComboBox
	$script:comboSection.Location = New-Object System.Drawing.Point(590, 49)
	$script:comboSection.Size = New-Object System.Drawing.Size(190, 25)
	$script:comboSection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$script:comboSection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$topPanel.Controls.Add($script:comboSection)

	$script:searchSectionKeyMap = @{}

	# Populate sections on book change.
	$script:comboBook.Add_SelectedIndexChanged({
		$script:comboSection.Items.Clear()
		$script:comboSection.SelectedIndex = -1
		$script:searchSectionKeyMap = @{}
		$bookName = "$($script:comboBook.SelectedItem)"
		if ([string]::IsNullOrWhiteSpace($bookName)) { return }
		foreach ($sec in (Get-SectionsForBook -BookName $bookName)) {
			[void]$script:comboSection.Items.Add($sec.Display)
			$script:searchSectionKeyMap[$sec.Display] = $sec.Key
		}
		if ($script:comboSection.Items.Count -gt 0) {
			$script:comboSection.SelectedIndex = 0
		}
	})
```

New:

```powershell
	$script:comboBook = New-Object System.Windows.Forms.ComboBox
	$script:comboBook.Location = New-Object System.Drawing.Point(590, 15)
	$script:comboBook.Size = New-Object System.Drawing.Size(190, 25)
	$script:comboBook.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$script:comboBook.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	[void]$script:comboBook.Items.Add('(All Books)')
	foreach ($b in (Get-SearchableBooks)) { [void]$script:comboBook.Items.Add($b) }
	$topPanel.Controls.Add($script:comboBook)
	...
	$script:comboSection = New-Object System.Windows.Forms.ComboBox
	$script:comboSection.Location = New-Object System.Drawing.Point(590, 49)
	$script:comboSection.Size = New-Object System.Drawing.Size(190, 25)
	$script:comboSection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$script:comboSection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	[void]$script:comboSection.Items.Add('(All Sections)')
	$topPanel.Controls.Add($script:comboSection)

	$script:searchSectionKeyMap = @{}

	# Populate sections on book change.
	$script:comboBook.Add_SelectedIndexChanged({
		$script:comboSection.Items.Clear()
		[void]$script:comboSection.Items.Add('(All Sections)')
		$script:comboSection.SelectedIndex = 0
		$script:searchSectionKeyMap = @{}
		$bookName = "$($script:comboBook.SelectedItem)"
		if ([string]::IsNullOrWhiteSpace($bookName) -or $bookName -eq '(All Books)') { return }
		foreach ($sec in (Get-SectionsForBook -BookName $bookName)) {
			[void]$script:comboSection.Items.Add($sec.Display)
			$script:searchSectionKeyMap[$sec.Display] = $sec.Key
		}
	})

	# Default to (All Books) / (All Sections) so search works immediately.
	$script:comboBook.SelectedIndex = 0
	$script:comboSection.SelectedIndex = 0
```

Then, in the search-click handler, treat (All Books) / (All Sections) as "no scope." Find the CrossRef block in the search handler:

Old:

```powershell
        $crossRefResults = @()
        $selectedBookName       = "$($script:comboBook.SelectedItem)".Trim()
        $selectedSectionDisplay = "$($script:comboSection.SelectedItem)".Trim()
        $selectedSectionKey     = $script:searchSectionKeyMap[$selectedSectionDisplay]

        if (-not [string]::IsNullOrWhiteSpace($selectedBookName) -and
            -not [string]::IsNullOrWhiteSpace($selectedSectionKey)) {
            ...
        } else {
            Write-Log "CrossRef skipped: no book/section selected."
        }
```

New:

```powershell
        $crossRefResults = @()
        $selectedBookName       = "$($script:comboBook.SelectedItem)".Trim()
        $selectedSectionDisplay = "$($script:comboSection.SelectedItem)".Trim()
        $selectedSectionKey     = $script:searchSectionKeyMap[$selectedSectionDisplay]

        # "(All Books)" means: search every book.  "(All Sections)" means:
        # search every section within the chosen book.  Empty means: no scope.
        $scopeAll = [string]::IsNullOrWhiteSpace($selectedBookName) -or $selectedBookName -eq '(All Books)'

        if ($scopeAll) {
            # Scan every searchable book across every section.
            foreach ($bookName in (Get-SearchableBooks)) {
                foreach ($sec in (Get-SectionsForBook -BookName $bookName)) {
                    try { $sectionData = Import-Csv -Path $sec.FilePath }
                    catch { $sectionData = @() }

                    if ($nsnNoMatch -or $oemNoMatch) { $filtered = @() } else {
                        $filtered = @($sectionData | Where-Object {
                            $nsnOk = $true
                            if ($nsnActive) {
                                $rowStock = ([string]$_.'STOCK NO.') -replace '[^0-9]', ''
                                $nsnOk = $rowStock -match $nsnPattern
                            }
                            $oemOk = $true
                            if ($oemActive) {
                                $rowPart = ([string]$_.'PART NO.' -replace '[^A-Za-z0-9]', '').ToUpper()
                                $oemOk = ($rowPart -and ($rowPart -match $oemPattern))
                            }
                            $descOk = $true
                            if ($descActive) {
                                $rowDesc = (([string]$_.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                                foreach ($token in $descTokens) {
                                    if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                                }
                            }
                            $nsnOk -and $oemOk -and $descOk
                        })
                    }

                    foreach ($item in @($filtered)) {
                        $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $bookName -Force
                        $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $sec.Display -Force
                        $crossRefResults += $item
                    }
                }
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($selectedSectionKey)) {
            # A specific book + specific section was chosen.
            $secs = Get-SectionsForBook -BookName $selectedBookName
            $match = $secs | Where-Object { $_.Key -eq $selectedSectionKey } | Select-Object -First 1
            if ($match) {
                try { $sectionData = Import-Csv -Path $match.FilePath }
                catch { $sectionData = @() }
                if ($nsnNoMatch -or $oemNoMatch) { $filtered = @() } else {
                    $filtered = @($sectionData | Where-Object {
                        $nsnOk = $true
                        if ($nsnActive) {
                            $rowStock = ([string]$_.'STOCK NO.') -replace '[^0-9]', ''
                            $nsnOk = $rowStock -match $nsnPattern
                        }
                        $oemOk = $true
                        if ($oemActive) {
                            $rowPart = ([string]$_.'PART NO.' -replace '[^A-Za-z0-9]', '').ToUpper()
                            $oemOk = ($rowPart -and ($rowPart -match $oemPattern))
                        }
                        $descOk = $true
                        if ($descActive) {
                            $rowDesc = (([string]$_.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                            foreach ($token in $descTokens) {
                                if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                            }
                        }
                        $nsnOk -and $oemOk -and $descOk
                    })
                }
                foreach ($item in @($filtered)) {
                    $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $selectedBookName       -Force
                    $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $selectedSectionDisplay -Force
                    $crossRefResults += $item
                }
            }
        } else {
            # Specific book, "(All Sections)" — scan every section of that book.
            foreach ($sec in (Get-SectionsForBook -BookName $selectedBookName)) {
                try { $sectionData = Import-Csv -Path $sec.FilePath }
                catch { $sectionData = @() }
                if ($nsnNoMatch -or $oemNoMatch) { $filtered = @() } else {
                    $filtered = @($sectionData | Where-Object {
                        $nsnOk = $true
                        if ($nsnActive) {
                            $rowStock = ([string]$_.'STOCK NO.') -replace '[^0-9]', ''
                            $nsnOk = $rowStock -match $nsnPattern
                        }
                        $oemOk = $true
                        if ($oemActive) {
                            $rowPart = ([string]$_.'PART NO.' -replace '[^A-Za-z0-9]', '').ToUpper()
                            $oemOk = ($rowPart -and ($rowPart -match $oemPattern))
                        }
                        $descOk = $true
                        if ($descActive) {
                            $rowDesc = (([string]$_.'PART DESCRIPTION' -replace '[^A-Za-z0-9]', ' ').ToLower() -replace '\s+', ' ').Trim()
                            foreach ($token in $descTokens) {
                                if (-not $rowDesc.Contains($token)) { $descOk = $false; break }
                            }
                        }
                        $nsnOk -and $oemOk -and $descOk
                    })
                }
                foreach ($item in @($filtered)) {
                    $item | Add-Member -NotePropertyName 'Handbook'     -NotePropertyValue $selectedBookName -Force
                    $item | Add-Member -NotePropertyName 'Section Name' -NotePropertyValue $sec.Display -Force
                    $crossRefResults += $item
                }
            }
        }
```

Also change the button gating so it stays enabled by default. Find (after the fixes from A1, this is the block you re-added):

Old:

```powershell
    $updateSearchEnabled = {
        $searchButton.Enabled = (
            $script:comboBook.SelectedIndex    -ge 0 -and
            $script:comboSection.SelectedIndex -ge 0
        )
    }.GetNewClosure()
    ...
    $searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -and $script:comboSection.SelectedIndex -ge 0)
```

New (just don't gate — always enabled):

```powershell
    $searchButton.Enabled = $true
```

You can leave the $updateSearchEnabled closure out entirely; it's no longer needed.

---

B3. eMARS p_user dropdown: doesn't populate + saved ID wiped

Three root causes:

1. Closure hides config. $refreshEmarsUsers calls Get-EmarsUserList → Get-EmarsSetting → $script:config (empty in closure). Fixed by B1c above.
2. Endpoint may not exist. Get-EmarsUserList calls ecbm.ecbm_user.getUsers, which is a placeholder guess. If it fails, the list stays empty (silent).
3. DropDownList drops saved text. $emarsUserBox.DropDownStyle = DropDownList means .Text = "jsmith" does nothing when items don't contain "jsmith". Then Save reads "" and writes EmarsUser = "" — that's the wipe you saw.

Fix B3a — Get-EmarsUserList diagnostic + fallback. Find:

```powershell
function Get-EmarsUserList {
    param([string]$SiteId = '')

    if ([string]::IsNullOrWhiteSpace($SiteId)) { $SiteId = Get-EmarsSetting -Name 'EmarsSite' }
    if ([string]::IsNullOrWhiteSpace($SiteId)) { throw "eMARS Site ID is not set." }

    # TODO: replace Proc and Query with the values observed in DevTools.
    $xml = Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_user.getUsers' `
        -Query "p_site=$SiteId" `
        -Label "getUsers(site $SiteId)"
    ...
}
```

New:

```powershell
function Get-EmarsUserList {
    param([string]$SiteId = '')

    if ([string]::IsNullOrWhiteSpace($SiteId)) { $SiteId = Get-EmarsSetting -Name 'EmarsSite' }
    if ([string]::IsNullOrWhiteSpace($SiteId)) { throw "eMARS Site ID is not set." }

    # Preferred path: ask eMARS.  This endpoint is a placeholder — if it
    # 404s or returns HTML, we fall back to EmarsUsers.csv (below).
    $users = @()
    try {
        $xml = Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_user.getUsers' `
            -Query "p_site=$SiteId" `
            -Label "getUsers(site $SiteId)"
        foreach ($row in (Get-EmarsRows -Doc $xml)) {
            $userId = Get-EmarsField -Row $row -Names @('USERID','USER_ID','P_USER') -Index 0
            $name   = Get-EmarsField -Row $row -Names @('USERNAME','FULLNAME','NAME') -Index 1
            if ([string]::IsNullOrWhiteSpace($userId)) { continue }
            $users += [PSCustomObject]@{
                UserId = $userId
                Name   = if ($name) { $name } else { $userId }
            }
        }
        if ($users.Count -gt 0) { return ,$users }
    } catch {
        Write-Log "Get-EmarsUserList: eMARS endpoint failed — $($_.Exception.Message); trying CSV fallback."
    }

    # Fallback path: Dropdown CSVs\EmarsUsers.csv with header
    #   Site ID,User ID,Full Name
    $cfg = if ($global:PMConfig) { $global:PMConfig } else { $script:config }
    if ($cfg) {
        $csvPath = Join-Path $cfg.DropdownCsvsDirectory 'EmarsUsers.csv'
        if (Test-Path $csvPath) {
            try {
                $rows = Import-Csv -Path $csvPath
                foreach ($r in @($rows | Where-Object { "$($_.'Site ID')" -eq $SiteId })) {
                    $uid = "$($r.'User ID')".Trim()
                    if ([string]::IsNullOrWhiteSpace($uid)) { continue }
                    $users += [PSCustomObject]@{
                        UserId = $uid
                        Name   = if ("$($r.'Full Name')".Trim()) { "$($r.'Full Name')".Trim() } else { $uid }
                    }
                }
                if ($users.Count -gt 0) {
                    Write-Log "Get-EmarsUserList: loaded $($users.Count) user(s) from EmarsUsers.csv for site $SiteId."
                } else {
                    Write-Log "Get-EmarsUserList: EmarsUsers.csv has no rows for site $SiteId."
                }
            } catch {
                Write-Log "Get-EmarsUserList: EmarsUsers.csv read failed — $($_.Exception.Message)"
            }
        }
    }
    return ,$users
}
```

Fix B3b — switch the combo to editable + preserve saved text.

Old:

```powershell
	$emarsUserBox = New-Object System.Windows.Forms.ComboBox
	$emarsUserBox.Size = New-Object System.Drawing.Size(440, 24)
	$emarsUserBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	New-SettingRow -Parent $gEmars -Top 62 -LabelText "eMARS User ID (p_user):" -InputControl $emarsUserBox
```

New:

```powershell
	$emarsUserBox = New-Object System.Windows.Forms.ComboBox
	$emarsUserBox.Size = New-Object System.Drawing.Size(440, 24)
	# DropDown (editable) so a saved ID survives even when the endpoint or
	# EmarsUsers.csv cannot supply the current list.  User picks from the list
	# when it works, types freely when it doesn't.
	$emarsUserBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
	New-SettingRow -Parent $gEmars -Top 62 -LabelText "eMARS User ID (p_user):" -InputControl $emarsUserBox
```

Fix B3c — attach $refreshEmarsUsers at setup time, and seed the saved ID.

Find (near the end of Setup-SettingsTab, right after Load-WtSettingsValues):

Old:

```powershell
    # --- Initial load ---
    Load-WtSettingsValues
```

New:

```powershell
    # --- Initial load ---
    Load-WtSettingsValues

    # Attach the user-list refresh at setup time (not just in Reload handler).
    $emarsSiteBox.Add_TextChanged($refreshEmarsUsers)
    & $refreshEmarsUsers
```

Find the $refreshEmarsUsers closure definition and make it preserve the current text even when the list is empty:

Old:

```powershell
	$refreshEmarsUsers = {
		try {
			$script:wtSettings.EmarsUserBox.Items.Clear()
			$site = "$($script:wtSettings.EmarsSiteBox.Text)".Trim()
			if ([string]::IsNullOrWhiteSpace($site)) { return }
			foreach ($u in (Get-EmarsUserList -SiteId $site)) {
				[void]$script:wtSettings.EmarsUserBox.Items.Add("$($u.UserId)  —  $($u.Name)")
			}
			# Preserve whatever the config already stored, if it appears in the list.
			$saved = "$($script:config.WorkTracking.EmarsUser)".Trim()
			if ($saved) {
				for ($i = 0; $i -lt $script:wtSettings.EmarsUserBox.Items.Count; $i++) {
					if ("$($script:wtSettings.EmarsUserBox.Items[$i])".StartsWith("$saved ")) {
						$script:wtSettings.EmarsUserBox.SelectedIndex = $i
						break
					}
				}
			}
		} catch {
			Write-Log "Get-EmarsUserList failed: $($_.Exception.Message)"
		}
	}.GetNewClosure()
```

New:

```powershell
	$refreshEmarsUsers = {
		try {
			$box = $script:wtSettings.EmarsUserBox
			# Remember whatever text is currently visible — that is either the
			# user's last selection or a manually-typed ID that must survive.
			$keepText = "$($box.Text)".Trim()

			$box.Items.Clear()
			$site = "$($script:wtSettings.EmarsSiteBox.Text)".Trim()
			if (-not [string]::IsNullOrWhiteSpace($site)) {
				$list = @(Get-EmarsUserList -SiteId $site)
				foreach ($u in $list) {
					[void]$box.Items.Add("$($u.UserId)  —  $($u.Name)")
				}
				Write-Log "refreshEmarsUsers: site='$site', loaded $($list.Count) user(s)"
			}

			# Prefer the saved config value if nothing is currently in the box.
			if ([string]::IsNullOrWhiteSpace($keepText)) {
				$cfg = if ($global:PMConfig) { $global:PMConfig } else { $script:config }
				if ($cfg) { $keepText = "$($cfg.WorkTracking.EmarsUser)".Trim() }
			}

			if (-not [string]::IsNullOrWhiteSpace($keepText)) {
				# If an entry matching the ID exists, select it.
				$matched = $false
				for ($i = 0; $i -lt $box.Items.Count; $i++) {
					if ("$($box.Items[$i])".StartsWith("$keepText ")) {
						$box.SelectedIndex = $i
						$matched = $true
						break
					}
				}
				# Otherwise type it into the editable box so it survives Save.
				if (-not $matched) { $box.Text = $keepText }
			}
		} catch {
			Write-Log "Get-EmarsUserList failed: $($_.Exception.Message)"
		}
	}.GetNewClosure()
```

Fix B3d — unstick the Reload handler. Find:

Old:

```powershell
            Load-WtSettingsValues
			$script:wtSettings.EmarsSiteBox.Add_TextChanged($refreshEmarsUsers)
			& $refreshEmarsUsers
            Write-Log "Reload: UI updated"
```

New:

```powershell
            Load-WtSettingsValues
            & $refreshEmarsUsers
            Write-Log "Reload: UI updated"
```

(The Add_TextChanged call moves to initial setup in B3c — attaching it a second time on every reload would fire the handler twice per keystroke.)

Fix B3e — Save handler writes what's actually in the box.

Old:

```powershell
			$emarsUserRaw = "$($c.EmarsUserBox.Text)".Trim()
			$emarsUser    = ($emarsUserRaw -split '\s+—\s+', 2)[0].Trim()
			if ([string]::IsNullOrWhiteSpace($emarsUser)) { $emarsUser = $emarsUserRaw }
```

New (unchanged logic, just a comment):

```powershell
			# Combo can carry "USERID  —  Full Name" (from list selection) or
			# just "USERID" (typed).  Split off the display suffix either way.
			$emarsUserRaw = "$($c.EmarsUserBox.Text)".Trim()
			$emarsUser    = ($emarsUserRaw -split '\s+—\s+', 2)[0].Trim()
			if ([string]::IsNullOrWhiteSpace($emarsUser)) { $emarsUser = $emarsUserRaw }
```

---

Verify

After all edits, before running:

```powershell
$errs = $null
[System.Management.Automation.Language.Parser]::ParseFile("$PWD\UI-Script.ps1", [ref]$null, [ref]$errs) | Out-Null
if ($errs) { $errs | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" } } else { "OK" }
```

Expected in the log at startup / first use:

· No The property 'Enabled' cannot be found (A1 / B2)
· No Set-ThemeWindowChrome error (A2)
· No flood of Get-Jobs: read ... — should be one or two per interaction (A3)
· Search tab: Book dropdown shows (All Books) selected, Section shows (All Sections) selected, Search button enabled — no auto-book selection (B2)
· Clicking a book in the Search tab populates Section with (All Sections) plus every section of that book (B1 + B2)
· Part Picker in the Job Log: sections populate when a book is picked (B1)
· Settings → eMARS → Save then reopen: your user ID is still in the box (B3b/c)
· After the site box has any text, UI.log shows refreshEmarsUsers: site='...' loaded N user(s) — N will be 0 if neither the endpoint nor EmarsUsers.csv returns anything, but the box no longer blanks out

---

What still needs a manual step

The ecbm.ecbm_user.getUsers endpoint is a guess. If your eMARS instance doesn't have it (or it returns HTML), the dropdown will stay empty — but thanks to B3b/c it will keep whatever ID you type and save it. To populate the list without the endpoint, add Dropdown CSVs\EmarsUsers.csv with:

```csv
Site ID,User ID,Full Name
345,jsmith,Joseph Shaw
345,jdoe,Jane Doe
```

Reload settings and the list will populate from CSV. Later, when you find the real endpoint name (via F12 → Network on any eMARS screen that lists users), change the Invoke-EmarsXmlGet -Proc 'ecbm.ecbm_user.getUsers' line to the real one and the CSV becomes a fallback only.