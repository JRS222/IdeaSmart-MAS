Root cause

There is one misplaced closing brace in Setup-SearchTab that closes the function early. Everything after that brace (the Search button, the three ListViews, the bottom bar, and their event handlers) is being executed at script scope, not inside the function — that's why $searchPanel, $topPanel, $script:comboBook, $script:textBoxNSN, $bottomBar, etc. are all $null when you reach $topPanel.Controls.Add(...).

The brace count is still balanced overall — but with the } in the wrong place — so PowerShell can't parse it cleanly (that's why you got Unexpected token '}' / ')' at first, then later a runtime "null-valued expression" once the file happened to balance).

The two edits

1. Delete the stray } right after the combo default-index assignment

Find (around line 14100 in your copy):

```powershell
	# Default to (All Books) / (All Sections) so search works immediately.
	$script:comboBook.SelectedIndex = 0
	$script:comboSection.SelectedIndex = 0
		}          # <-- this is the stray brace; DELETE it

	

    # Search button
```

Replace with:

```powershell
	# Default to (All Books) / (All Sections) so search works immediately.
	$script:comboBook.SelectedIndex = 0
	$script:comboSection.SelectedIndex = 0

    # Search button
```

2. Add the missing } after Write-Log "Search tab setup completed."

Find (right before function Get-SupervisorEmail):

```powershell
    Write-Log "Search tab setup completed."


# Function to get or set the supervisor email
function Get-SupervisorEmail {
```

Replace with:

```powershell
    Write-Log "Search tab setup completed."
}   # <-- ADD this closing brace for Setup-SearchTab


# Function to get or set the supervisor email
function Get-SupervisorEmail {
```

That's the whole search tab fix. After these two edits the Search tab will actually get built inside the function, $topPanel, $searchPanel, $comboBook, etc. will all exist, and the Open Figures / Take Part(s) Out buttons will enable correctly.

---

Secondary issue: Write-Log inside $refreshEmarsUsers

Your output also shows:

```
Write-Log : The term 'Write-Log' is not recognized ...
At ... UI-Script.ps1:13622 char:4
Write-Log "Get-EmarsUserList failed: $($_.Exception.Messa ...
```

This is the .GetNewClosure() trap the file itself documents in the block starting at # IMPORTANT (PowerShell 5.1): inside a .GetNewClosure() block, $script: refers to the CLOSURE's own (empty) script scope. Inside that closure:

· $script:wtSettings is $null (the closure's $script: scope is empty)
· Write-Log and Get-EmarsUserList are also not resolvable from the closure's function scope

Minimal fix — move the closure creation to after $script:wtSettings is assigned, and capture the functions as locals (this is the exact pattern the file already uses for Set-ThemeWindowChrome).

Find the current $refreshEmarsUsers = { ... }.GetNewClosure() block (around line 13610) and delete it entirely. Then, further down, find this block:

```powershell
    # --- Store control references ---
    $script:wtSettings = @{
        ...
    }

    # Wire auto-recompute on any budget input.
    foreach ($box in @($dayLenBox, $lunchBox, $washupBox, $startupBox,
                       $break1Box, $break2Box, $endWashBox, $paperBox)) {
        $box.Add_ValueChanged({ Update-WorkTargetDisplay })
    }

    # --- Initial load ---
    Load-WtSettingsValues

    # Attach the user-list refresh at setup time (not just in Reload handler).
    $emarsSiteBox.Add_TextChanged($refreshEmarsUsers)
    & $refreshEmarsUsers
```

Replace with:

```powershell
    # --- Store control references ---
    $script:wtSettings = @{
        ...
    }

    # Wire auto-recompute on any budget input.
    foreach ($box in @($dayLenBox, $lunchBox, $washupBox, $startupBox,
                       $break1Box, $break2Box, $endWashBox, $paperBox)) {
        $box.Add_ValueChanged({ Update-WorkTargetDisplay })
    }

    # Capture script state + functions as locals BEFORE creating the closure.
    # (.GetNewClosure() gets an empty $script: scope in PS 5.1 — the file's
    #  own comment near Initialize-ResponsiveLayout explains this.)
    $wt          = $script:wtSettings
    $getUsersFn  = Get-Command -Name Get-EmarsUserList -CommandType Function
    $logFn       = Get-Command -Name Write-Log        -CommandType Function
    $cfgForUsers = $script:config

    $refreshEmarsUsers = {
        try {
            $box = $wt.EmarsUserBox
            $keepText = "$($box.Text)".Trim()
            $box.Items.Clear()
            $site = "$($wt.EmarsSiteBox.Text)".Trim()
            if (-not [string]::IsNullOrWhiteSpace($site)) {
                $list = @(& $getUsersFn -SiteId $site)
                foreach ($u in $list) {
                    [void]$box.Items.Add("$($u.UserId)  —  $($u.Name)")
                }
                & $logFn "refreshEmarsUsers: site='$site', loaded $($list.Count) user(s)"
            }

            if ([string]::IsNullOrWhiteSpace($keepText)) {
                $src = if ($global:PMConfig) { $global:PMConfig } else { $cfgForUsers }
                if ($src) { $keepText = "$($src.WorkTracking.EmarsUser)".Trim() }
            }
            if (-not [string]::IsNullOrWhiteSpace($keepText)) {
                $matched = $false
                for ($i = 0; $i -lt $box.Items.Count; $i++) {
                    if ("$($box.Items[$i])".StartsWith("$keepText ")) {
                        $box.SelectedIndex = $i; $matched = $true; break
                    }
                }
                if (-not $matched) { $box.Text = $keepText }
            }
        } catch {
            & $logFn "Get-EmarsUserList failed: $($_.Exception.Message)"
        }
    }.GetNewClosure()

    # --- Initial load ---
    Load-WtSettingsValues

    # Attach the user-list refresh at setup time.
    $emarsSiteBox.Add_TextChanged($refreshEmarsUsers)
    & $refreshEmarsUsers
```

---

Verify after patching

From PS C:\Users\YYQTP0\Desktop>:

```powershell
# 1. Confirm the file parses cleanly (no output = OK)
[System.Management.Automation.Language.Parser]::ParseFile(
    "$PWD\UI-Script.ps1", [ref]$null, [ref]$errors)
$errors
```

You should get zero parse errors. Then run:

```powershell
.\UI-Script.ps1
```

Expected behaviour after the fix:

Before After
Search tab opens empty, but typing in NSN/Book/Section boxes throws null-ref Search tab renders normally; Book: / Section: dropdowns are populated
"Open Figures" and "Take Part(s) Out" fail Both buttons work
Every run logs Write-Log ... not recognized at refreshEmarsUsers Clean log; user dropdown populates on site change
Search results never appear Search populates all three ListViews (Availability, Same Day, Cross Reference)

The .GetNewClosure() fix is the one PowerShell 5.1 quirk that isn't obvious from the error message — the file's own header comment says it best: "inside a .GetNewClosure() block, $script: refers to the CLOSURE's own (empty) script scope, not the main script scope." Capturing $wt, $getUsersFn, and $logFn as locals (which .GetNewClosure() does capture) routes around that.