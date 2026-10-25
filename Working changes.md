You're right about the eMARS user — different credential, keep it manual. Dropping that part.

The bug is the Test button — not the site ID check itself.

Look at $emarsTestBtn.Add_Click({...}) in Setup-SettingsTab. It references $emarsSiteBox, $emarsUserBox, $emarsOriginBox, $emarsSubsiteBox, $emarsFacBox as local variables:

```powershell
$cfg.WorkTracking.EmarsSite = "$($emarsSiteBox.Text)".Trim()
```

Those locals live in Setup-SettingsTab's scope and die when that function returns. By the time your click fires the handler, they're $null. In PS 5.1 without strict mode, $null.Text doesn't throw — it returns $null, so "$null".Trim() becomes "".

So clicking Test eMARS Connection does this:

1. Sets $cfg.WorkTracking.EmarsSite = "" (silently wiping whatever you typed)
2. Persists it to disk — EmarsSite: "" in Config.json
3. Calls Test-EmarsConnection
4. Get-EmarsSetting -Name 'EmarsSite' → ""
5. Shows "eMARS Site ID is not set — fill it in (and save) first."

You typed it, Save wrote it, Test erased it. That's the symptom.

This is the exact class of bug the codebase already handles elsewhere — the close-button handlers carry a comment about .GetNewClosure() being required in PS 5.1 for this reason. The Test handler just missed the pattern.

The Save handler is fine because it uses $script:wtSettings.EmarsSiteBox (script-scoped hashtable holding the live TextBox references), not the dead locals. That's why Save works and Test doesn't.

Fix

Rewrite the Test handler to use $script:wtSettings the same way Save does:

```powershell
$emarsTestBtn.Add_Click({
    try {
        $c = $script:wtSettings
        if (-not $c) { throw "Settings controls are not initialized yet." }

        $cfg = $script:config
        $cfg.WorkTracking.EmarsOrigin  = "$($c.EmarsOriginBox.Text)".Trim()
        $cfg.WorkTracking.EmarsUser    = "$($c.EmarsUserBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSite    = "$($c.EmarsSiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsSubsite = "$($c.EmarsSubsiteBox.Text)".Trim()
        $cfg.WorkTracking.EmarsFacId   = "$($c.EmarsFacBox.Text)".Trim()
        ($cfg | ConvertTo-Json -Depth 12) | Set-Content -Path $script:configPath -Encoding UTF8
        $script:config = $cfg

        Write-Log "Test eMARS Connection: values applied and saved; running endpoint checks..."
        $results = Test-EmarsConnection
        foreach ($line in $results) { Write-Log "Test eMARS Connection: $line" }
        [System.Windows.Forms.MessageBox]::Show(
            ("eMARS connectivity check:`r`n`r`n" + ($results -join "`r`n") + "`r`n`r`n(Details are also in UI.log)"),
            "eMARS Test", "OK", "Information")
    } catch {
        Write-Log "Test eMARS Connection failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Test failed: $($_.Exception.Message)", "eMARS Test", "OK", "Error")
    }
})
```

$script:wtSettings is populated right before Load-WtSettingsValues at the end of Setup-SettingsTab, and it holds live references to the TextBox instances — so the handler always reads the current text, no matter when the user clicks.

(The alternative is .GetNewClosure() on the existing scriptblock. Same result, but $script:wtSettings is what Save already uses and it won't rot when someone adds another eMARS field later.)

Confirming the diagnosis

1. Set the site ID, click Save.
2. Open Config.json — verify "EmarsSite": "345" (or whatever).
3. Click Test.
4. Re-open Config.json — on the current build, it'll now say "EmarsSite": "".

After the fix, step 4 leaves the value intact and UI.log shows Test eMARS Connection: values applied and saved; running endpoint checks....

If you still get "not set" after the fix, check the Save: wrote ... chars to ...\Config.json line in UI.log — the path is hardcoded as $PSScriptRoot\PartsMgmt\Config.json, so if the app was set up with a custom root via the Installation Wizard, that path could be the wrong file. But the symptom you're describing is fully explained by the Test button, so I'd fix that first and re-check.