2026-10-25 12:45:47 - Application closed
PS C:\Users\YYQTP0\Desktop> .\UI-Script.ps1
2026-10-25 12:47:17 - Config loaded successfully
2026-10-25 12:47:17 - Active UI theme: Default
2026-10-25 12:47:18 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:47:19 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:47:19 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:47:19 - Job Log sub-tab setup completed.
2026-10-25 12:47:20 - Sync-PmArchive: processed=0 removed=0
2026-10-25 12:47:20 - Load-PmMachines: dir=C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\PM Checklists fileCount=55
2026-10-25 12:47:21 - Refresh-PmMachineTabs: selected=2026-10-25 all=55 shown=2
2026-10-25 12:47:23 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:47:23 - PM Checklist sub-tab setup completed.
2026-10-25 12:47:25 - Weekly Worksheet sub-tab setup completed.
[Historian] machine filter: kind='All' -> 20 machine(s): ATU 1, ATU 2, ATU 3, ATU 4, DBCS 25, DBCS 28, DBCS 30, DBCS 31, DBCS 32, DBCS 33, DBCS 34, DBCS 35, DBCS 49, DBCS 58, DBCS 59, DBCS 61, DBCS 62, DBCS 63, DBCS 64, WORKSHEET
2026-10-25 12:47:31 - Historian sub-tab setup completed.
2026-10-25 12:47:31 - Processing Books from config:
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-229 Delivery Bar Code Sorter Phase 4 - DBCS (Vol. E)\MS-229 Delivery Bar Code Sorter Phase 4 - DBCS (Vol. E).xlsx
2026-10-25 12:47:31 - Added file: MS-229 Delivery Bar Code Sorter Phase 4 - DBCS Vol E
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 (Vol. C)\MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 (Vol. C).xlsx
2026-10-25 12:47:31 - Added file: MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 Vol C
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-212 Wide Field Of View Camera - WFOV (Vol. B)\MS-212 Wide Field Of View Camera - WFOV (Vol. B).xlsx
2026-10-25 12:47:31 - Added file: MS-212 Wide Field Of View Camera - WFOV Vol B
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-243 Automatic Tray Unsleever - ATU (Vol. B)\MS-243 Automatic Tray Unsleever - ATU (Vol. B).xlsx
2026-10-25 12:47:31 - Added file: MS-243 Automatic Tray Unsleever - ATU Vol B
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-242 Enhanced Airline Assignment - EAA (Vol. B)\MS-242 Enhanced Airline Assignment - EAA (Vol. B).xlsx
2026-10-25 12:47:31 - Added file: MS-242 Enhanced Airline Assignment - EAA Vol B
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-313 High Speed Tray Sorter - HSTS (Vol. B)\MS-313 High Speed Tray Sorter - HSTS (Vol. B).xlsx
2026-10-25 12:47:31 - Added file: MS-313 High Speed Tray Sorter - HSTS Vol B
2026-10-25 12:47:31 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai (Vol. G2)\MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai (Vol. G2).xlsx
2026-10-25 12:47:31 - Added file: MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai Vol G2
2026-10-25 12:47:32 - Setting up Search tab...
The property 'Enabled' cannot be found on this object. Verify that the property exists and can be set.
At C:\Users\YYQTP0\Desktop\UI-Script.ps1:14015 char:3
+         $searchButton.Enabled = (
+         ~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : PropertyNotFound

The property 'Enabled' cannot be found on this object. Verify that the property exists and can be set.
At C:\Users\YYQTP0\Desktop\UI-Script.ps1:14015 char:3
+         $searchButton.Enabled = (
+         ~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : PropertyNotFound                                                                                                                                                                                                  The property 'Enabled' cannot be found on this object. Verify that the property exists and can be set.                  At C:\Users\YYQTP0\Desktop\UI-Script.ps1:14028 char:2                                                                   +     $searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -a ...                                             +     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~                                                     + CategoryInfo          : InvalidOperation: (:) [], RuntimeException                                                    + FullyQualifiedErrorId : PropertyNotFound                                                                                                                                                                                                  2026-10-25 12:47:33 - Search tab setup completed.                                                                       2026-10-25 12:47:33 - Setting up Settings tab...                                                                        2026-10-25 12:47:35 - Load-WtSettingsValues: configPath='C:\Users\YYQTP0\Desktop\PartsMgmt\Config.json'                 2026-10-25 12:47:35 - Load-WtSettingsValues: techName='Joseph Shaw' email='default@example.com'                         2026-10-25 12:47:36 - Settings tab setup completed.
2026-10-25 12:47:36 - UI setup completed
2026-10-25 12:47:37 - Theme applied: Default
Set-ThemeWindowChrome : The term 'Set-ThemeWindowChrome' is not recognized as the name of a cmdlet, function, script
file, or operable program. Check the spelling of the name, or if a path was included, verify that the path is correct
and try again.
At C:\Users\YYQTP0\Desktop\UI-Script.ps1:912 char:21
+                     Set-ThemeWindowChrome -Form $Root
+                     ~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : ObjectNotFound: (Set-ThemeWindowChrome:String) [], CommandNotFoundException
    + FullyQualifiedErrorId : CommandNotFoundException

2026-10-25 12:47:55 - Test eMARS Connection: values applied and saved; running endpoint checks...
2026-10-25 12:47:55 - Invoke-PmHttpGet: GET http://emarssu3.eng.usps.gov/pemars/ecbm.ecbm_checklist.getBypassCodes
2026-10-25 12:47:56 - Invoke-PmHttpGet: 200 — 319 char(s)
2026-10-25 12:47:56 - Invoke-PmHttpGet: GET http://emarssu3.eng.usps.gov/pemars/ecbm.ecbm_report_j.getAllChecklists?pSiteID=345&pSubsiteID=01&pAcronym=DBCS
2026-10-25 12:47:56 - Invoke-PmHttpGet: 200 — 6845 char(s)
2026-10-25 12:47:57 - Invoke-PmHttpGet: GET http://emarssu3.eng.usps.gov/pemars/ecbm.ecbm_checklist.getOpenlist?p_machine_id=827433&p_site=345&p_facid=345
2026-10-25 12:47:57 - Invoke-PmHttpGet: 200 — 359 char(s)
2026-10-25 12:47:57 - Invoke-PmHttpGet: GET http://emarssu3.eng.usps.gov/pemars/ecbm.ecbm_checklist.getAChecklist?p_machineid=827433&p_machinetype=DBCS&p_classcode=CK&p_site=345&p_trackno=37915541
2026-10-25 12:47:58 - Invoke-PmHttpGet: 200 — 61515 char(s)
2026-10-25 12:47:58 - Test eMARS Connection: getBypassCodes: OK (4 code(s))
2026-10-25 12:47:58 - Test eMARS Connection: getAllChecklists(DBCS): OK (36 machine(s))
2026-10-25 12:47:58 - Test eMARS Connection: getOpenlist(DBCS 16): OK (checklist 37915541)
2026-10-25 12:47:58 - Test eMARS Connection: getAChecklist(37915541): OK (20 task(s))
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:29 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:30 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:31 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:32 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:32 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:48:32 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
