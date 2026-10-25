PS C:\Users\YYQTP0\Desktop> .\UI-Script.ps1
2026-10-25 12:44:58 - Config loaded successfully
2026-10-25 12:44:58 - Active UI theme: Default
2026-10-25 12:45:00 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:45:01 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:45:01 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:45:01 - Job Log sub-tab setup completed.
2026-10-25 12:45:02 - Sync-PmArchive: processed=0 removed=0
2026-10-25 12:45:02 - Load-PmMachines: dir=C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\PM Checklists fileCount=55
2026-10-25 12:45:03 - Refresh-PmMachineTabs: selected=2026-10-25 all=55 shown=2
2026-10-25 12:45:06 - Get-Jobs: read 45 job(s) from C:\Users\YYQTP0\Desktop\PartsMgmt\Work Tracking\Jobs.json
2026-10-25 12:45:06 - PM Checklist sub-tab setup completed.
2026-10-25 12:45:07 - Weekly Worksheet sub-tab setup completed.
[Historian] machine filter: kind='All' -> 20 machine(s): ATU 1, ATU 2, ATU 3, ATU 4, DBCS 25, DBCS 28, DBCS 30, DBCS 31, DBCS 32, DBCS 33, DBCS 34, DBCS 35, DBCS 49, DBCS 58, DBCS 59, DBCS 61, DBCS 62, DBCS 63, DBCS 64, WORKSHEET
2026-10-25 12:45:13 - Historian sub-tab setup completed.
2026-10-25 12:45:13 - Processing Books from config:
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-229 Delivery Bar Code Sorter Phase 4 - DBCS (Vol. E)\MS-229 Delivery Bar Code Sorter Phase 4 - DBCS (Vol. E).xlsx
2026-10-25 12:45:13 - Added file: MS-229 Delivery Bar Code Sorter Phase 4 - DBCS Vol E
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 (Vol. C)\MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 (Vol. C).xlsx
2026-10-25 12:45:13 - Added file: MS-298 DELIVERY BAR CODE SORTER Phase 3-5 LAU PH2 - DBCS 3-5 LAU PH2 Vol C
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-212 Wide Field Of View Camera - WFOV (Vol. B)\MS-212 Wide Field Of View Camera - WFOV (Vol. B).xlsx
2026-10-25 12:45:13 - Added file: MS-212 Wide Field Of View Camera - WFOV Vol B
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-243 Automatic Tray Unsleever - ATU (Vol. B)\MS-243 Automatic Tray Unsleever - ATU (Vol. B).xlsx
2026-10-25 12:45:13 - Added file: MS-243 Automatic Tray Unsleever - ATU Vol B
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-242 Enhanced Airline Assignment - EAA (Vol. B)\MS-242 Enhanced Airline Assignment - EAA (Vol. B).xlsx
2026-10-25 12:45:13 - Added file: MS-242 Enhanced Airline Assignment - EAA Vol B
2026-10-25 12:45:13 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-313 High Speed Tray Sorter - HSTS (Vol. B)\MS-313 High Speed Tray Sorter - HSTS (Vol. B).xlsx
2026-10-25 12:45:13 - Added file: MS-313 High Speed Tray Sorter - HSTS Vol B
2026-10-25 12:45:14 - Checking file: C:\Users\YYQTP0\Desktop\PartsMgmt\Parts Books\MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai (Vol. G2)\MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai (Vol. G2).xlsx
2026-10-25 12:45:14 - Added file: MS-178 Automated Flat Sorting Machine 100 - Automatic Induction - AFSM 100-ai Vol G2
2026-10-25 12:45:14 - Setting up Search tab...
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
    + FullyQualifiedErrorId : PropertyNotFound

The property 'Enabled' cannot be found on this object. Verify that the property exists and can be set.
At C:\Users\YYQTP0\Desktop\UI-Script.ps1:14028 char:2
+     $searchButton.Enabled = ($script:comboBook.SelectedIndex -ge 0 -a ...
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : PropertyNotFound

2026-10-25 12:45:15 - Search tab setup completed.
2026-10-25 12:45:15 - Setting up Settings tab...
2026-10-25 12:45:18 - Load-WtSettingsValues: configPath='C:\Users\YYQTP0\Desktop\PartsMgmt\Config.json'
2026-10-25 12:45:18 - Load-WtSettingsValues: techName='Joseph Shaw' email='default@example.com'
2026-10-25 12:45:18 - Settings tab setup completed.
2026-10-25 12:45:18 - UI setup completed
2026-10-25 12:45:19 - Theme applied: Default
Set-ThemeWindowChrome : The term 'Set-ThemeWindowChrome' is not recognized as the name of a cmdlet, function, script
file, or operable program. Check the spelling of the name, or if a path was included, verify that the path is correct
and try again.
At C:\Users\YYQTP0\Desktop\UI-Script.ps1:912 char:21
+                     Set-ThemeWindowChrome -Form $Root
+                     ~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : ObjectNotFound: (Set-ThemeWindowChrome:String) [], CommandNotFoundException
    + FullyQualifiedErrorId : CommandNotFoundException

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

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
    + FullyQualifiedErrorId : PropertyNotFound

Cancel
2026-10-25 12:45:47 - Application closed
PS C:\Users\YYQTP0\Desktop>
