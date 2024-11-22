# Comprehensive Maintenance Software Guide

## Overview
This maintenance software consists of two PowerShell scripts:
1. Parts-Books-Creator.ps1 (Creates Parts Books Script)
2. UI-Script.ps1 (Main Application)
3. Setup.ps1 (Set Up Application)

## Initial Setup

### Step 1: Run the Setup Script
1. Right-click on `Setup.ps1` and select "Run with PowerShell"
2. When prompted, select a directory where you want to install the maintenance software
3. The script will create the following directory structure:
   ```
   Selected Directory/
   ├── Labor/
   ├── Call Logs/
   ├── Parts Room/
   ├── Parts Books/
   ├── Dropdown CSVs/
   └── Scripts/
      ├── UI-Script.ps1
      └── Parts-Books-Creator.ps1
      └── Setup.ps1
   ```
4. Wait for the setup to complete - this only needs to be done once

### Step 2: Required Files and CSV Formats
After setup, ensure you have the following files in your directory:

#### Config.json
- Created automatically during setup
- Contains paths and configuration settings

#### Required CSV Files (in Dropdown CSVs folder):

1. **Actions.csv**
   ```
   Value
   Replace
   Repair
   Clean
   Adjust
   ```

2. **Causes.csv**
   ```
   Value
   Broken
   Worn
   Misaligned
   Loose
   ```

3. **Machines.csv**
   ```
   Machine,EquipmentNumber
   DBCS,123456
   AFCS,789012
   ```

4. **Nouns.csv**
   ```
   Value
   Belt
   Motor
   Bearing
   Sensor
   ```

5. **Sites.csv**
   ```
   Site ID,Full Name
   ABC,Alpha Bravo Center
   XYZ,X-Ray Yankee Zone
   ```

## Detailed Feature Guide

### Parts Books Tab

#### Features:
- **Open Parts Room**: Opens the current parts room Excel file
- **Create Parts Book**: Launches the Parts Book Creator utility
- **Parts Books List**: Shows available parts books that can be opened

#### How to Use:
1. Click "Open Parts Room" to view current inventory
2. Use "Create Parts Book" to generate a new parts book
3. Click on any listed parts book to open it directly

### Call Logs Tab

#### Features:
- **Add New Call Log**: Record new maintenance calls
- **Add Machine**: Add new equipment to the system
- **Send Logs**: Export logs for a specific date
- **Automatic Labor Log Creation**: Creates labor log entries for calls over 30 minutes

#### How to Add a Call Log:
1. Click "Add New Call Log"
2. Enter date (defaults to current)
3. Select Machine ID from dropdown
4. Select Cause, Action, and Noun from dropdowns
5. Enter Time Down and Time Up
6. Add any relevant notes
7. Click "Add" to save

#### How to Add a Machine:
1. Click "Add Machine"
2. Enter Machine Acronym
3. Enter Equipment Number
4. Click "Add Machine" to save

### Labor Log Tab
Part Details View
Double-clicking on any labor log entry that has parts attached will open a detailed parts window showing:


#### Features:
- **Add Labor Log Entry**: Record labor performed
- **Edit Labor Log Entry**: Modify existing entries
- **Add Parts to Work Order**: Attach parts to work orders
- **Refresh Labor Logs**: Update the display

#### How to Add Labor Log Entry:
1. Click "Add Labor Log Entry"
2. Enter:
   - Date
   - Work Order Number (or leave blank for "Need W/O #")
   - Task Description
   - Machine ID
   - Duration
   - Notes
3. Click "Add" to save

### Actions Tab

#### Features and Usage:

**Parts Management:**
- **Update Parts Books**: Refresh parts book data
- **Update Parts Room**: Update inventory information
- **Take a Part Out**: Remove parts from inventory
- **Search for a Part**: Open search interface
- **Request a Part to be Ordered**: Submit parts order request

**Work Management:**
- **Request a Work Order**: Create new work order request
- **Make an MTSC Ticket**: Opens MTSC ticket interface
- **Search Knowledge Base**: Search maintenance knowledge base

**Parts Room Management:**
- **Add Same Day Parts Room**: Add new same-day parts location
- **Add 1-Day Parts Room**: Set up 1-day parts location (future)
- **Add 2-Day Parts Room**: Set up 2-day parts location (future)

### Search Tab

#### Features:
- **Comprehensive Search**: Search by multiple criteria
- **Results Display**: Shows availability across all locations
- **Part Selection**: Select and add parts to work orders

#### Search Options:
1. **NSN Search**
   - Enter full NSN or partial with wildcard (*)
   - Example: "1234*" finds all NSNs starting with 1234

2. **OEM/Part Number Search**
   - Enter manufacturer part number
   - Supports wildcards
   - Searches across all OEM fields

3. **Description Search**
   - Text-based search of part descriptions
   - Supports partial matches

#### Results Sections:
- **Availability**: Shows local parts room inventory
- **Same Day Parts**: Shows parts available at same-day locations
- **Cross Reference**: Shows related parts and references

## Notes
- Always ensure you run the scripts by right-clicking and selecting "Run with PowerShell"
- Keep your CSV files up to date in the Dropdown CSVs folder
- Back up your data regularly
- Contact your system administrator if you encounter any issues

## Troubleshooting
- If the application doesn't start, verify that all required directories and files exist
- Check the UI.log file for error messages
- Ensure PowerShell execution policy allows running the scripts
- Verify all CSV files are properly formatted

## Additional Generated Files

### CallLogs.csv Format:
```
Date,Machine,Cause,Action,Noun,Time Down,Time Up,Notes
2024-01-01,DBCS-123,Broken,Replace,Belt,08:00,09:30,Regular maintenance
```

### LaborLogs.csv Format:
```
Date,Work Order,Description,Machine,Duration,Parts,Notes
2024-01-01,WO-123,Replaced drive belt,DBCS-123,1.5,NSN12345 (1) - Local,Routine replacement
```

### Parts Room CSV Format:
```
Part (NSN),Description,QTY,13 Period Usage,Location,OEM 1,OEM 2,OEM 3,Changed Part (NSN)
12345,Drive Belt,10,5,A1-B2,MFG123,MFG456,MFG789,67890
```
# Known Issues and Limitations

## Search Functionality Issues

### Search Cross-Contamination
There is a critical issue with the search functionality:
- Using the search in the Labor Log tab interferes with the Search tab's accuracy
- After using Labor Log search, the Search tab will return incorrect or incomplete results
- **Required Workaround**: Close and restart the software between using these different search functions
- Best Practice: Complete all Search tab operations before using Labor Log search, or vice versa

### Search Tab Best Practices
To avoid search issues:
1. Decide which search function you need to use first
2. Complete all searches in that tab
3. Close the software
4. Reopen the software if you need to use the other search function
5. Remember that switching between search functions without a restart will compromise search accuracy

## Labor Log Tab Issues

### Tooltip Display Bug
When hovering over entries in the Labor Log tab, you may experience:
- Error messages appearing instead of the expected tooltip
- The tooltip should show detailed parts information when hovering over entries
- This is a known issue awaiting resolution

### Notification Icon Not Working
- The red notification dot that should appear for unacknowledged entries is currently non-functional
- This feature is designed to show when work orders need numbers assigned
- The underlying tracking system works, but the visual indicator does not display properly

### Parts Persistence Issue
When adding parts to work orders:
- Parts information will be visible in the current session
- However, this information does not persist after closing and reopening the software
- You'll need to re-add parts information after each software restart
- This is a known limitation pending future updates

## Dynamic Update Issues

The software requires a restart to recognize certain changes:
- Newly added parts books
- New parts rooms
- Updated same-day parts locations
- Any modifications to the directory structure
- Switching between search functions
- Adding parts to work orders

**Workaround**: After making any of these changes, close and reopen the software to see the updates.

## Best Practices to Avoid Issues

1. **Parts Management**:
   - Keep detailed records outside the software as backup
   - Document parts additions in a separate system until persistence is fixed
   - Consider taking screenshots of parts assignments for record-keeping

2. **Work Order Management**:
   - Assign work order numbers promptly
   - Don't rely on the notification system
   - Regularly review entries for missing work order numbers

3. **Software Updates**:
   - Close and reopen the software after making structural changes
   - Verify changes are visible after restart
   - Maintain backups of all CSV files

4. **Search Operations**:
   - Plan search operations to minimize software restarts
   - Complete all searches in one tab before switching
   - Document search results before closing software
   - Verify search results match expected outcomes
