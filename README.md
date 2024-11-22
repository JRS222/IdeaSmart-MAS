# Comprehensive Maintenance Software Guide

## Overview
This maintenance software consists of two PowerShell scripts:
1. Parts-Books-Creator.ps1 (Setup Script)
2. UI.ps1 (Main Application)

## Initial Setup

### Step 1: Run the Setup Script
1. Right-click on `Parts-Books-Creator.ps1` and select "Run with PowerShell"
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
      ├── UI.ps1
      └── Parts-Books-Creator.ps1
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

#### How to Add Parts to Work Order:
1. Select a labor log entry
2. Click "Add Parts to Work Order"
3. Search for parts using:
   - NSN
   - Part Number
   - Description
4. Select desired parts and quantities
5. Click "Attach to W/O" to save

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

This comprehensive format ensures all necessary information is captured and properly organized within the system.
