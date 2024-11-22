
---

# Advanced Parts Management Software: Revolutionizing Equipment Maintenance Efficiency

## Introduction

In equipment maintenance, where every second of downtime can be costly, efficiency and accuracy are paramount. Technicians need to identify, locate, and manage parts quickly and precisely. Our **Advanced Parts Management Software** does just that—streamlining parts and equipment management to enable faster, smarter work. This comprehensive software provides centralized data, automated processes, and user-friendly tools, empowering your team to tackle maintenance challenges with unparalleled efficiency.

---

## **1. Parts Books Tab**

### Purpose

The **Parts Books Tab** serves as the central repository for accessing detailed parts books tailored to each machine. Each book includes essential columns such as stock numbers, descriptions, and reference figures. This data is organized in a way that streamlines parts identification and retrieval, ensuring technicians can quickly find what they need without sifting through multiple resources.

![[Pasted image 20241112112945.png]]
_FIGURE 1.1 Parts Books Tab_

![[Pasted image 20241112214752.png]]
_Figure 1.2 Parts Book Sample (AFCS200)_
![[Pasted image 20241113064159.png]]
_Figure 1.3 Parts Room Sample_

### Distinctions Between Parts Books and Parts Room Files

While the **Parts Room** file consolidates information across all parts books for a holistic view, each **Parts Book** file focuses on the unique data relevant to a specific machine. This targeted structure provides efficiency and precision when working on individual machines.

### Column Breakdown in Parts Book Files

#### Key Columns

- **NO.**: This column lists the part number within the specific figure referenced in the **REF.** column. It allows technicians to directly correlate the part with its illustration in the handbook, making it easy to identify within diagrams.

```mermaid
flowchart TD
    A[Start] --> B[Parse Figure HTML]
    B --> C{Item Number Present?}
    C -->|Yes| D[Extract Item Number]
    C -->|No| E[Generate Sequential Number]
    D --> F{Valid Format?}
    F -->|Yes| G[Store in NO. Column]
    F -->|No| H[Format Correction]
    E --> G
    H --> G
    G --> I[End]
```
_Figure 1.4 NO. Column Logic_

- **STOCK NO.**: Displays the standardized **NSN** or **NSL** (Not Stocked Locally) label. NSNs allow cross-referencing with the national stockroom, while NSLs identify parts that may not be immediately available but can be sourced or ordered as needed.


```mermaid
flowchart TD
    A[Start] --> B[Parse Figure HTML]
    B --> C{Stock Number Present?}
    C -->|Yes| D[Extract Stock Number]
    C -->|No| E{Is NSL Part?}
    D --> F{Format Check}
    E -->|Yes| G[Mark as NSL]
    E -->|No| H[Research Required]
    F -->|Valid| I[Store in STOCK NO. Column]
    F -->|Invalid| J[Format Correction]
    G --> I
    H --> K[Manual Entry]
    J --> I
    K --> I
    I --> L[End]
```
_Figure 1.5 NSN Column Logic_


- **PART DESCRIPTION**: Contains detailed information about the part, such as specifications (e.g., "Screw, cap, hexagon head, steel, zinc coated, class 8.8, M8 x 1.25 x 30 mm long"). This ensures that technicians can verify they are selecting the correct part without ambiguity.

```mermaid
flowchart TD
    A[Start] --> B[Parse Figure HTML]
    B --> C{Description Present?}
    C -->|Yes| D[Extract Description]
    C -->|No| E[Research Required]
    D --> F{Format Check}
    F -->|Valid| G[Clean Special Characters]
    F -->|Invalid| H[Format Correction]
    E --> I[Manual Entry]
    G --> J[Store in PART DESCRIPTION Column]
    H --> J
    I --> J
    J --> K{Length Check}
    K -->|Too Long| L[Truncate]
    K -->|OK| M[End]
    L --> M
```
_Figure 1.6 Description Column Logic_

- **PART NO.**: This column lists the OEM information. It allows easy cross-referencing with suppliers and manufacturers.

- **REF.**: Specifies the figure references (e.g., "Figure 2-10") within the handbook where the part can be found. This column is essential for directing technicians to the exact figure that includes the part, saving time and reducing errors.

```mermaid
flowchart TD
    A[Start] --> B[Get Figure Number]
    B --> C{Valid Format?}
    C -->|Yes| D[Extract Section-Figure]
    C -->|No| E[Format Correction]
    D --> F{HTML File Exists?}
    F -->|Yes| G[Create Hyperlink]
    F -->|No| H[Store Text Only]
    E --> I[Manual Review]
    G --> J[Store in REF. Column]
    H --> J
    I --> J
    J --> K[End]
```
_Figure 1.7 Reference Column Logic_

- **QTY**: Indicates the quantity on hand or required, streamlining inventory planning and ensuring that parts are in stock or requested in advance.

- **LOCATION**: Shows the storage location or notes that a part is not stocked locally, aiding in quick retrieval or sourcing decisions.

```mermaid
flowchart TD
    A[Start Cell Population] --> B[Load Parts Room CSV]
    B --> C{Process Each Row in Section}
    
    C --> D{Match by STOCK NO.?}
    D -->|Yes| E[Copy Location from Parts Room]
    D -->|No| F{Match by PART NO.?}
    
    F -->|Yes| G{Multiple OEM Matches?}
    F -->|No| H[Set Not Stocked Locally]
    
    G -->|Yes| I[Use First Match Location]
    G -->|No| J[Use Single Match Location]
    
    E --> K[Write to Location Column]
    H --> K
    I --> K
    J --> K
    
    K --> L{More Rows?}
    L -->|Yes| C
    L -->|No| M[End]

    style H fill:#f9f,stroke:#333
    style E fill:#9f9,stroke:#333
```
_Figure 1.8 Location Column Logic_

- **CAGE**: Provides the Commercial and Government Entity (CAGE) code, essential for identifying the part's manufacturer or supplier.



### Column Breakdown in Parts Room Files

#### Key Columns

- **PART NO.**: Identifies the unique part number associated with each item. This enables quick cross-referencing between parts books and inventory files to ensure the correct part is available.

- **DESCRIPTION**: Details the part characteristics, such as type and specifications. It’s crucial for ensuring technicians have the necessary information for parts selection.

- **STOCK NO.**: Lists the National Stock Number (NSN) if applicable. This column is essential for inventory verification and identifying parts that align with national standards.
```mermaid
flowchart TD
    A[Start] --> B{Is part in EMARS?}
    B -->|Yes| C[Extract NSN from EMARS]
    B -->|No| D{Is it a standard part?}
    D -->|Yes| E[Use Standard NSN]
    D -->|No| F[Mark as NSL]
    C --> G[Store in Part NSN Column]
    E --> G
    F --> G
    G --> H{Changed NSN?}
    H -->|Yes| I[Move old NSN to 'Changed Part NSN' column]
    H -->|No| J[End]
    I --> J
```
_Figure 1.9 STOCK NO. Column Logic_

- **QTY**: Displays the quantity on hand, facilitating inventory tracking and restocking when necessary.

- **LOCATION**: Indicates where the part is stored within the facility or notes that it is not locally stocked, aiding in efficient retrieval and stock tracking.

- **CAGE**: Lists the Commercial and Government Entity (CAGE) code for the part’s supplier, essential for sourcing and procurement.

- **New Parts Book Reference Columns**: Columns that allow a user to quickly cross-reference between the Parts Volume and his Site. Particularly useful when you have a part you need to replace and the only part number you can find is on a part nearby. 

```mermaid
flowchart TD
    A[Start] --> B[Process New Parts Book]
    B --> C[Create New Column for Book]
    C --> D{For Each Part in Book}
    D --> E{Match with Parts Room?}
    
    E -->|NSN Match| F[Get All Figures for Part]
    E -->|OEM Match| G[Check Changed NSN]
    E -->|No Match| H[Skip Part]
    
    F --> I[Format as Figure References]
    G --> J{Valid NSN Change?}
    J -->|Yes| K[Update Part NSN]
    J -->|No| L[Mark Changed NSN]
    
    K --> I
    L --> I
    
    I --> M{Existing References?}
    M -->|Yes| N[Append New References]
    M -->|No| O[Create New Reference]
    
    N --> P[Format References]
    O --> P
    
    P --> Q{More Parts?}
    Q -->|Yes| D
    Q -->|No| R[Save Updates]
    
    H --> Q
    
    R --> S[Update Excel File]
    S --> T[End]
    
    P --> U[Sort Figure Numbers]
    U --> V[Group by Section]
    V --> W[Separate by Pipe]
    W --> X[Format: X-Y pipe A-B pipe C-D]
```
_Figure 1.10 Book Reference Column Logic_

---
### Dynamic Integration with the Parts Room File

When a new parts book is created, its name is added as a column in the **Parts Room** file. This enables technicians to see, at a glance, which parts appear across different machines and specific figures within each book. For example, if a part is featured in multiple figures within a book, these figures (e.g., "Figure 2-11 | Figure 2-8") are populated in the respective book-specific column. This helps technicians focus on relevant sections, saving valuable time during diagnostics and repair.

### Benefits

- **Targeted Information Access**: Enables technicians to quickly locate machine-specific data within each parts book.
  
- **Enhanced Cross-Referencing**: With book-specific columns in the **Parts Room**, technicians can see exactly which figures to consult, improving efficiency and accuracy.

- **Simplified Navigation**: Figure references in the **Parts Book** files link directly to the parts diagrams, minimizing the time spent navigating complex manuals.

---

## **2. Call Logs Tab**

### Purpose

The **Call Logs Tab** provides structured access to service calls, ensuring efficient maintenance tracking. It logs essential call details, updates statuses, and automates labor transfers, ensuring that crucial data is always available for analysis.

### Key Features

- **Recent Calls List**: Displays machine identifiers, causes, actions, timestamps, and notes, keeping the team informed about issues and resolutions in real time.
  
- **Automatic Labor Log Transfer**: Moves calls over 30 minutes into the **Labor Log Tab** for more detailed tracking. This automation ensures that significant maintenance activities are captured without extra steps.
  
- **End-of-Shift Reporting**: Simplifies the reporting process by consolidating all calls and activities at the end of each shift. This feature encourages technicians to accurately document calls, enhancing long-term data reliability.

![[Pasted image 20241114045740.png]]
_Figure 2.1 Call Logs Tab_
### Benefits

- **Improved Communication**: Keeps team members updated on equipment status, preventing miscommunications.
  
- **Efficient Documentation**: Automates logging for convenience and accuracy.
  
- **Enhanced Accountability**: Tracks every action, supporting performance assessments.

![[Pasted image 20241114045932.png]]
_Figure 2.2  Call Log Adding Window_

![[Pasted image 20241114050252.png]]
_Figure 2.3 Adding a new machine to the options for calls_

---

## **3. Labor Log Tab**

### Purpose

The **Labor Log Tab** is a powerful tool for creating work orders, tracking parts usage, and enabling detailed records. It includes search capabilities designed for parts retrieval and is integrated with EMARS for seamless work order management. Any Call Log that is over 30 minutes gets moved to the Labor Log view as well and tells the user that they need a work order number. This also allows them to utilize the adding of parts.

![[Pasted image 20241115030201.png]]
_Figure 3.1 Labor Log Tab_
### Key Features

- **Create Work Orders**: With fields for the **Date**, **Work Order**, **Description**, **Machine**, **Duration**, **Parts**, and **Notes** work orders are auto-populated with the parts database. Technicians can then add work-specific details and seamlessly integrate with EMARS.

![[Pasted image 20241115030401.png]]
_Figure 3.2 Adding a labor log_

- **Search for Parts**: The built-in parts search saves technicians time by enabling quick retrieval from the current parts database.

![[Pasted image 20241115030646.png]]
_Figure 3.3 Searching for a part_

![[Pasted image 20241115030818.png]]
_FIgure 3.4 Adding a part to a work order_

![[Pasted image 20241115030954.png]] 
_Figure 3.5 Prompt to add part. This will come up for each part._

![[Pasted image 20241115031155.png]]
_Figure 3.6 Once you have selected you parts it populates_

![[Pasted image 20241115031436.png]]
_Figure 3.7 How it populates the parts it goes_ <NSN><(quantity)>-<Location>, <NSN><(quantity)>-<Location>

```mermaid
flowchart TB
    subgraph Initialize[Startup Process]
        IConfig[Initialize Configuration] -->|Load Paths| Paths[Get File Paths]
        Paths -->|Check Files| Files[Load Call/Labor Files]
        Files -->|Initialize| Views[Create Call/Labor Views]
    end

    subgraph CallLog[Call Log Data Loading]
        Import[Import CallLogs.csv] -->|Read Entries| CallData[Create Call Log Objects]
        CallData -->|For Each Entry| Process[Process Each Call]
        Process -->|Check Duration| Qualify{Qualifies for Labor Log?}
        Qualify -->|No: < 30min| Skip[Skip Entry]
        Qualify -->|Yes: >= 30min| UniqueCheck
    end

    subgraph Validation[Entry Validation]
        UniqueCheck{Check If Unique} -->|Create Key| Key["Create uniqueKey:
        Date_Machine_TimeDown"]
        Key --> Hash[Check processedCallLogs Hash]
        Hash --> Exists{Entry Exists?}
        Exists -->|Yes| SkipDupe[Skip Duplicate]
        Exists -->|No| CreateEntry[Create Labor Entry]
    end

    subgraph LaborEntry[Labor Log Entry Creation]
        CreateEntry --> Duration[Calculate Duration]
        Duration -->|"timeUp - timeDown
        / 60 (hours)"| DurFormat[Format Duration]
        
        DurFormat --> BuildEntry[Build Labor Entry]

        subgraph Mapping[Column Mapping]
            direction TB
            BuildEntry --> C1[Date: log.Date]
            BuildEntry --> C2["Work Order: 'Need W/O #'"]
            BuildEntry --> C3["Description: 
            Cause/Action/Noun Combined"]
            BuildEntry --> C4[Machine: log.Machine]
            BuildEntry --> C5["Duration: 
            Calculated Hours"]
            BuildEntry --> C6["Parts: Empty"]
            BuildEntry --> C7[Notes: log.Notes]
        end
    end

    subgraph Save[Save Process]
        BuildEntry --> AddItem[Add to Labor ListView]
        AddItem --> SaveCSV[Save to LaborLogs.csv]
        SaveCSV --> UpdateHash[Update processedCallLogs]
    end

    Save --> Final[Process Next Entry]
    Final --> Process

    style Mapping fill:#e1f5fe,stroke:#b3e5fc
    style Initialize fill:#e8f5e9,stroke:#c8e6c9
    style Validation fill:#fff3e0,stroke:#ffe0b2
```
_Figure 3.8 How call logs are moved and how columns are populated_

The most important parts are color-coded:
- Green: Initialization process
- Blue: Column mapping
- Orange: Validation process
  
  - **Enhance Reporting** with data for audits and budgeting.
### Benefits

- **Comprehensive Tracking**: Ensures complete, accurate labor logs.
  
- **Time Savings**: Simplifies work order creation and documentation.
  
- **Data Insights**: Offers historical data to improve maintenance planning.

---

## **4. Actions Tab**

### Purpose

The **Actions Tab** serves as a **temporary hub** for actions that have not yet been assigned to specific tabs but remain integral to the workflow. It also includes features that connect and streamline various components of the program. This ensures ongoing functionality while providing a centralized workspace for fast, automated tasks.

![[Pasted image 20241115033215.png]]
_Figure 4.1 Quick look at the Actions tab_

### Key Features
- **Workflow Bridge**: Facilitates steps that involve multiple features of the software, tying together interconnected processes in a single interface.
---

## **5. Search Tab**

### Purpose

The **Search Tab** offers flexible, fast search capabilities for technicians. Parts can be located using NSNs, OEM numbers, or descriptions, significantly speeding up parts retrieval.

### Key Features

- **Wildcard Search**: Enables flexible search terms with wildcards, allowing technicians to find parts even with incomplete data.
  
- **Cross-Reference Capabilities**: Cross-references NSNs with OEM numbers and descriptions, providing all possible options.
  
- **Multi-Location Search**: Extends searches to local and national stockrooms and nearby facilities, improving the chances of finding parts quickly.
  
- **Cross-Referenced Figures**: For parts in multiple handbooks, results show relevant figures and allow quick access. This feature reduces manual handbook navigation drastically.

### How it works:

# Search Pattern Reference

## Cross Reference Search Specifics
| Field | Search Target | Notes |
|-------|--------------|-------|
| NSN | STOCK NO. | Matches against Parts Books' STOCK NO. field |
| Part No | PART NO. | Matches against Parts Books' PART NO. field |
| Description | PART DESCRIPTION | Matches against Parts Books' PART DESCRIPTION field |

## Cross Reference Display Fields
| Field | Source | Description |
|-------|--------|-------------|
| REF | Section reference | From Parts Book section reference |
| STOCK NO. | Parts Book data | NSN from Parts Book |
| PART NO. | Parts Book data | Manufacturer part number |
| Location | Parts Book data | Location if available |
| QTY | Not applicable | Cross reference doesn't track quantity |

## Cross Reference Search Behavior
- Searches all sections in all configured Parts Books
- REF numbers preserved for figure lookup
- Matches must meet all provided search criteria (AND logic)
- Empty search fields match all entries in that field
- Search is performed across all configured Parts Books simultaneously
- Results include book and section information for reference

## View-Specific Search Comparisons
| Feature | Main Parts Room | Same Day Parts | Cross Reference |
|---------|----------------|----------------|-----------------|
| Source Display | CSV Name | Site Name | Book Name |
| QTY Available | Yes | Yes | No |
| Location Info | Current | Current | From Manual |
| Part Number Source | OEM Fields | OEM Fields | PART NO. Field |
| NSN Format | Current | Current | Manual Format |
| Results Update | Real-time | Real-time | Static |

## Search Field Interactions Across Views
| Search Type | Main Parts Room | Same Day Parts | Cross Reference |
|------------|-----------------|----------------|-----------------|
| NSN Only | Searches Part (NSN) | Searches Part (NSN) | Searches STOCK NO. |
| Part No Only | Searches OEM fields | Searches OEM fields | Searches PART NO. |
| Description Only | Searches Description | Searches Description | Searches PART DESCRIPTION |
| Combined | AND logic between fields | AND logic between fields | AND logic between fields |

## Important Notes:
1. All three views are searched simultaneously
2. Each view maintains its own result set
3. Cross Reference results don't include quantity information
4. Location information means different things in different views
5. Part numbers are matched differently in Cross Reference vs. other views
6. Results are displayed in separate sections of the interface
7. Search terms are applied consistently across all views but may match different fields

```mermaid
flowchart TB
    subgraph Input[Search Input Fields]
        NSN[NSN Search Box]
        PartNo[Part Number Search Box]
        Desc[Description Search Box]
    end

    subgraph MainPartsRoom[Main Parts Room Search]
        MPR[Main Parts Room CSV] -->|For Each Row| MPRFilter{Filter Criteria}
        MPRFilter -->|NSN Match| MPRNSN{NSN Check}
        MPRFilter -->|Part No Match| MPRPN{Part No Check}
        MPRFilter -->|Description Match| MPRDesc{Desc Check}
        
        MPRNSN -->|Yes| MPRNext[Next Filter]
        MPRNSN -->|No| MPRSkip[Skip Row]
        
        MPRPN -->|Match Any OEM Field| MPRNext
        MPRPN -->|No Match| MPRSkip
        
        MPRDesc -->|Yes| MPRAdd[Add to Results]
        MPRDesc -->|No| MPRSkip

        MPRAdd -->|Display Fields| MPRShow["Show:
        - Source (CSV Name)
        - STOCK NO.
        - PART NO.
        - Location
        - QTY"]
    end

    subgraph SameDay[Same Day Parts Room Search]
        SDR[Same Day Room CSVs] -->|For Each CSV| SDSite[Process Each Site]
        SDSite -->|For Each Row| SDFilter{Filter Criteria}
        SDFilter -->|NSN Match| SDNSN{NSN Check}
        SDFilter -->|Part No Match| SDPN{Part No Check}
        SDFilter -->|Description Match| SDDesc{Desc Check}
        
        SDNSN -->|Yes| SDNext[Next Filter]
        SDNSN -->|No| SDSkip[Skip Row]
        
        SDPN -->|Match Any OEM Field| SDNext
        SDPN -->|No Match| SDSkip
        
        SDDesc -->|Yes| SDAdd[Add to Results]
        SDDesc -->|No| SDSkip

        SDAdd -->|Display Fields| SDShow["Show:
        - Source (Site Name)
        - STOCK NO.
        - PART NO.
        - Location
        - QTY"]
    end

    subgraph CrossRef[Cross Reference Search]
        direction TB
        CR[Parts Books] -->|For Each Book| CRBook[Process Each Book]
        CRBook -->|Load Sections| CRSection[Process Each Section]
        CRSection -->|For Each Row| CRFilter{Filter Criteria}
        
        CRFilter -->|NSN Match| CRNSN{Check STOCK NO.}
        CRFilter -->|Part No Match| CRPN{Check PART NO.}
        CRFilter -->|Description Match| CRDesc{Check PART DESCRIPTION}
        
        CRNSN -->|Yes| CRNext[Continue to Next Check]
        CRNSN -->|No| CRSkip[Skip Entry]
        
        CRPN -->|Yes| CRNext
        CRPN -->|No| CRSkip
        
        CRDesc -->|Yes| CRAdd[Add to Results]
        CRDesc -->|No| CRSkip

        CRAdd -->|Display Fields| CRShow["Show:
        - Handbook
        - REF.
        - STOCK NO.
        - PART NO.
        - Location
        - QTY"]
    end

    Input -->|Search Criteria| MainPartsRoom
    Input -->|Search Criteria| SameDay
    Input -->|Search Criteria| CrossRef

    style Input fill:#e1f5fe
    style MainPartsRoom fill:#e8f5e9
    style SameDay fill:#fff3e0
    style CrossRef fill:#f3e5f5

    subgraph SearchNotes[Search Processing Notes]
        direction TB
        Note1["• All searches case-insensitive"]
        Note2["• Wildcards (*) allowed in all fields"]
        Note3["• Empty fields match everything"]
        Note4["• NSN search ignores hyphens"]
        Note5["• Part No matches any OEM field"]
        Note6["• Fields combined with AND logic"]
    end
```
![[Pasted image 20241115071149.png]]
#### _OEM Search Example_

![[Pasted image 20241115071608.png]]
#### _NSN Search Example_

![[Pasted image 20241115072114.png]]
#### _Description Search Example_

![[Pasted image 20241121220445.png]]
![[Pasted image 20241121221124.png]]

By selecting any number of parts and clicking "Open Figures" it will open directly to the drawing. This is much, much faster than navigating slowly through the website. I can put in a partial number or a number of a nearby part and open, instantly, any drawing I suspect that could be helpful.
```mermaid
flowchart TB
    subgraph UserAction[User Interaction]
        Start([Start]) --> Button[Click 'Open Figures' Button]
        Button --> GetRef[Get Reference from Row]
        GetRef --> ParseRef[Parse Figure Reference]
    end

    subgraph ReferenceProcessing[Reference Processing]
        ParseRef --> Split[Split Multiple References]
        Split --> Valid{Valid References?}
        Valid -->|No| Error[Show Error Message]
        Valid -->|Yes| Process[Process Each Reference]
    end

    subgraph FileHandling[File Handling]
        Process --> BuildPath[Build File Path]
        BuildPath --> CheckFile{File Exists?}
        CheckFile -->|No| NotFound[Show Not Found Error]
        CheckFile -->|Yes| LoadHTML[Load HTML File]
    end

    subgraph Display[Display Processing]
        LoadHTML --> ExtractFig[Extract Figure Content]
        ExtractFig --> CreateWindow[Create New Window]
        CreateWindow --> RenderFig[Render Figure]
        RenderFig --> AddControls[Add Navigation Controls]
    end

    subgraph Navigation[Window Controls]
        AddControls --> Zoom[Add Zoom Controls]
        AddControls --> Print[Add Print Option]
        AddControls --> Close[Add Close Button]
    end

    Error --> End([End])
    NotFound --> End
    Navigation --> End

    classDef process fill:#e1f5fe,stroke:#b3e5fc
    classDef decision fill:#fff3e0,stroke:#ffe0b2
    classDef action fill:#e8f5e9,stroke:#c8e6c9
    
    class UserAction,Display action
    class ReferenceProcessing process
    class Valid,CheckFile decision
    class FileHandling process
    class Navigation action
```

### Benefits

- **Speed**: Reduces time spent finding parts by over 90%.
  
- **Accuracy**: Increases accuracy with detailed, cross-referenced results.
  
- **Convenience**: Combines all search functions into one interface.

---

## **6. Optimized Workflow for Parts Search**

### The Transformation

Our software redefines the time-intensive parts identification workflow by providing a seamless, end-to-end solution.

#### Old Workflow Process (50–92 minutes per task):

1. **Identify Issue on Machine**: Inspect machine to determine the issue.
   
2. **Access ACE Computer**: Locate and use a dedicated ACE computer to access parts and repair information.
   
3. **Navigate to MTSC Website**: Find machine acronym link and relevant handbook.
   
4. **Identify Handbook and Section**: Manually locate correct section.
   
5. **Identify the Part**: Scroll through sections to locate specific part information.
   
6. **Locate Part in National Stockroom Database**: Search by NSN, OEM, or description.
   
7. **Check Other Facility Part Rooms** if unavailable locally.
   
8. **Order Part**: Submit a request to parts clerk.
   
9. **Document Call Log**: Record the issue, actions taken, and time manually.

#### New Workflow with Advanced Parts Management Software (Under 10 minutes per task):

1. **Diagnose Issue**: Use historical data from the **Call Logs Tab**.
   
2. **Search for Part**: Instantly search by any known identifier.
   
3. **Access Detailed Part Information**: View parts data with cross-referenced figures in the **Parts Books Tab**.
   
4. **Check Stock Availability**: See real-time stock levels and locations in the **Parts Room**.
   
5. **Order Part**: Use the **Actions Tab** to order the part immediately if out of stock.
   
6. **Automated Logging**: Automatically updates **Call Logs**

 and **Labor Logs**.

```mermaid
flowchart TB
    subgraph Old["Old Workflow (50-92 minutes)"]
        direction TB
        A1[1. Identify Issue on Machine] -->|5-10 min| A2[2. Access ACE Computer]
        A2 -->|5-10 min| A3[3. Navigate MTSC Website]
        A3 -->|10-15 min| A4[4. Identify Handbook and Section]
        A4 -->|10-20 min| A5[5. Identify Part]
        A5 -->|5-15 min| A6[6. Search National Stockroom]
        A6 -->|5-10 min| A7[7. Check Other Facilities]
        A7 -->|5-7 min| A8[8. Order Part]
        A8 -->|5 min| A9[9. Document Call Log]
    end

    subgraph New["New Workflow (<10 minutes)"]
        direction TB
        B1[1. Diagnose Issue] -->|1-2 min| B2[2. Search for Part]
        B2 -->|1-2 min| B3[3. Access Part Information]
        B3 -->|1-2 min| B4[4. Check Stock Availability]
        B4 -->|1-2 min| B5[5. Order Part]
        B5 -->|Auto| B6[6. Automated Logging]
    end

    classDef oldStyle fill:#ffebee,stroke:#ef9a9a
    classDef newStyle fill:#e8f5e9,stroke:#a5d6a7
    
    class A1,A2,A3,A4,A5,A6,A7,A8,A9 oldStyle
    class B1,B2,B3,B4,B5,B6 newStyle

    %% Add time savings visualization
    TimeComparison["Total Time Savings:
    Old Process: 50-92 minutes
    New Process: <10 minutes
    Efficiency Gain: >80%"]

    style TimeComparison fill:#e3f2fd,stroke:#90caf9
```

### Time Saved

- **Old Workflow**: 50-92 minutes per task.
  
- **New Workflow**: Under 10 minutes per task

---

## **7. Windows PowerShell Integration**

### Purpose

The software leverages Windows PowerShell scripts to automate complex tasks, including data imports, HTML processing, and Excel report generation, reducing manual intervention and enhancing accuracy.

### Key Features

- **Automated Data Handling**: Scripts streamline data extraction and updating, ensuring the latest data is always accessible.
  
- **Dynamic Parts Books Generation**: Scripts automate Excel creation for each new book, keeping data current with minimal manual input.
  
- **Error Handling and Notifications**: Notifications immediately inform users of any issues, supporting a smooth experience.

*Template location*: *Include a diagram showing how PowerShell scripts automate data flow.*

### Benefits

- **Improved Accuracy**: Reduces the risk of data entry errors.
  
- **Efficiency**: Automates data handling, freeing staff to focus on essential tasks.
  
- **Scalability**: Allows the system to grow with ease, adapting to new machines, parts, or handbooks as needed.

---

## **Conclusion**

Our **Maintenance Assistance Software** is a transformative solution that not only manages parts but optimizes your entire maintenance workflow. Each feature is designed with technicians’ needs in mind—from the **Parts Books Tab** that centralizes all parts data, to the intelligent **Search Tab** that significantly reduces search times. Future enhancements, such as the planned historical labor search, will provide even deeper insights for long-term analysis and reporting.

Invest in this system to reduce downtime, cut costs, and improve operational efficiency—taking your maintenance operations to new heights.

---