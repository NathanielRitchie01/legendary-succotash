Clear-Host
$Host.UI.RawUI.WindowTitle = "NR WCS Performance Dashboard - AUKC01"

# Force TLS 1.2 or higher (TLS 1.0 and 1.1 are deprecated)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13


# Life hack from Nathaniel, stuff up function scopes and so now create a work around in the shape of a global variabel
$Global:currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Global:currentUOM = "EACHES" # Default UOM, can be changed by user selection in functions that require it. Will be used for KPI calculations and display.

function futureupdate{
    # Later for moduleCheck lol
    Install-Module Graphical -scope CurrentUser
    Import-Module Graphical
}


#Function to go through correct order of start up processes - only because I not used to script style

#Helper Compare Functions

function isValidSysSelectionEnum{
    param ([int]$userSelection)
    return $userSelection -ne -1
}

# Global ENUM do not touch please - all SQL is based upon this for selection of UOM.
enum UOM {
    Eaches = 1
    Totes = 2
    Cartons = 3
}

enum DECANT_UOM{
    EACHES = 1
    TOTES = 2
    CARTONS = 3
}

enum GTP_UOM {
    EACHES = 1
    #Totes = 2
    CARTONS = 3
}

enum HOURS_DAY {
    #ENUM - VALUE
    Hour00 = 0
    Hour01 = 1
    Hour02 = 2
    Hour03 = 3
    Hour04 = 4
    Hour05 = 5
    Hour06 = 6
    Hour07 = 7
    Hour08 = 8
    Hour09 = 9
    Hour10 = 10
    Hour11 = 11
    Hour12 = 12
    Hour13 = 13
    Hour14 = 14
    Hour15 = 15
    Hour16 = 16
    Hour17 = 17
    Hour18 = 18
    Hour19 = 19
    Hour20 = 20
    Hour21 = 21
    Hour22 = 22
    Hour23 = 23
}


enum PICK_STATE{
    #ENUM with Default Value
    CREATING
    EXPECTED
    PENDING
    WAIT_ALLOCATION
    UNSATISFIABLE
    ALLOCATED
    UNPICKABLE
    WAIT_STOCK
    RESERVED
    STARTED
    PICKED
    COLLATING
    COLLATED
    PACKABLE
    PACKING
    PACKED
    BUFFERED
    UNREACHABLE
    MARSHALLING
    MARSHALLED
    LOADING
    LOADED
    DESPATCHED
    FINISHED
    ABANDONED
    CANCELLED
}

enum PICK_STATE_DASHBOARD {
    <#These are the ones we care about mainly...#>
    PENDING 
    WAIT_ALLOCATION
    UNSATISFIABLE
    UNPICKABLE
    WAIT_STOCK
    RESERVED
    STARTED
    PICKED
    MARSHALLED
    LOADED
    DESPATCHED
    FINISHED
    ABANDONED
    CANCELLED
}
# SQL Server connection settings - Windows Authentication
$global:SQLServer   = "SQLDBAUP010"
$global:SQLDatabase = "prodmis"
$global:UseWindowsAuth = $true

# Check for SqlServer module and install if not present
function moduleCheck {

    if (-not (Get-Module -ListAvailable -Name SqlServer)) {
        try {
            Write-Host "SqlServer module not found. Installing..." -ForegroundColor Yellow
            Install-Module -Name SqlServer -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
            Write-Host "SqlServer module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install SqlServer module. Please install it manually." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "SqlServer module is already installed." -ForegroundColor Green
    }
}


<#

ALL MENU DISPLAYS SECTION START

#>

# Manin menu function that user sees at the start of everything
function MainMenu {
    while ($true) {
        Clear-Host

        Write-Host ""
        Write-Host "       ______________________________________________________________"
        Write-Host ""
        Write-Host "                 Logs Methods:"
        Write-Host ""
        Write-Host "             [1] Fill Percentage                          - AUKC01"
        Write-Host "             [2] Consumable Usage (WIP)                   - AUKC01"
        Write-Host "             [3] Operation KPIs Menu                      - AUKC01"
        Write-Host "             [4] Picks By Day                             - AUKC01"
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [5] Random Selection 5"
        Write-Host "             [6] Random Selection 6"
        Write-Host "             [7] Random Selection 7"
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [8] Troubleshoot"
        Write-Host "             [E] Extras"
        Write-Host "             [H] Help"
        Write-Host "             [0] Exit"
        Write-Host "       ______________________________________________________________"
        Write-Host ""

        $choice = Read-Host "Choose a menu option"

        switch ($choice.ToUpper()) {
            "1" { FillPercentage }
            "2" { ConsumableUsage }
            "3" { KPIsMenu }
            "4" { EachesPerDay }
            "5" { Write-Host "Option 5 not implemented"; Pause }
            "6" { Write-Host "Option 6 not implemented"; Pause }
            "7" { Write-Host "Option 7 not implemented"; Pause }
            "8" { Troubleshoot }
            "E" { Extras }
            "H" { HelpMenu }
            "0" { exit }
            default {
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep 1
            }
        }
    }
}

# Fill percentage does not have a menu unfortuantly?

# KPI menu option when user selects option 3
function KPIsMenu {
    while ($true) {
        Clear-Host
        Write-Host ""
        Write-Host "       ______________________________________________________________"
        Write-Host ""
        Write-Host "                 KPIs Menu:"
        Write-Host ""
        Write-Host "             [1]  Decant Performance"
        Write-Host "             [2]  GTP Picking Performance"
        Write-Host "             [3]  Packing Performance"
        Write-Host "             [4]  Receiving Performance"
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [5]  Quality Metrics"
        Write-Host "             [6]  Cycle Count Accuracy"
        Write-Host "             [7]  Order Fulfillment Rate"
        Write-Host "             [8]  Putaway Efficiency"
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [9]  Inventory Turnover"
        Write-Host "             [10] Dock-to-Stock Time"
        Write-Host "             [11] Labor Productivity"
        Write-Host "             [12] Returns Processing"
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [B] Back to Main Menu"
        Write-Host "       ______________________________________________________________"
        Write-Host ""

        $choice = Read-Host "Choose a KPI option"

        switch ($choice.ToUpper()) {
            "1"  { DecantPerformance }
            "2"  { GTPPickingPerformance }
            "3"  { Write-Host "Packing Performance not yet implemented" -ForegroundColor Yellow; Pause }
            "4"  { Write-Host "Receiving Performance not yet implemented" -ForegroundColor Yellow; Pause }
            "5"  { Write-Host "Quality Metrics not yet implemented" -ForegroundColor Yellow; Pause }
            "6"  { Write-Host "Cycle Count Accuracy not yet implemented" -ForegroundColor Yellow; Pause }
            "7"  { Write-Host "Order Fulfillment Rate not yet implemented" -ForegroundColor Yellow; Pause }
            "8"  { Write-Host "Putaway Efficiency not yet implemented" -ForegroundColor Yellow; Pause }
            "9"  { Write-Host "Inventory Turnover not yet implemented" -ForegroundColor Yellow; Pause }
            "10" { Write-Host "Dock-to-Stock Time not yet implemented" -ForegroundColor Yellow; Pause }
            "11" { Write-Host "Labor Productivity not yet implemented" -ForegroundColor Yellow; Pause }
            "12" { Write-Host "Returns Processing not yet implemented" -ForegroundColor Yellow; Pause }
            "B"  { return }  # Goes back to Main Menu
            default {
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep 1
            }
        }
    }
}




<#

ALL MENU DISPLAYS SECTION END

#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#>
<#

QUERIES SECTION START

#>

function FillPercentageQuery {
    
    return @"
;WITH order_fill AS (
    SELECT
        order_id,
        AVG(fill_percent) AS order_avg_fill
    FROM x_du
    WHERE fill_percent IS NOT NULL
    GROUP BY order_id
)
SELECT
    o.num_lines,
    COUNT(*) AS order_count,
    AVG(ofl.order_avg_fill) AS avg_fill_percent
FROM order_fill AS ofl
JOIN x_order AS o
    ON ofl.order_id = o.order_id
GROUP BY o.num_lines
ORDER BY o.num_lines;
"@

}

function DecantPerformanceQuery {
    
    param   (
        [string]$targetDate = (queryDateUpdate), #(Get-Date -Format "yyyy-MM-dd"),  # Default to today's date

        # Will change to param for user input
        [DECANT_UOM]$uom = 2
        #1 is EACH
        #2 is TOTE
        #3 is CARTONS  # Assuming 1 represents Eaches, can be adjusted based on actual UOM values in the database
    )

    Write-Host "The current default selection: $uom" -ForegroundColor Green
    $userSelection = userUOMSelection -enumType ([DECANT_UOM])
    
    try {
        if (isValidSysSelectionEnum -$userSelection) {
            $uom = $userSelection
        }
    }
    catch {
        Write-Error "Bugger these enums..."
    }
   
    [string]$DECANT_QUERY_EACHES
    [string]$DECANT_QUERY_CARTONS
    [string]$DECANT_QUERY_TOTES
    $Global:currentUOM = [DECANT_UOM]$uom # Update global variable for display and KPI calculations. I know this is bad but it is late and I am tired. Will refactor later to remove global variable dependency.

    $DECANT_QUERY_EACHES = @"
            SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            SUM(quantity) AS [Eaches]
        FROM (
            SELECT
                quantity,
                change_uid,
                event_time,
                tote_id,
                DATEDIFF(SECOND, 
                    MIN(event_time) OVER (PARTITION BY tote_id, sku_id), 
                    event_time
                ) AS secs_from_first
            FROM mi_decant
            WHERE CAST(event_time AS DATE) = '$targetDate'
                AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
                AND change_uid IS NOT NULL
                AND quantity IS NOT NULL
        ) deduped
        WHERE secs_from_first = 0 OR secs_from_first > 300  -- 300 seconds = 5 minutes
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $DECANT_QUERY_CARTONS = @"
            SELECT 
                change_uid AS [User],
                DATEPART(HOUR, event_time) AS [Hour],
                COUNT(tote_id) AS [Eaches]
            FROM (
                SELECT 
                    change_uid,
                    event_time,
                    tote_id,
                    DATEDIFF(SECOND, 
                        MIN(event_time) OVER (PARTITION BY tote_id, sku_id), 
                        event_time
                    ) AS secs_from_first
                FROM mi_decant
                WHERE CAST(event_time AS DATE) = '$targetDate'
                    AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
                    AND change_uid IS NOT NULL
                    AND quantity IS NOT NULL
            ) deduped
            WHERE secs_from_first = 0 OR secs_from_first > 300  -- 300 seconds = 5 minutes
            GROUP BY change_uid, DATEPART(HOUR, event_time)
            ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $DECANT_QUERY_TOTES = @"
            SELECT 
                change_uid AS [User],
                DATEPART(HOUR, event_time) AS [Hour],
                COUNT(DISTINCT(case_id)) AS [Eaches]
            FROM (
                SELECT 
                    case_id,
                    change_uid,
                    event_time,
                    tote_id,
                    DATEDIFF(SECOND, 
                        MIN(event_time) OVER (PARTITION BY tote_id, sku_id), 
                        event_time
                    ) AS secs_from_first
                FROM mi_decant
                WHERE CAST(event_time AS DATE) = '$targetDate'
                    AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
                    AND change_uid IS NOT NULL
                    AND quantity IS NOT NULL
            ) deduped
            WHERE secs_from_first = 0 OR secs_from_first > 300  -- 300 seconds = 5 minutes
            GROUP BY change_uid, DATEPART(HOUR, event_time)
            ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    switch ([int]$uom) {
        1 { return $DECANT_QUERY_EACHES }
        2 { return $DECANT_QUERY_TOTES }
        3 { return $DECANT_QUERY_CARTONS }
        default { Write-Host "Invalid UOM selection. No idea how you could achieve this Defaulting to $uom." -ForegroundColor Yellow; return $GTP_QUERY_EACHES }
    }

    # messsssyyyyy - make a switch statement later.    
}

function GTPPickingPerformanceQuery {
    param (
        [string]$targetDate = (queryDateUpdate),  # Default to today's date
        [GTP_UOM]$uom = 1
    )
    
    Write-Host "The current default selection: $uom, Totes is not available!" -ForegroundColor Green
    $userSelection = userUOMSelection -enumType ([GTP_UOM])
    
    Try {
        if (isValidSysSelectionEnum $userSelection) {
            $uom = $userSelection
        }
    }
    catch {
        Write-Host "Hate these enums..." -ForegroundColor Red
    }

    $Global:currentUOM = [GTP_UOM]$uom # Update global variable for display and KPI calculations. I know this is bad but it is late and I am tired. Will refactor later to remove global variable dependency.

    # Similar to decant performance, we will need to adjust the query based on UOM selection.
    # Do we need to worry about batch picking?!? Do we need to seperate CMC to many questions..
    [string]$GTP_QUERY_EACHES
    [string]$GTP_QUERY_TOTES
    [string]$GTP_QUERY_CARTONS

    $GTP_QUERY_EACHES = @" 
    SELECT 
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        SUM(qty) AS [Eaches]
    FROM 
        mi_pick
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND oel_class = 'OEL_PICK_PICKED'
        AND change_uid IS NOT NULL
        AND qty IS NOT NULL
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    # STOCK TOTES IS GONNA BE A ROUGHHHH ONE TO CALCULATE.
    # Reason being is one stock tote can go to three different cartons in GTP station
    # We do 1 to 3 picking set up... rip. Maybe remove entire thing as options...
    $GTP_QUERY_TOTES = @"
    NOT IN USE!!! Enjoy the error if you manage to call this...
"@

    $GTP_QUERY_CARTONS = @"
    SELECT 
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(stock_tm_id) AS [Eaches]
    FROM 
        mi_pick
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND oel_class = 'OEL_PICK_PICKED'
        AND change_uid IS NOT NULL
        AND qty IS NOT NULL
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    switch ([int]$uom) {
        1 { return $GTP_QUERY_EACHES }
        2 { return $GTP_QUERY_TOTES }
        3 { return $GTP_QUERY_CARTONS }
        default { Write-Host "Invalid UOM selection. No idea how you could achieve this Defaulting to $uom." -ForegroundColor Yellow; return $GTP_QUERY_EACHES }
    }
}

function TestQuery {
    return @"
    SELECT TOP 10 *
    FROM x_du
"@
}

function ConsumableUsageQuery {
    #ToDo -> Look through the TM. 
    [string]$ConsumableQueryString
    
    $ConsumableQueryString = @"
    SELECT 
    CAST(state_change_time AS DATE) AS [Date],
    planned_tm_sub_type_id AS [CartonType],
    COUNT(DISTINCT CASE 
        WHEN planned_tm_sub_type_id = 'CMC CARTON' THEN du_id 
        ELSE tm_id 
        END) AS [CartonCount]
    FROM x_du
    WHERE du_state IN ('LOADED','BUFFERED','PACKED','FINISHED')
        AND order_packing_type != 'REPACK'
        AND state_change_time >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
        AND state_change_time <= GETDATE()
    GROUP BY CAST(state_change_time AS DATE), planned_tm_sub_type_id
    ORDER BY CAST(state_change_time AS DATE), planned_tm_sub_type_id
"@

return $ConsumableQueryString
}

function EachesPerDayQuery {
    
    param(
        [string]$targetDate = (queryDateUpdate),
        [string]$userFilter = $true #Will make a switch later.
    )

    [string]$PriorityTime = "priority_time"
    [String]$DespatchTime = "required_despatch_time"
    
    switch ($userFilter) {
        $true { $filter = $DespatchTime; break }
        Default { $filter = $PriorityTime}
    }
    

    return @"
    SELECT 
        pick_state AS [State],
        DATEPART(DAY,$filter) AS [DateDay],
        SUM(each_qty) AS [Eaches]
    FROM x_pick
    GROUP BY pick_state, DATEPART(DAY,$filter);
"@

}

function GTPUtilisationQuery {
    return @"
    SELECT
    event_time,
    tm_type,
    tm_id,
        'GTP' + RIGHT('0' + SUBSTRING(
            location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        ), 2) AS GTP_code
    FROM [prodmis].[dbo].[mi_tm]
    WHERE (event_time > DATEADD(DAY, -3, GETDATE()))
    AND change_field = 'location'
    AND (location LIKE 'GTP-PICK-%:01' OR location LIKE 'GTP-PUT-%:01')
    ORDER BY event_time ASC;
"@
}

function BrandDistributionQuery {
    return @"
    SELECT 
        x_sku.brand_id,
        SUBSTRING(x_stock.tm_location, 3, 2) AS AisleID,
        x_stock.*, GETDATE() AS query_time 
    FROM x_stock
    JOIN x_sku ON x_stock.sku_id = x_sku.sku_id
        WHERE x_stock.tm_id LIKE '8%'
    AND x_stock.qty > 0
    AND x_stock.tm_location LIKE 'MS%';
"@
}



<#

 QUERIES SECTION END
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#>
<#

 PROCESSING SCRIPTS SECTION START

#>
function UserSelectDate{
    
    # Prompt for date selection
    $dateInput = Read-Host "Enter date (YYYY-MM-DD) or press Enter for today"
    
    if ([string]::IsNullOrWhiteSpace($dateInput)) {
        $targetDate = Get-Date -Format "yyyy-MM-dd"
    } else {
        try {
            $targetDate = [DateTime]::Parse($dateInput).ToString("yyyy-MM-dd")
        }
        catch {
            Write-Host "Invalid date format. Using today's date." -ForegroundColor Yellow
            $targetDate = Get-Date -Format "yyyy-MM-dd"
        }
    }

    return $targetDate
}

function userUOMSelection {
    param (
        [type]$enumType
    )

    # Get all enum names and values dynamically from whatever enum is passed in
    $enumValues = [Enum]::GetValues($enumType)

    Write-Host "`nSelect a Unit of Measure:"
    foreach ($val in $enumValues) {
        Write-Host "  $([int]$val) - $val"
    }
    Write-Host ""

    $userInput = Read-Host "Enter selection (or press Enter to use default)"

    # Empty input — signal caller to use its default
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        return -1
    }

    # Validate input against the enum
    $parsed = 0
    if ([int]::TryParse($userInput, [ref]$parsed)) {
        if ($enumValues -contains [Enum]::ToObject($enumType, $parsed)) {
            return $parsed
        }
    }

    Write-Host "Invalid selection. Using default." -ForegroundColor Yellow
    return -1
}

function queryDateUpdate{
    $targetDate = UserSelectDate
    return $targetDate
}

function TableDisplay {
    
    param(
        [Parameter(Mandatory=$true)]
        [array]$Data,
        
        [Parameter(Mandatory=$true)]
        [string]$RowProperty,      # e.g., "User"
        
        [Parameter(Mandatory=$true)]
        [string]$ColumnProperty,   # e.g., "Hour"
        
        [Parameter(Mandatory=$true)]
        [string]$ValueProperty,    # e.g., "Eaches"
        [hashtable]$ColorThresholds = @{},  # Optional color coding

        
        
        [Type]$ColumnEnumOverride = $null,  # NEW: Enum type to use for columns instead

        [switch]$ShowEnumValues, # NEW: Whether to show enum names or values in column headers
        
        [int]$MinRowPadding = 15,
        [int]$MinColPadding = 6,
        [int]$ExtraPadding = 2  # Extra space between columns for breathing room
    )

    # Determine columns - either from enum or from data
    if ($null -ne $ColumnEnumOverride) {
        # Use enum values as columns
        $columns = [Enum]::GetValues($ColumnEnumOverride) | Sort-Object
    } else {
        # Extract unique columns from data (original behavior)
        $columns = $Data | Select-Object -ExpandProperty $ColumnProperty -Unique | Sort-Object
    }

    # Extract unique rows from data
    $rows = $Data | Select-Object -ExpandProperty $RowProperty -Unique | Sort-Object

    # Calculate optimal row width (based on longest row label)
    $maxRowLength = ($rows | ForEach-Object { $_.ToString().Length } | Measure-Object -Maximum).Maximum
    $maxRowLength = [Math]::Max($maxRowLength, $RowProperty.Length)
    $RowPadding = [Math]::Max($MinRowPadding, $maxRowLength + $ExtraPadding)

    # Calculate optimal column widths (each column can have different width)
    $columnWidths = @{}
    
    #Cluster Truck!
    foreach ($col in $columns) {
        $colDisplay = if ($ColumnProperty -eq "Hour") 
        { "{0:D2}h" -f [int]$col 
        } elseif ($null -ne $ColumnEnumOverride) {
            # if we're using an enum override. 
                if ($ShowEnumValues) { [int]$col } else {$col.ToString()}
            } else { 
                $col 
            }

        $maxColWidth = $colDisplay.ToString().Length
        
        # Check all values in this column 
        $columnValues = $Data | Where-Object { $_.$ColumnProperty -eq $col } | 
                        Select-Object -ExpandProperty $ValueProperty
        $maxValueWidth = ($columnValues | ForEach-Object { 
            if ($null -ne $_ -and $_ -ne 0) { $_.ToString().Length } else { 1 } 
        } | Measure-Object -Maximum).Maximum
        
        if ($null -ne $maxValueWidth) {
            $maxColWidth = [Math]::Max($maxColWidth, $maxValueWidth)
        }
        
        $columnWidths[$col] = [Math]::Max($MinColPadding, $maxColWidth + $ExtraPadding)
    }

    # Calculate total width for the table
    $totalColWidth = ($columnWidths.Values | Measure-Object -Sum).Sum
    $totalWidth = $RowPadding + $totalColWidth + 10

    # Header row
    Write-Host ($RowProperty.PadRight($RowPadding)) -NoNewline -ForegroundColor Cyan
    foreach ($col in $columns) {
        $colDisplay = if ($ColumnProperty -eq "Hour") { 
            "{0:D2}h" -f [int]$col 
    } elseif ($null -ne $ColumnEnumOverride) {
            if ($ShowEnumValues) { [int]$col} else { $col.ToString() }
        } else { 
            $col 
        }

        $width = $columnWidths[$col]
        Write-Host ($colDisplay.ToString().PadLeft($width)) -NoNewline -ForegroundColor Cyan
    }
    Write-Host ("  Total".PadLeft(8)) -ForegroundColor Cyan
    Write-Host ("-" * $totalWidth) -ForegroundColor Gray

    # Data rows
    foreach ($row in $rows) {
        $rowTotal = 0
        Write-Host ($row.ToString().PadRight($RowPadding)) -NoNewline

        foreach ($col in $columns) {
            $width = $columnWidths[$col]
            
            # Get the value for this row/column intersection
            $value = ($Data | Where-Object { 
                $_.$RowProperty -eq $row -and $_.$ColumnProperty -eq $col 
            }).$ValueProperty

            if ($null -eq $value -or $value -eq 0) {
                Write-Host (" " * ($width - 1) + "-") -NoNewline -ForegroundColor DarkGray
            } else {
                $rowTotal += $value
                
                # Apply color coding if thresholds provided
                $color = Get-CellColor -Value $value -Thresholds $ColorThresholds
                Write-Host ("{0,$width}" -f $value) -NoNewline -ForegroundColor $color
            }
        }
        Write-Host ("{0,8}" -f $rowTotal) -ForegroundColor Cyan
    }

    # Column totals row
    Write-Host ("-" * $totalWidth) -ForegroundColor Gray
    Write-Host (($ColumnProperty.ToUpper() + " TOTAL").PadRight($RowPadding)) -NoNewline -ForegroundColor Cyan
    
    $grandTotal = 0
    foreach ($col in $columns) {
        $width = $columnWidths[$col]
        $colTotal = ($Data | Where-Object { $_.$ColumnProperty -eq $col } | 
                     Measure-Object -Property $ValueProperty -Sum).Sum
        if ($colTotal -gt 0) {
            Write-Host ("{0,$width}" -f $colTotal) -NoNewline -ForegroundColor Cyan
            $grandTotal += $colTotal
        } else {
            Write-Host (" " * ($width - 1) + "-") -NoNewline -ForegroundColor DarkGray
        }
    }
    Write-Host ("{0,8}" -f $grandTotal) -ForegroundColor Green

}



function Get-CellColor {
    param(
        [int]$Value,
        [hashtable]$Thresholds
    )

    if ($Thresholds.Count -eq 0) {
        return "White"
    }

    if ($Value -ge $Thresholds.High) {
        return "Green"
    } elseif ($Value -ge $Thresholds.Medium) {
        return "Yellow"
    } else {
        return "Red"
    }
}


function FillPercentage {
    param (
        [string]$query = (FillPercentageQuery)
    )

    Clear-Host
    Write-Host "Raw Fill Data Query" -ForegroundColor Green

    try {
        # Execute the query using SQLdirector function
        $data = SQLdirector -query $query 
        $data | Format-Table -AutoSize
        Pause
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Pause
    }
    finally {
        Write-Host "Finished executing Fill Percentage." -ForegroundColor Yellow
    }
}

function DecantPerformance {
    param (
        [string]$query = (DecantPerformanceQuery),
        [int]$refreshInterval = 7000
    )

    #Unsure if want another function to obtain DecantKPISettings but all welllll.
    $Decant_high = 100
    $Decant_medium = 50    
    $continueRunning = $true
    #Refresh does not work!

    


    while ($continueRunning) {
       
        Clear-Host
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host "                  DECANT PERFORMANCE - HOURLY BREAKDOWN" -ForegroundColor Cyan
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "UOM DISPLAYED BELOW IS: $($Global:currentUOM)" -ForegroundColor Red

        try {
             # Execute the query
            $data = SQLdirector -query $query
            if ($data.Rows.Count -eq 0) {
                Write-Host "No decant data found for $targetDate" -ForegroundColor Yellow
                Pause
                return
            }

            # Convert to PowerShell objects for easier manipulation
            $results = @()
            
            foreach ($row in $data) {
                $results += [PSCustomObject]@{
                    User   = $row.User
                    Hour   = $row.Hour
                    Eaches = $row.Eaches
                }
            
            }
            
            TableDisplay -Data $results `
                    -RowProperty "User" `
                    -ColumnProperty "Hour" `
                    -ValueProperty "Eaches" `
                    -ColumnEnumOverride ([HOURS_DAY]) `
                    -ColorThresholds @{High = $Decant_high; Medium = $Decant_medium}
                    


            # Last updated timestamp
            Write-Host ""
            Write-Host "Last Updated: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Green
            Write-Host "Press Enter to exit." -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error retrieving decant performance data: $_" -ForegroundColor Red
            Pause
        }

        # Refresh countdown with exit options
        Write-Host ""
        $secondsRemaining = $refreshInterval
        $keyPressed = $false

        while ($secondsRemaining -gt 0 -and -not $keyPressed) {
            
            if ($host.UI.RawUI.KeyAvailable) {
                $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                
                # Check for Enter (13), Esc (27), or Q (81)
                if ($key.VirtualKeyCode -eq 13 -or $key.VirtualKeyCode -eq 27 -or $key.VirtualKeyCode -eq 81) {
                    $keyPressed = $true
                    $continueRunning = $false
                    Clear-Host
                    Write-Host "Returning to menu..." -ForegroundColor Yellow
                    break
                }
            }

            $minutes = [math]::Floor($secondsRemaining / 60)
            $seconds = $secondsRemaining % 60

            Write-Host ("`rNext refresh in: {0:D2}:{1:D2}  (Press Enter/Esc/Q to exit)" -f $minutes, $seconds) -NoNewline -ForegroundColor Cyan
            
            Start-Sleep -Seconds 1
            $secondsRemaining--
        }

        if (-not $keyPressed) {
            Write-Host "`n`rRefreshing data..." -ForegroundColor Yellow
        }
    }
}

function GTPPickingPerformance {
    
    param (
        [string]$query = (GTPPickingPerformanceQuery),
        [int]$refreshInterval = 7000
    )
    
    # Huhhhhhhhhhhhhhh
    $GTPPicking_high = 200; 
    $GTPPicking_medium = 100;

    #cannotttt wait to get rid of this global variable dependency. I should not have done this...
    $continueRunning = $true

    while ($continueRunning) {
       
        Clear-Host
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host "                  GTP PICKING PERFORMANCE - HOURLY BREAKDOWN" -ForegroundColor Cyan
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "UOM DISPLAYED BELOW IS: $($Global:currentUOM)" -ForegroundColor Red
        #Need to update the function later...

        try {
             # Execute the query
            $data = SQLdirector -query $query
            if ($data.Rows.Count -eq 0) {
                Write-Host "No GTP PICKING data found for $targetDate" -ForegroundColor Yellow
                Pause
                return
            }

            # Convert to PowerShell objects for easier manipulation
            $results = @()
            
            foreach ($row in $data) {
                $results += [PSCustomObject]@{
                    User   = $row.User
                    Hour   = $row.Hour
                    Eaches = $row.Eaches
                }
            
            }
            
            TableDisplay -Data $results `
                    -RowProperty "User" `
                    -ColumnProperty "Hour" `
                    -ValueProperty "Eaches" `
                    -ColumnEnumOverride ([HOURS_DAY]) `
                    -ColorThresholds @{High = $GTPPicking_high; Medium = $GTPPicking_medium}
                    

            # Last updated timestamp
            Write-Host ""
            Write-Host "Query Ran at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Green
            Write-Host "Press Enter to exit." -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error retrieving GTP Picking performance data: $_" -ForegroundColor Red
            Pause
        }

        # Refresh countdown with exit options
        Write-Host ""
        $secondsRemaining = $refreshInterval
        $keyPressed = $false

        while ($secondsRemaining -gt 0 -and -not $keyPressed) {
            
            if ($host.UI.RawUI.KeyAvailable) {
                $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                
                # Check for Enter (13), Esc (27), or Q (81)
                if ($key.VirtualKeyCode -eq 13 -or $key.VirtualKeyCode -eq 27 -or $key.VirtualKeyCode -eq 81) {
                    $keyPressed = $true
                    $continueRunning = $false
                    Clear-Host
                    Write-Host "Returning to menu..." -ForegroundColor Yellow
                    break
                }
            }

            $minutes = [math]::Floor($secondsRemaining / 60)
            $seconds = $secondsRemaining % 60

            Write-Host ("`rNext refresh in: {0:D2}:{1:D2}  (Press Enter/Esc/Q to exit)" -f $minutes, $seconds) -NoNewline -ForegroundColor Cyan
            
            Start-Sleep -Seconds 1
            $secondsRemaining--
        }

        if (-not $keyPressed) {
            Write-Host "`n`rRefreshing data..." -ForegroundColor Yellow
        }
    }
    


}

function ConsumableUsage {
    param (
        [string]$query = (ConsumableUsageQuery)
    )

    Clear-Host
    Write-Host "Consumable Usage Query" -ForegroundColor Green
    Write-Host "Showing count of cartons used in the last 7 days by carton type." -ForegroundColor Green
    $fromDate = (Get-Date).AddDays(-7)
    $toDate = Get-Date
    Write-Host "Date Range: $($fromDate.ToShortDateString()) - $($toDate.ToShortDateString())" -ForegroundColor Green
    Write-Host "WARNING: THIS QUERY IS NOT SEARCHING MI_DU THEREFORE MIGHT BE MISSING ALOT" -ForegroundColor Red
    try {
        # Execute the query using SQLdirector function
        $data = SQLdirector -query $query 

        $results = @()
                    
        foreach ($row in $data) {
            $results += [PSCustomObject]@{
                Date   = $row.Date
                CartonType   = $row.CartonType
                Eaches = $row.CartonCount
            }

        }

        TableDisplay -Data $results `
                -RowProperty "Date" `
                -ColumnProperty "CartonType" `
                -ValueProperty "Eaches" 
        
        Write-Host ""
        Write-Host "WARNING: THIS QUERY IS NOT SEARCHING MI_DU THEREFORE MIGHT BE MISSING ALOT" -ForegroundColor Red
        Pause
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Pause
    }
    

}

function EachesPerDay {
    param (
        [string]$query = (EachesPerDayQuery)
    )

    Clear-Host
    Write-Host "Eaches Per Day Query" -ForegroundColor Green

    try {
        # Execute the query using SQLdirector function
        $data = SQLdirector -query $query 
        
        # Convert to PowerShell objects for easier manipulation
        $results = @()
                    
        foreach ($row in $data) {
            $results += [PSCustomObject]@{
                DateDay   = $row.DateDay
                State   = $row.State
                Eaches = $row.Eaches
            }

        }

        TableDisplay -Data $results `
                -RowProperty "DateDay" `
                -ColumnProperty "State" `
                -ValueProperty "Eaches" `
                -ColumnEnumOverride ([PICK_STATE_DASHBOARD])
                 
                
        Pause
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Pause
    }
    

}
<#

 PROCESSING SCRIPTS SECTION END

#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#>
<#

 SQL PROCESSING SECTION START

 THIS SHOULD REQUIRE MINIM CHANGES TO THE SQL CONNECTION CODE IN THE FUTURE IF WE DECIDE TO SWITCH TO SQL AUTH OR CHANGE SERVERS, ETC.

#>

# SQLdirector function - simplified for Windows Authentication
function SQLdirector {
    param (
        [string]$query
    )

    try {
        $data = RunSqlQuery -SQLServer $global:SQLServer `
                            -SQLDatabase $global:SQLDatabase `
                            -Query $query
        return $data
    }
    catch {
        Write-Host "Database query error: $_" -ForegroundColor Red
        throw
    }
}

# RunSqlQuery function - Windows Authentication with modern encryption
function RunSqlQuery {
    param (
        [string]$SQLServer,
        [string]$SQLDatabase,
        [string]$Query
    )

    # Modern connection string with Windows Auth and encryption
    $connectionString = "Server=$SQLServer;Database=$SQLDatabase;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30;"
    
    $connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $command    = $connection.CreateCommand()
    $command.CommandText = $Query
    $adapter    = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $table      = New-Object System.Data.DataTable

    try {
        $connection.Open()
        [void]$adapter.Fill($table)
        return $table
    }
    catch {
        Write-Host "Database connection error: $_" -ForegroundColor Red
        throw
    }
    finally {
        if ($connection.State -eq 'Open') {
            $connection.Close()
        }
        $connection.Dispose()
    }
}
<#

 SQL PROCESSING SECTION END

#>
function InventoryQueryMenu {
    Write-Host "Inventory Query Menu not yet implemented" -ForegroundColor Yellow
    Pause
}

function Method2 {
    Write-Host "Method2 not yet implemented" -ForegroundColor Yellow
    Pause
}

function Troubleshoot {
    Write-Host "Troubleshoot not yet implemented" -ForegroundColor Yellow
    Pause
}

function Extras {
    Write-Host "Extras not yet implemented" -ForegroundColor Yellow
    Pause
}

function HelpMenu {
    Clear-Host
    Write-Host "Help Menu" -ForegroundColor Cyan
    Write-Host "==========" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This script provides various database query options."
    Write-Host "Using Windows Authentication with your current credentials."
    Write-Host "Reach out to Nathaniel Ritchie if you ever need help!!!"
    Write-Host ""
    Pause
}

function runProgram {
    # Checking the module installing
     moduleCheck
 
     $size.Width = 150
     $size.Height = 50
     $Host.UI.RawUI.WindowSize = $size
 
     # Launch main menu
     MainMenu
 }
 


runProgram