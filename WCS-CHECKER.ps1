Clear-Host
$Host.UI.RawUI.WindowTitle = "Nathaniel Ritchie RandomScript"

# Force TLS 1.2 or higher (TLS 1.0 and 1.1 are deprecated)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13


# Seans test mode - nr
$nrtestmode = $true

# Global ENUM do not touch please - all SQL is based upon this for selection of UOM.

enum UOM {
    Eaches = 1
    Totes = 2
    Cartons = 3
}

enum DECANT_UOM{
    Eaches = 1
    Totes = 2
    Cartons = 3
}


enum PICK_STATE{
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
        Write-Host "             [2] Inventory Query Menu                     - AUKC01"
        Write-Host "             [3] KPIsMenu                                 - AUKC01"
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
            "2" { InventoryQueryMenu }
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
        Write-Host "             [2]  Picking Performance"
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
            "2"  { Write-Host "Picking Performance not yet implemented" -ForegroundColor Yellow; Pause }
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

#>
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
        [string]$targetDate = "2026-02-12", #(Get-Date -Format "yyyy-MM-dd"),  # Default to today's date
        

        # Will change to param for user input
        [DECANT_UOM]$uom = 3
        #1 is EACH
        #2 is TOTE
        #3 is CARTONS  # Assuming 1 represents Eaches, can be adjusted based on actual UOM values in the database
    )


    # messsssyyyyy - make a switch statement later.    
    if ($uom -eq 1) {
        return @"
        SELECT 
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        SUM(quantity) AS [Eaches]
    FROM mi_decant
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
        AND change_uid IS NOT NULL
        AND quantity IS NOT NULL
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@
    } else {
        
        # We want to count number of TOTES. Therefore we change the query.
        if ($uom -eq 2){
            return @"
            SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(tote_id) AS [Eaches]
        FROM mi_decant
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
            AND change_uid IS NOT NULL
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

        } else {
            if ($uom -eq 3){
                return @"
                SELECT 
                change_uid AS [User],
                DATEPART(HOUR, event_time) AS [Hour],
                COUNT(DISTINCT(case_id)) AS [Eaches]
            FROM mi_decant
            WHERE CAST(event_time AS DATE) = '$targetDate'
                AND oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
                AND change_uid IS NOT NULL
                AND quantity IS NOT NULL
            GROUP BY change_uid, DATEPART(HOUR, event_time)
            ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

            }
        }

    }
    

}


function EachesPerDayQuery {
    
    param(
        [string]$targetDate = "2026-02-12", #(Get-Date -Format "yyyy-MM-dd")  # Default to today's date
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


function TableDisplay{
    
        # Just pulling in all the information.
        # Data is the raw data. This is not to be manipulated here - only for display purposes.
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
            
            [int]$MinRowPadding = 15,
            [int]$MinColPadding = 6,
            [int]$ExtraPadding = 2  # Extra space between columns for breathing room
        )
    
        # Extract unique rows and columns from data
        $rows = $Data | Select-Object -ExpandProperty $RowProperty -Unique | Sort-Object
        $columns = $Data | Select-Object -ExpandProperty $ColumnProperty -Unique | Sort-Object
    
        # Calculate optimal row width (based on longest row label)
        $maxRowLength = ($rows | ForEach-Object { $_.ToString().Length } | Measure-Object -Maximum).Maximum
        $maxRowLength = [Math]::Max($maxRowLength, $RowProperty.Length)
        $RowPadding = [Math]::Max($MinRowPadding, $maxRowLength + $ExtraPadding)
    
        # Calculate optimal column widths (each column can have different width)
        $columnWidths = @{}
        foreach ($col in $columns) {
            $colDisplay = if ($ColumnProperty -eq "Hour") { "{0:D2}h" -f $col } else { $col }
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
            $colDisplay = if ($ColumnProperty -eq "Hour") { "{0:D2}h" -f $col } else { $col }
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
        [int]$refreshInterval = 60
    )

    #Unsure if want another function to obtain DecantKPISettings but all welllll.
    $Decant_high = 100
    $Decant_medium = 50    
    $continueRunning = $true


    
       
        Clear-Host
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host "                  DECANT PERFORMANCE - HOURLY BREAKDOWN" -ForegroundColor Cyan
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host ""

        
        #$targetDate = Get-Date -Format "yyyy-MM-dd"
        $targetDate = "Linklost..." #placeholder for now - need to edit the query string to pass this in properly.
        Write-Host "Querying decant data for: $targetDate" -ForegroundColor Green
        Write-Host ""

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
                    -ColorThresholds @{High = $Decant_high; Medium = $Decant_medium}


            Write-Host ""
            Write-Host "==================================================================" -ForegroundColor Cyan
            Write-Host "Color Legend: " -NoNewline
            Write-Host "Green >= ?SGJN eaches/hr  " -NoNewline -ForegroundColor Green
            Write-Host "Yellow >= DSLFKJS eaches/hr  " -NoNewline -ForegroundColor Yellow
            Write-Host "White < SDL eaches/hr" -ForegroundColor White
            Write-Host "==================================================================" -ForegroundColor Cyan
            Write-Host ""


            # Last updated timestamp
            Write-Host ""
            Write-Host "Last Updated: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Green
            
        }
        catch {
            Write-Host "Error retrieving decant performance data: $_" -ForegroundColor Red
            Pause
        }

        <# Refresh fix but not working so fix needs a fix.
        Write-Host ""
        $secondsRemaining = $refreshInterval
        $keyPressed = $false

        while ($secondsRemaining -gt 0 -and -not $keyPressed -or $nrtestmode) {
            
            if ($host.UI.RawUI.KeyAvailable) {
                $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                
                if ($key.VirtualKeyCode -eq 13 -or $key.VirtualKeyCode -eq 27 -or $key.VirtualKeyCode -eq 81) {
                    $keyPressed = $true
                    $continueRunning = $false
                    Clear-Host
                    Write-Host "Returning to menu..." -ForegroundColor Yellow
                }
            }

            $minutes = [math]::Floor($secondsRemaining / 60)
            $seconds = $secondsRemaining % 60

            [int]$intMinutes = $minutes
            [int]$intSeconds = $seconds

            Write-Host ("`rNext refresh in: {0:D2}:{1:D2}  (Press Enter/Esc/Q to exit)" -f $intMinutes, $intSeconds) -NoNewline -ForegroundColor Cyan
            
            Start-Sleep -Seconds 1
            $secondsRemaining--

            if (-not $keyPressed){
             Write-Host "`n`rRefreshing data..." -ForegroundColor Yellow
            
            } 
        #>
        
    
    
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
                -ValueProperty "Eaches"
                #-ColorThresholds @{High = $Decant_high; Medium = $Decant_medium}

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
    Write-Host ""
    Pause
}


# Starting the Scripts. 
MainMenu