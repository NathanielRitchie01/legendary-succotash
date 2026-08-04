# Script requires minimum PowerShell 5.1 wonder if I can force an update and if that fails then just exit with a message to user.

<#
.SYNOPSIS
    WCS Checker - A PowerShell script to query and display some metrics from our WCS system.
.DESCRIPTION
    A PowerShell-based warehouse control system dashboard that queries SQL Server.
.NOTES
    Author  : Nathaniel Ritchie
    Site    : AUKC01
    Requires: SqlServer module (auto-installed if missing)
    Auth    : Windows Integrated Authentication.
.EXAMPLE
    .\WCS-Checker.ps1
    Launches the WCS Checker dashboard with interactive menus for various performance metrics.
.NOTES
    A WinForms GUI front-end also exists: WCS-Checker-GUI.ps1 (same folder). It dot-sources
    this file to reuse its functions - see the Start-WCSCheckerConsole guard near the bottom.
#>


#region declareEnums

enum UOM {
    EACHES = 1
    TOTES = 2
    CARTONS = 3
}

enum DECANT_UOM{
    EACHES = 1
    TOTES = 2
    CARTONS = 3
}

enum GTP_UOM {
    EACHES = 1
    #Totes not supported. - 1-to-3 picking makes tote counting unreliable
    CARTONS = 3
}

enum REPLEN_UOM {
    #Eaches not supporty - Linking between case - tm is not worth it. EACHES = 1
    #Totes not supported. - Totes are not in automation for picking.
    CARTONS = 3
    PALLET = 4
}

enum PUTAWAY_UOM {
    EACHES = 1
    #Totes not supported. Totes is automation? Should not be sending geting delivery from cleint for totes I hope..
    CARTONS = 3
    PALLET = 4
}

enum RECEIVING_UOM {
    EACHES = 1
    #Totes not supported. Totes is automation? Should not be sending geting delivery from cleint for totes I hope..
    CARTONS = 3
    PALLET = 4
}

enum INBOUND_TRAMMING_UOM {
    # Only pallets thanks.
    PALLET = 4
}

enum MANUALPICKING_UOM {

    CARTONS = 3
    PALLET = 4
}


enum GTPUTILISATION_UOM {
    CARTON_CMC = 1
    TOTES   = 2
    CARTON_CMC_TOTES_COMBINED = 3
}

enum RETURNS_UOM {
    # Numbers here don't follow the EACHES/TOTES/CARTONS/PALLET convention above -
    # returns doesn't have a physical UOM, these are just different measures of the same activity.
    TOTAL_QTY    = 1
    SALEABLE     = 2
    NON_SALEABLE = 3
    DELIVERIES   = 4
}

enum VAS_UOM {
    EACHES  = 1
    #Totes not applicable.
    CARTONS = 3
}

enum QI_UOM {
    SKUS    = 1
    #Matches convention loosely - tm_id here is effectively a pallet at the QI PALLET station.
    PALLETS = 4
}

enum HOURS_DAY {
    Hour00 = 0;  Hour01 = 1;  Hour02 = 2;  Hour03 = 3
    Hour04 = 4;  Hour05 = 5;  Hour06 = 6;  Hour07 = 7
    Hour08 = 8;  Hour09 = 9;  Hour10 = 10; Hour11 = 11
    Hour12 = 12; Hour13 = 13; Hour14 = 14; Hour15 = 15
    Hour16 = 16; Hour17 = 17; Hour18 = 18; Hour19 = 19
    Hour20 = 20; Hour21 = 21; Hour22 = 22; Hour23 = 23
    
}

enum PICK_STATE{
    CREATING; EXPECTED; PENDING; WAIT_ALLOCATION; UNSATISFIABLE
    ALLOCATED; UNPICKABLE; WAIT_STOCK; RESERVED; STARTED
    PICKED; COLLATING; COLLATED; PACKABLE; PACKING; PACKED
    BUFFERED; UNREACHABLE; MARSHALLING; MARSHALLED
    LOADING; LOADED; DESPATCHED; FINISHED; ABANDONED; CANCELLED
}

enum PICK_STATE_DASHBOARD {
    <#Subset of PICK_STATE but only ones we care for on dashboard#>
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
#endregion








#region configuration
<#
.SYNOPSIS
    Returns a hashtable of connection and display settings.
.DESCRIPTION
    Central place for all configurable values. Edit this ONE function
    when the server, database, or KPI thresholds change.
.OUTPUTS
    [hashtable] with keys: SQLServer, SQLDatabase,
    DecantThresholds, GTPThresholds
.EXAMPLE
    $cfg = Get-DashboardConfig
    $cfg.SQLServer          # → "SQLDBAUP010"
    $cfg.DecantThresholds   # → @{ High = 100; Medium = 50 }
.NOTES
    This is AI btw - no idea if this is how I want to structure config but it is better than scattering variables throughout code. Can also add other settings like display thresholds in here later.
#>
function Get-DashboardConfig {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()



    $config = @{
        SQLServer             = "SQLDBAUP010"
        SQLDatabase           = "prodmis"

        # Ideally these thresholds would be based on historical performance data and aligned with business goals. For now, just placeholders to demonstrate colour coding.
        # Allso consider making these configurable at runtime in future updates.
        # And also other function for overarching rather than just this? Who Knows...
        DecantThresholds      = @{ High = 100; Medium = 50 }
        GTPThresholds         = @{ High = 200; Medium = 100 }
        ReplenishmentThresholds = @{ High = 150; Medium = 75 }
        GTPUTILISATIONThresholds  = @{ High = 100; Medium = 75 }
        PutawayThresholds     = @{ High = 100; Medium = 90 }
        ReceivingThresholds = @{ High = 100; Medium = 90 }
        InboundTrammingThresholds = @{ High = 50; Medium = 20 }

        # New KPIs added 2026-07 - same "placeholder, not tuned yet" disclaimer as above.
        ReturnsThresholds      = @{ High = 30; Medium = 10 }
        CubiThresholds         = @{ High = 50; Medium = 20 }
        VASThresholds          = @{ High = 50; Medium = 20 }
        QIThresholds           = @{ High = 50; Medium = 20 }
        ParcelRepackThresholds = @{ High = 50; Medium = 20 }
        DespatchThresholds     = @{ High = 100; Medium = 50 }   
        DepartmentLabourThresholds = @{ Green = 40; Yellow = 20 }
        ManualPickingThresholds = @{ Green = 40; Yellow = 20 }
    }   

    

    return $config
}
#endregion








#region databaseConnection


<#
.SYNOPSIS
    Ensures the SqlServer module is installed.
.DESCRIPTION
    Need to check if user has the SqlServer module installed. If not, it attempts to install it.
.EXAMPLE
    Assert-SqlModule
#>
function Assert-SqlModule {
    [CmdletBinding()]
    param()
    

    if ($Script:TestMode) {
        Write-Host "TEST MODE: SqlServer module check skipped." -ForegroundColor Magenta
        return
    }

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
    }
    else {
        Write-Host "SqlServer module is already installed." -ForegroundColor Green
    }
}


<#
.SYNOPSIS
    Ensures the Graphical module is installed.
.DESCRIPTION
    Need to check if user has the Graphical module installed. If not, it attempts to install it.
.EXAMPLE
    Assert-GraphicalModule
.NOTES
    This is NOT being used right now. In future updates - need to review.
#>
function Assert-GraphicalModule {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Graphical)) {
        try {
            Write-Host "Graphical module not found. Installing..." -ForegroundColor Yellow
            Install-Module -Name Graphical -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
            Write-Host "Graphical module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install Graphical module. Please install it manually." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Graphical module is already installed." -ForegroundColor Green
    }
}


<#
.SYNOPSIS
    Executes a SQL query against the configured database.
.DESCRIPTION
    Opens a SqlClient connection using users Windows Integrated Authentication,
    executes the given query, and returns the result as a DataTable.
.PARAMETER Query
    The T-SQL query string to execute.
.PARAMETER Server
    SQL Server instance name. Defaults to the config value.
.PARAMETER Database
    Database name. Defaults to the config value.
.OUTPUTS
    [System.Data.DataTable] — the query result set.
.EXAMPLE
    # Simple query
    $data = Invoke-SqlQueryDirect -Query "SELECT TOP 10 * FROM x_du"

    # With explicit server
    $data = Invoke-SqlQueryDirect -Query $myQuery -Server "MYSERVER" -Database "mydb"
#>
function Invoke-SqlQueryDirect {
    [CmdletBinding()]
    [OutputType([System.Data.DataTable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Query,

        [string]$Server   = (Get-DashboardConfig).SQLServer,
        [string]$Database = (Get-DashboardConfig).SQLDatabase
    )

   


    $connectionString = "Server=$Server;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30;"


    $connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $command    = $connection.CreateCommand()
    $command.CommandText = $Query
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $table   = New-Object System.Data.DataTable

    try {
        $connection.Open()
        
        [void]$adapter.Fill($table)
        
        # NR-NOTE: PowerShell treats DataTable as enumerable (via IListSource) when it flows
        # through a function's output - without the comma, callers can silently get back an
        # array of DataRow instead of the DataTable itself (varies by row count: 0 rows -> $null,
        # 1 row -> a single DataRow, 2+ rows -> DataRow[]). The comma forces exactly one object
        # through so $data = Invoke-SqlQueryDirect ... is always a real, intact DataTable.
        return ,$table
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

#endregion



#region userInputHelpers
<#
.SYNOPSIS
    Prompts the user for a date selection.
.DESCRIPTION
    Shows a Read-Host prompt for a date in YYYY-MM-DD format.
    Returns today's date if the user presses Enter or enters an invalid value.
.OUTPUTS
    [string] — date in "yyyy-MM-dd" format.
.EXAMPLE
    $date = Read-DateSelection
    # User enters "2025-03-15"  →  returns "2025-03-15"
    # User presses Enter        →  returns today e.g. "2025-06-01"
    # User enters "garbage"     →  returns today with a warning
#>
function Read-DateSelection {
    [CmdletBinding()]
    [OutputType([string])]
    param()


    $dateInput = Read-Host "Enter date (YYYY-MM-DD) or press Enter for today"


    if ([string]::IsNullOrWhiteSpace($dateInput)) {
        $today = (Get-Date -Format "yyyy-MM-dd")
        return $today
    }

    try {
        $parsed = [DateTime]::Parse($dateInput).ToString("yyyy-MM-dd")
        return $parsed
    }
    catch {
        Write-Host "Invalid date format. Using today's date." -ForegroundColor Yellow
        return (Get-Date -Format "yyyy-MM-dd")
    }
}


<#
.SYNOPSIS
    Prompts the user to pick a value from any enum type.
.DESCRIPTION
    Dynamically lists all values of the provided enum and lets the user select one.
    Returns -1 if the user presses Enter (signals "use default") or enters an invalid value.
.PARAMETER EnumType
    The [type] of the enum to display — e.g. [DECANT_UOM], [GTP_UOM].
.OUTPUTS
    [int] — the selected enum integer value, or -1 to signal "keep default".
.EXAMPLE
    # If the user types "2" for an enum where 2 = TOTES:
    $sel = Read-EnumSelection -EnumType ([DECANT_UOM])   # → returns 2

    # If the user presses Enter:
    $sel = Read-EnumSelection -EnumType ([DECANT_UOM])   # → returns -1
#>


function Read-EnumSelection {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [type]$EnumType
    )


    $enumValues = [Enum]::GetValues($EnumType)

    Write-Host "`nSelect a Unit of Measure:"
    foreach ($val in $enumValues) {
        Write-Host "  $([int]$val) - $val"
    }
    Write-Host ""

    $userInput = Read-Host "Enter selection (or press Enter to use default)"

    if ([string]::IsNullOrWhiteSpace($userInput)) { 
        return -1 }

    $parsed = 0
    if ([int]::TryParse($userInput, [ref]$parsed)) {
        
        if ($enumValues -contains [Enum]::ToObject($EnumType, $parsed)) {
            return $parsed
        }
    }

    Write-Host "Invalid selection. Using default." -ForegroundColor Yellow
    return -1
}


<#
.SYNOPSIS
    Checks whether the user made an active selection (i.e. did not choose "default").
.PARAMETER UserSelection
    The integer returned from Read-EnumSelection.
.OUTPUTS
    [bool] — $true if the user chose a valid enum value, $false if they chose default (-1).
.EXAMPLE
    Test-ValidEnumSelection -UserSelection 2    # → $true
    Test-ValidEnumSelection -UserSelection -1   # → $false  (user pressed Enter)
#>
function Test-ValidEnumSelection {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [int]$UserSelection
    )

    return $UserSelection -ne -1
}

#endregion








#region displayHelpers
<#
.SYNOPSIS
    Returns a console colour name based on a value and threshold pair.
.PARAMETER Value
    The numeric value to evaluate.
.PARAMETER Thresholds
    A hashtable with keys 'High' and 'Medium'.
    Values >= High → Green, >= Medium → Yellow, else → Red.
    An empty hashtable returns "White" (no colouring).
.OUTPUTS
    [string] — a PowerShell console colour name.
.EXAMPLE
    Get-CellColor -Value 150 -Thresholds @{ High = 100; Medium = 50 }   # → "Green"
    Get-CellColor -Value 75  -Thresholds @{ High = 100; Medium = 50 }   # → "Yellow"
    Get-CellColor -Value 20  -Thresholds @{ High = 100; Medium = 50 }   # → "Red"
    Get-CellColor -Value 20  -Thresholds @{}                            # → "White"
#>
function Get-CellColor {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [int]$Value,
        [hashtable]$Thresholds
    )

    if ($Thresholds.Count -eq 0) { 
        return "White" }

    if ($Value -ge $Thresholds.High)   { 
        return "Green" 
     }
    if ($Value -ge $Thresholds.Medium) { 
        return "Yellow"
     }

    return "Red"
}


#Lmao this comment is huge. Too tired to care but love AI for making commennts.
#I think the comment making and the cmdletbinding it's done has helped me save time
#but you have been warned!
<#
.SYNOPSIS
    Renders a pivot-style console table with optional colour coding.
.DESCRIPTION
    Takes flat row/column/value data and displays it as a padded, coloured
    pivot table in the console, with row totals, column totals, and a grand total.
.PARAMETER Data
    Array of PSCustomObjects containing at least the three properties named
    by RowProperty, ColumnProperty, and ValueProperty.
.PARAMETER RowProperty
    The property name used for row labels (e.g. "User").
.PARAMETER ColumnProperty
    The property name used for column headers (e.g. "Hour").
.PARAMETER ValueProperty
    The property name whose values fill the cells (e.g. "Eaches").
.PARAMETER ColorThresholds
    Optional. Hashtable with 'High' and 'Medium' keys for colour coding cells.
    Values >= High → Green, >= Medium → Yellow, < Medium → Red.
.PARAMETER ColumnEnumOverride
    Optional. An enum [Type] whose values define the column headers instead of
    deriving them from the data.  Useful for showing all 24 hours even when some
    have no data.
.PARAMETER ShowEnumValues
    Switch. When used with ColumnEnumOverride, shows the integer value instead
    of the enum name in column headers.
.PARAMETER MinRowPadding
    Minimum character width for the row label column. Default 15.
.PARAMETER MinColPadding
    Minimum character width for each data column. Default 6.
.PARAMETER ExtraPadding
    Extra spaces added between columns for readability. Default 2.
.EXAMPLE
    # Minimal usage — no colour, columns derived from data:
    Show-PivotTable -Data $results -RowProperty "Date" -ColumnProperty "CartonType" -ValueProperty "CartonCount"

    # With colour thresholds and an enum for columns:
    Show-PivotTable -Data $results `
        -RowProperty "User" -ColumnProperty "Hour" -ValueProperty "Eaches" `
        -ColumnEnumOverride ([HOURS_DAY]) `
        -ColorThresholds @{ High = 100; Medium = 50 }
#>
function Show-PivotTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Data,

        [Parameter(Mandatory)]
        [string]$RowProperty,

        [Parameter(Mandatory)]
        [string]$ColumnProperty,

        [Parameter(Mandatory)]
        [string]$ValueProperty,

        [hashtable]$ColorThresholds = @{},
        [Type]$ColumnEnumOverride   = $null,
        [switch]$ShowEnumValues,
        [int]$MinRowPadding = 15,
        [int]$MinColPadding = 6,
        [int]$ExtraPadding  = 2
    )

    # --- Determine columns ---
    if ($null -ne $ColumnEnumOverride) {
        $columns = [Enum]::GetValues($ColumnEnumOverride) | Sort-Object
    }
    else {
        $columns = $Data | Select-Object -ExpandProperty $ColumnProperty -Unique | Sort-Object
    }

    $rows = $Data | Select-Object -ExpandProperty $RowProperty -Unique | Sort-Object

    # --- Calculate widths ---
    $maxRowLength = ($rows | ForEach-Object { $_.ToString().Length } | Measure-Object -Maximum).Maximum
    $maxRowLength = [Math]::Max($maxRowLength, $RowProperty.Length)
    $RowPadding   = [Math]::Max($MinRowPadding, $maxRowLength + $ExtraPadding)

    $columnWidths = @{}
    foreach ($col in $columns) {
        $colDisplay = if ($ColumnProperty -eq "Hour") {
            "{0:D2}h" -f [int]$col
        }
        elseif ($null -ne $ColumnEnumOverride) {
            if ($ShowEnumValues) { [int]$col } else { $col.ToString() }
        }
        else { $col }

        $maxColWidth   = $colDisplay.ToString().Length
        $columnValues  = $Data | Where-Object { $_.$ColumnProperty -eq $col } |
                         Select-Object -ExpandProperty $ValueProperty
        $maxValueWidth = ($columnValues | ForEach-Object {
            if ($null -ne $_ -and $_ -ne 0) { $_.ToString().Length } else { 1 }
        } | Measure-Object -Maximum).Maximum

        if ($null -ne $maxValueWidth) {
            $maxColWidth = [Math]::Max($maxColWidth, $maxValueWidth)
        }
        $columnWidths[$col] = [Math]::Max($MinColPadding, $maxColWidth + $ExtraPadding)
    }

    $totalColWidth = ($columnWidths.Values | Measure-Object -Sum).Sum
    $totalWidth    = $RowPadding + $totalColWidth + 10

    # --- Header row ---
    Write-Host ($RowProperty.PadRight($RowPadding)) -NoNewline -ForegroundColor Cyan
    foreach ($col in $columns) {
        $colDisplay = if ($ColumnProperty -eq "Hour") {
            "{0:D2}h" -f [int]$col
        }
        elseif ($null -ne $ColumnEnumOverride) {
            if ($ShowEnumValues) { [int]$col } else { $col.ToString() }
        }
        else { $col }

        $width = $columnWidths[$col]
        Write-Host ($colDisplay.ToString().PadLeft($width)) -NoNewline -ForegroundColor Cyan
    }
    Write-Host ("  Total".PadLeft(8)) -ForegroundColor Cyan
    Write-Host ("-" * $totalWidth) -ForegroundColor Gray

    # --- Data rows ---
    foreach ($row in $rows) {
        $rowTotal = 0
        Write-Host ($row.ToString().PadRight($RowPadding)) -NoNewline

        foreach ($col in $columns) {
            $width = $columnWidths[$col]
            $value = ($Data | Where-Object {
                $_.$RowProperty -eq $row -and $_.$ColumnProperty -eq $col
            }).$ValueProperty

            if ($null -eq $value -or $value -eq 0) {
                Write-Host (" " * ($width - 1) + "-") -NoNewline -ForegroundColor DarkGray
            }
            else {
                $rowTotal += $value
                $color = Get-CellColor -Value $value -Thresholds $ColorThresholds
                Write-Host ("{0,$width}" -f $value) -NoNewline -ForegroundColor $color
            }
        }
        Write-Host ("{0,8}" -f $rowTotal) -ForegroundColor Cyan
    }

    # --- Column totals ---
    Write-Host ("-" * $totalWidth) -ForegroundColor Gray
    Write-Host (($ColumnProperty.ToUpper() + " TOTAL").PadRight($RowPadding)) -NoNewline -ForegroundColor Cyan

    $grandTotal = 0
    foreach ($col in $columns) {
        $width    = $columnWidths[$col]
        $colTotal = ($Data | Where-Object { $_.$ColumnProperty -eq $col } |
                     Measure-Object -Property $ValueProperty -Sum).Sum
        if ($colTotal -gt 0) {
            Write-Host ("{0,$width}" -f $colTotal) -NoNewline -ForegroundColor Cyan
            $grandTotal += $colTotal
        }
        else {
            Write-Host (" " * ($width - 1) + "-") -NoNewline -ForegroundColor DarkGray
        }
    }
    Write-Host ("{0,8}" -f $grandTotal) -ForegroundColor Green
}

<#
.SYNOPSIS
    Waits for the user to choose whether to refresh or return to the menu.
.DESCRIPTION
    Blocking prompt shown after a dashboard renders. Typing R (then Enter) refreshes
    the same screen. Anything else - a blank Enter, Q, or any other input - returns
    to the menu.
.NOTES
    NR-NOTE: this replaces the old auto-refreshing countdown timer, which relied on
    $host.UI.RawUI.KeyAvailable/ReadKey and raw console cursor positioning - both of
    which are inconsistent across hosts (worked in some terminals, silently broken or
    scrolling endlessly in others). Read-Host is a plain blocking line-read that works
    the same way everywhere, at the cost of needing an actual Enter press to register.
    One consequence: a literal Esc keypress can't be distinguished here (Console.ReadLine
    just clears the input line on Esc in most hosts) - pressing Esc then Enter still
    exits, it just takes the extra Enter.
.OUTPUTS
    [bool] — $true = refresh again, $false = user wants to exit.
.EXAMPLE
    if (-not (Wait-UserAction)) { return }
#>
function Wait-UserAction {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $response = Read-Host "Press Enter to return to the menu, or type R and press Enter to refresh"

    if ($response -match '^[Rr]$') {
        Write-Host "Refreshing data..." -ForegroundColor Yellow
        return $true
    }

    Write-Host "Returning to menu..." -ForegroundColor Yellow
    return $false
}


<#
.SYNOPSIS
    Converts a DataTable to a flat array of PSCustomObjects with standardised property names.
.DESCRIPTION
    Maps source column names to a consistent set of output property names so that
    Show-PivotTable always receives the same shape of data regardless of query.
.PARAMETER DataTable
    The raw DataTable from Invoke-SqlQueryDirect.
.PARAMETER PropertyMap
    A hashtable mapping output names to source column names.
    Example: @{ User = "User"; Hour = "Hour"; Value = "Eaches" }
.OUTPUTS
    [PSCustomObject[]]
.EXAMPLE
    $results = ConvertTo-PivotData -DataTable $data -PropertyMap @{
        User = "User"; Hour = "Hour"; Value = "Eaches"
    }
    # Each object has .User, .Hour, .Value
#>
function ConvertTo-PivotData {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        $DataTable,

        [Parameter(Mandatory)]
        [hashtable]$PropertyMap
    )

    $results = @()
    foreach ($row in $DataTable) {
        $obj = [ordered]@{}
        foreach ($key in $PropertyMap.Keys) {
            $obj[$key] = $row.($PropertyMap[$key])
        }
        $results += [PSCustomObject]$obj
    }
    return $results
}


<#
.SYNOPSIS
    Converts every column of a DataTable to a PSCustomObject array, turning DBNull into real $null.
.DESCRIPTION
    SQL Server NULLs come back from a DataTable as [System.DBNull]::Value, not PowerShell $null.
    Most of this script's queries never surface a NULL in the value column they care about (they
    filter "IS NOT NULL" upstream), so it's never been an issue - but the Overview screen's combined
    query has columns that are genuinely NULL by design (e.g. RECEIVING never populates Totes), and
    Measure-Object throws "is not numeric" on a raw DBNull. This is the fix - use this instead of
    working off the raw DataTable when a query can have legitimately-NULL numeric columns.
.PARAMETER DataTable
    The raw DataTable from Invoke-SqlQueryDirect.
.OUTPUTS
    [PSCustomObject[]] - one object per row, one property per column, DBNull replaced with $null.
.EXAMPLE
    $rows = ConvertTo-CleanObjects -DataTable $data
    ($rows | Measure-Object -Property Totes -Sum).Sum   # no longer throws even if every Totes value was NULL
#>
function ConvertTo-CleanObjects {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [System.Data.DataTable]$DataTable
    )

    $columnNames = $DataTable.Columns | Select-Object -ExpandProperty ColumnName

    $results = @()
    foreach ($row in $DataTable) {
        $obj = [ordered]@{}
        foreach ($col in $columnNames) {
            $val = $row.$col
            $obj[$col] = if ($val -is [DBNull]) { $null } else { $val }
        }
        $results += [PSCustomObject]$obj
    }
    return $results
}



#endregion










#region sqlQueriesBuilder

<#
.SYNOPSIS
    Returns the shared #pick_task_base build SQL (with purge-order tagging).
.DESCRIPTION
    Single source of truth for pick_task_base. Every query function that needs
    task_type classification should splice this in rather than copy-pasting it.
    Update purge logic here ONCE and it propagates everywhere.
.NOTES
    2026-08-04 update: base_task_type now splits manual picking by source, matching
    Get-KPIOverviewQuery - 'pt_build_vna' -> 'MANUAL_PICK_VNA', 'pt_build_pr' ->
    'MANUAL_PICK_SPR'. Anything else that isn't GTP/Replenishment now falls into
    'MANUAL_PICK_ERR' instead of silently being lumped into a generic 'MANUAL_PICK'
    bucket. Callers that used to filter on task_type = 'MANUAL_PICK' need updating
    to 'MANUAL_PICK_VNA' and/or 'MANUAL_PICK_SPR' - see Get-ManualPickingPerformanceQuery.
.PARAMETER DateFilter
    The SQL expression to compare event_time against (e.g. "@targetDate" or
    a literal date string). Lets callers control whether they filter as a
    DECLAREd variable or a raw literal.
.OUTPUTS
    [string] - T-SQL block. Caller is responsible for DECLARE'ing/binding the
    date value it references before this runs.
#>
function Get-PickTaskBaseSql {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$TargetDate
    )

    return @"
DECLARE @targetDate date = '$TargetDate';

DROP TABLE IF EXISTS #pick_task_raw;
DROP TABLE IF EXISTS #purge_orders;
DROP TABLE IF EXISTS #pick_order;
DROP TABLE IF EXISTS #pick_task_base;

SELECT pick_task_id, oel_process_id, base_task_type
INTO #pick_task_raw
FROM (
    SELECT
        mpt.pick_task_id,
        mpt.oel_process_id,
        CASE
            WHEN mpt.oel_process_id = 'pt_build_gtp-dms' THEN 'GTP_PICK'
            WHEN mpt.oel_process_id = 'pt_build_pr' THEN 'MANUAL_PICK_SPR'
            WHEN mpt.oel_process_id = 'pt_build_vna' THEN 'MANUAL_PICK_VNA'
            WHEN mpt.oel_process_id LIKE 'pt_build_replen%' THEN 'REPLENISHMENT'
            ELSE 'MANUAL_PICK_ERR'
        END AS base_task_type,
        ROW_NUMBER() OVER (PARTITION BY mpt.pick_task_id ORDER BY mpt.event_time DESC) AS rn
    FROM mi_pick_task mpt
    WHERE mpt.change_uid IS NOT NULL
        AND mpt.oel_process_id IN (
            'pt_build_replen-from-vna', 'pt_build_replen-from-pr',
            'pt_build_replen-from-returns', 'pt_build_pr',
            'pt_build_gtp-dms', 'pt_build_vna'
        )
        AND mpt.event_time >= @targetDate
) x
WHERE rn = 1;
CREATE CLUSTERED INDEX IX_pick_task_raw ON #pick_task_raw (pick_task_id);

SELECT DISTINCT order_id
INTO #purge_orders
FROM mi_purge_order_processing;
CREATE CLUSTERED INDEX IX_purge_orders ON #purge_orders (order_id);

SELECT pick_task_id, order_id
INTO #pick_order
FROM (
    SELECT
        mp.pick_task_id,
        mp.order_id,
        ROW_NUMBER() OVER (PARTITION BY mp.pick_task_id ORDER BY mp.event_time DESC) AS rn
    FROM mi_pick mp
    WHERE mp.pick_task_id IN (SELECT pick_task_id FROM #pick_task_raw)
) y
WHERE rn = 1;
CREATE CLUSTERED INDEX IX_pick_order ON #pick_order (pick_task_id);

SELECT
    r.pick_task_id,
    r.oel_process_id,
    r.base_task_type + CASE WHEN po.order_id IS NOT NULL THEN '_PURGE' ELSE '' END AS task_type
INTO #pick_task_base
FROM #pick_task_raw r
LEFT JOIN #pick_order o ON o.pick_task_id = r.pick_task_id
LEFT JOIN #purge_orders po ON po.order_id = o.order_id;
"@
}

<#
.SYNOPSIS
    Returns the Fill Percentage SQL query.
.OUTPUTS
    [string] — T-SQL query.
.EXAMPLE
    $sql = Get-FillPercentageQuery
#>
function Get-FillPercentageQuery {
    [CmdletBinding()]
    [OutputType([string])]
    param()

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

<#
.SYNOPSIS
    Returns a Decant Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of three query variants:
      - EACHES  (UOM 1): SUM(quantity)       — individual item count
      - TOTES   (UOM 2): COUNT(DISTINCT case_id) — distinct tote count
      - CARTONS (UOM 3): COUNT(tote_id)      — carton/tote_id count
    Deduplicates events within a 5-minute window per tote_id + sku_id.
.PARAMETER TargetDate
    Date to query in "yyyy-MM-dd" format. Prompts user if not supplied.
.PARAMETER Uom
    The DECANT_UOM enum value. Default is TOTES (2). The user is prompted to
    change this at runtime.
.OUTPUTS
    [hashtable] with keys:
      - Query  [string]    — the T-SQL string
      - UOM    [DECANT_UOM] — the resolved unit of measure
.EXAMPLE
    # User is prompted for date + UOM interactively:
    $result = Get-DecantPerformanceQuery
    $result.Query   # → the SQL string
    $result.UOM     # → e.g. [DECANT_UOM]::TOTES

    # Explicit parameters (no prompts):
    $result = Get-DecantPerformanceQuery -TargetDate "2025-06-01" -Uom ([DECANT_UOM]::EACHES)
#>
function Get-DecantPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [DECANT_UOM]$Uom = [DECANT_UOM]::TOTES
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    # NR-NOTE: only prompt if the caller didn't already pass -Uom - lets this be called
    # headlessly (e.g. from the GUI) without blocking on Read-Host.
    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([DECANT_UOM])

        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [DECANT_UOM]$userSelection
        }
    }

    # Again, we have three different queries based on UOM selection. The main complexity is deduplicating events that occur within 5 minutes for the same tote_id + sku_id, which likely represent QA checks and release at the  decant space being recorded multiple times in WCS.

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

    $DECANT_QUERY_TOTES = @"
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
    
    $DECANT_QUERY_CARTONS = @"
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

    $query = switch ([int]$Uom) {
        1 { $DECANT_QUERY_EACHES }
        2 { $DECANT_QUERY_TOTES }
        3 { $DECANT_QUERY_CARTONS }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Totes." -ForegroundColor Yellow;
            $DECANT_QUERY_TOTES 
        }
    }

    return @{Query = $query; UOM = $Uom}


}


<#
.SYNOPSIS
    Returns a GTP Picking Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Totes is NOT supported):
      - EACHES  (UOM 1): SUM(qty)           — individual item count
      - CARTONS (UOM 3): COUNT(stock_tm_id)  — carton count
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    GTP_UOM enum value. Default EACHES (1). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [GTP_UOM]
.EXAMPLE
    $result = Get-GTPPickingQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [GTP_UOM]::EACHES

    $result = Get-GTPPickingQuery -TargetDate "2025-06-01" -Uom ([GTP_UOM]::CARTONS)
#>
function Get-GTPPickingQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [string]$targetDate,
        [GTP_UOM]$Uom = [GTP_UOM]::EACHES
    )
    
    if (-not $targetDate){$targetDate = Read-DateSelection}
    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM: $Uom (Totes is not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([GTP_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [GTP_UOM]$userSelection
        }
    }

    $GTP_QUERY_EACHES = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)

    SELECT 
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        SUM(qty) AS [Eaches]
    FROM 
        mi_pick
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND oel_class = 'OEL_PICK_PICKED'
        AND change_uid IS NOT NULL
        AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'GTP_PICK')
        AND qty IS NOT NULL
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    #You can try, but good luck. No point. We have a GTP through put calc.
    #Not value for an operator to know how many totes they picked, just how many eaches or outbound cartons.
    $GTP_QUERY_TOTES = @"
    NOT IN USE!!! Enjoy the error if you manage to call this...
"@

    #Why is it stock_tm_id and not tm_id? Who knows. WCS be like that
    <#ALSO to ADDD okay:
     Cartons query is picking up:
     CMC Cases
     Outbound Cartons (which is good, we want to count these)
     Cartons used for repack. (if we use tm_id instead we miss these, and miss batch picking!) Joys of WCS logging...
    #>
    $GTP_QUERY_CARTONS = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)
    
    SELECT 
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(stock_tm_id) AS [Eaches]
    FROM 
        mi_pick
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND oel_class = 'OEL_PICK_PICKED'
        AND change_uid IS NOT NULL
        AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'GTP_PICK')
        AND qty IS NOT NULL
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $GTP_QUERY_EACHES }
        2 { $GTP_QUERY_TOTES }
        3 { $GTP_QUERY_CARTONS }
        default { 
        Write-Host "Invalid UOM selection. No idea how lol Default to eaches." -ForegroundColor Yellow; 
        $GTP_QUERY_EACHES 
        }
    }

    return @{Query = $query; UOM = $Uom}
}


<#
.SYNOPSIS
    Returns a Replenishment Picking Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Eaches, and Totes is NOT supported):
      - CARTONS (UOM 3): COUNT(DISTINCT(case_id))  — carton count
      - PALLETS (UOM 4): COUNT(DISTINCT(pick_task_id))  — individual rough pallet count
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    REPLEN_UOM enum value. Default CARTONS (3). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [REPLEN_UOM]
.EXAMPLE
    $result = Get-ReplenishmentPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [REPLEN_UOM]::CARTONS

    $result = Get-ReplenishmentPerformanceQuery -TargetDate "2025-06-01" -Uom ([REPLEN_UOM]::CARTONS)
#>
function Get-ReplenishmentPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [REPLEN_UOM]$Uom = [REPLEN_UOM]::CARTONS
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Eaches and Totes not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([REPLEN_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [REPLEN_UOM]$userSelection
        }
    }

    # Again, we have three different queries based on UOM selection. The main complexity is deduplicating events that occur within 5 minutes for the same tote_id + sku_id, which likely represent QA checks and release at the  decant space being recorded multiple times in WCS.

    $REPLEN_QUERY_EACHES = @"
    Invalid Selection. Eaches not available. How did you even get this far?
"@

    $REPLEN_QUERY_TOTES = @"
    Invalid Selection. Totes not available. How did you even get this far?
"@
    
    $REPLEN_QUERY_CARTONS = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)


        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(case_id)) AS [Eaches]
        FROM mi_manual_full_case_picking
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_MANUAL_FULL_CASE_PICKING_PICK_COMPLETED'
            AND change_uid IS NOT NULL
            AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'REPLENISHMENT')
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $REPLEN_QUERY_PALLET = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)

        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(pick_task_id)) AS [Eaches]
        FROM mi_manual_full_case_picking
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_MANUAL_FULL_CASE_PICKING_PICK_COMPLETED'
            AND change_uid IS NOT NULL
            AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'REPLENISHMENT')
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $REPLEN_QUERY_EACHES }
        2 { $REPLEN_QUERY_TOTES }
        3 { $REPLEN_QUERY_CARTONS }
        4 { $REPLEN_QUERY_PALLET }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Cartons." -ForegroundColor Yellow;
            $REPLEN_QUERY_CARTONS 
        }
    }

    return @{Query = $query; UOM = $Uom}


}

<#
.SYNOPSIS
    Returns a Manual Picking Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Eaches, and Totes is NOT supported):
      - CARTONS (UOM 3): COUNT(DISTINCT(tm_id)) via mi_pick, filtered to manual pick tasks
      - PALLETS (UOM 4): COUNT(DISTINCT(pick_task_id)) via mi_manual_full_case_picking,
        filtered to replenishment tasks (kept as-is; not part of this update)
.NOTES
    2026-08-04 update: task_type classification (from Get-PickTaskBaseSql) now splits
    manual picking into 'MANUAL_PICK_VNA' and 'MANUAL_PICK_SPR' instead of one generic
    'MANUAL_PICK' bucket. The CARTONS variant now filters task_type IN ('MANUAL_PICK_VNA',
    'MANUAL_PICK_SPR') so this screen keeps showing combined manual-picking volume, same
    as before. If you'd rather see VNA and SPR broken out separately (like the KPI
    Overview screen does), split this into two query variants / two registry rows instead.
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    MANUALPICKING_UOM enum value. Default CARTONS (3). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [MANUALPICKING_UOM]
.EXAMPLE
    $result = Get-ManualPickingPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [MANUALPICKING_UOM]::CARTONS

    $result = Get-ManualPickingPerformanceQuery -TargetDate "2025-06-01" -Uom ([MANUALPICKING_UOM]::CARTONS)
#>
function Get-ManualPickingPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [MANUALPICKING_UOM]$Uom = [MANUALPICKING_UOM]::CARTONS
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Eaches and Totes not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([MANUALPICKING_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [MANUALPICKING_UOM]$userSelection
        }
    }

    # Again, we have three different queries based on UOM selection. The main complexity is deduplicating events that occur within 5 minutes for the same tote_id + sku_id, which likely represent QA checks and release at the  decant space being recorded multiple times in WCS.

    $MANUALPICKING_QUERY_EACHES = @"
    Invalid Selection. Eaches not available. How did you even get this far?
"@

    $MANUALPICKING_QUERY_TOTES = @"
    Invalid Selection. Totes not available. How did you even get this far?
"@
    
    $MANUALPICKING_QUERY_CARTONS = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)

        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(tm_id)) AS [Eaches]
        FROM mi_pick
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_PICK_PICKED'
            AND change_uid IS NOT NULL
            AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type IN ('MANUAL_PICK_VNA', 'MANUAL_PICK_SPR'))
            AND qty IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $MANUALPICKING_QUERY_PALLET = @"
    $(Get-PickTaskBaseSql -TargetDate $targetDate)

        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(pick_task_id)) AS [Eaches]
        FROM mi_manual_full_case_picking
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_MANUAL_FULL_CASE_PICKING_PICK_COMPLETED'
            AND change_uid IS NOT NULL
            AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'REPLENISHMENT')
            AND qty IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $MANUALPICKING_QUERY_EACHES }
        2 { $MANUALPICKING_QUERY_TOTES }
        3 { $MANUALPICKING_QUERY_CARTONS }
        4 { $MANUALPICKING_QUERY_PALLET }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Cartons." -ForegroundColor Yellow;
            $MANUALPICKING_QUERY_CARTONS 
        }
    }

    return @{Query = $query; UOM = $Uom}


}


<#
.SYNOPSIS
    Returns the Consumable Usage SQL query (last 7 days).
.OUTPUTS
    [string] — T-SQL query.
.EXAMPLE
    $sql = Get-ConsumableUsageQuery
#>

function Get-ConsumableUsageQuery {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    #ToDo -> Look through the TM. 

    return @"
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

}


<#
.SYNOPSIS
    Returns the Eaches Per Day SQL query.
.DESCRIPTION
    Groups picks by state and day. The day column is derived from either
    required_despatch_time (default) or priority_time.
.PARAMETER TargetDate
    Date filter. Prompts user if not supplied.
.PARAMETER UseDespatchTime
    If $true (default), groups by required_despatch_time.
    If $false, groups by priority_time.
.OUTPUTS
    [string] — T-SQL query.
.EXAMPLE
    # Default — group by despatch time:
    $sql = Get-EachesPerDayQuery                        

    # Group by priority time instead:
    $sql = Get-EachesPerDayQuery -UseDespatchTime $false
.NOTES
    We cannot use a targetDate filter here as it only stores 2 days of data.
    Just pull everything. Will need to adjust later to merge with another table.
#>
function Get-EachesPerDayQuery {
    [CmdletBinding()]
    [OutputType([string])]  

    param(
        #[string]$targetDate,
        [string]$useDespatchTime = $true
    )

    #if (-not $targetDate){$targetDate = Read-DateSelection}

    $filter = if ($useDespatchTime) { "required_despatch_time" } else { "priority_time" }

    return @"
    SELECT 
        pick_state AS [State],
        CONVERT(CHAR(10), $filter, 120) AS [DateDay],
        SUM(each_qty) AS [Eaches]
    FROM x_pick
    GROUP BY pick_state, CONVERT(CHAR(10), $filter, 120);
"@

}

<#
.SYNOPSIS
    Returns the GTP Utilisation SQL query (last 3 days).
.OUTPUTS
    [string] — T-SQL query.
#>
# NR 2026-07-09, this is a god awful implementation. IT is so laggy that I removed it from even being able to be selected in the new edition...
function Get-GTPUtilisationQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate,
        [GTPUTILISATION_UOM]$Uom = [GTPUTILISATION_UOM]::CARTON_CMC
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Eaches and Totes not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([GTPUTILISATION_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [GTPUTILISATION_UOM]$userSelection
        }
    }


    $GTPUTILISATION_QUERY_CARTONS_CMC = 
    @"
    SELECT
        DATEPART(HOUR, event_time) AS [Hour],
        'GTP' + RIGHT('0' + SUBSTRING(
            location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        ), 2) AS [User],
        COUNT(*) AS [Eaches]
    FROM mi_tm
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND change_field = 'location'
        AND (location LIKE 'GTP-PICK-%:01' OR location LIKE 'GTP-PUT-%:01')
        AND (tm_type = 'carton' OR (tm_type = 'TOTE' AND tm_id LIKE '7%'))
    GROUP BY
        DATEPART(HOUR, event_time),
        SUBSTRING(location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        )
    ORDER BY [Hour] ASC;
"@

  
    $GTPUTILISATION_QUERY_TOTES = @"
    SELECT
        DATEPART(HOUR, event_time) AS [Hour],
        'GTP' + RIGHT('0' + SUBSTRING(
            location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        ), 2) AS [User],
        COUNT(*) AS [Eaches]
    FROM mi_tm
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND change_field = 'location'
        AND (location LIKE 'GTP-PICK-%:01' OR location LIKE 'GTP-PUT-%:01')
        AND (tm_type = 'TOTE' AND tm_id NOT LIKE '7%')
    GROUP BY
        DATEPART(HOUR, event_time),
        SUBSTRING(location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        )
    ORDER BY [Hour] ASC;
"@

    $GTPUTILISATION_QUERY_COMBINED = @"
    SELECT
        DATEPART(HOUR, event_time) AS [Hour],
        'GTP' + RIGHT('0' + SUBSTRING(
            location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        ), 2) AS [User],
        COUNT(*) AS [Eaches]
    FROM mi_tm
    WHERE CAST(event_time AS DATE) = '$targetDate'
        AND change_field = 'location'
        AND (location LIKE 'GTP-PICK-%:01' OR location LIKE 'GTP-PUT-%:01')
    GROUP BY
        DATEPART(HOUR, event_time),
        SUBSTRING(location,
            LEN(location) - CHARINDEX('-', REVERSE(location)) + 1,
            CHARINDEX(':', location) - (LEN(location) - CHARINDEX('-', REVERSE(location)) + 1)
        )
    ORDER BY [Hour] ASC;
"@


    $query = switch([int]$Uom) {
        1 { $GTPUTILISATION_QUERY_CARTONS_CMC }
        2 { $GTPUTILISATION_QUERY_TOTES }
        3 { $GTPUTILISATION_QUERY_COMBINED }
        default { 
        Write-Host "Invalid UOM selection. No idea how lol Default to eaches." -ForegroundColor Yellow; 
        $GTPUTILISATION_QUERY_CARTONS_CMC 
        }
    }

    return @{Query = $query; UOM = $Uom}

}

<#
.SYNOPSIS
    Returns the Brand Distribution SQL query.
.OUTPUTS
    [string] — T-SQL query.
#>
function Get-BrandDistributionQuery {
    [CmdletBinding()]
    [OutputType([string])]
    param()

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
.SYNOPSIS
    Returns a Inbound - Putaway Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Totes is NOT supported):
      - EACHES  (UOM 1): SUM(quantity)       — individual item count
      - CARTONS (UOM 3): COUNT(DISTINCT(case_id))  — carton count
      - PALLETS (UOM 4): COUNT(DISTINCT(pallet_id))  — individual rough pallet count
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    PUTAWAY_UOM enum value. Default CARTONS (3). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [PUTAWAY_UOM]
.EXAMPLE
    $result = Get-PutawayPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [PUTAWAY_UOM]::CARTONS

    $result = Get-PutawayPerformanceQuery -TargetDate "2025-06-01" -Uom ([PUTAWAY_UOM]::CARTONS)
#>
function Get-PutawayPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [PUTAWAY_UOM]$Uom = [PUTAWAY_UOM]::CARTONS
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Totes not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([PUTAWAY_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [PUTAWAY_UOM]$userSelection
        }
    }

    $OEL_CLASS_FILTER_CASE = "OEL_CASE_PUTAWAY_CASE_STORED"  # For memory sake
    $OEL_CLASS_FILTER_PALLET = "OEL_CASE_PUTAWAY_COMPLETE"  # For memory sake: Remove qty != null filter as this event doesn't log quantity, but does log pallet_id which we can count distinct of for pallet performance.
    # Again, we have three different queries based on UOM selection. The main complexity is deduplicating events that occur within 5 minutes for the same tote_id + sku_id, which likely represent QA checks and release at the  decant space being recorded multiple times in WCS.
    # 2026-07-09. Why are there three different damn OEL class for the same action! Who if i could git blame I will...
    $PUTAWAY_QUERY_EACHES = @"
        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            SUM(quantity) AS [Eaches]
        FROM mi_case_putaway
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class IN ('OEL_CASE_PUTAWAY_CASE_STORED', 'OEL_CASE_PUTAWAY_COMPLETE')
            AND change_uid IS NOT NULL
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $PUTAWAY_QUERY_TOTES = @"
    Invalid Selection. Totes not available. How did you even get this far?
"@
    
    $PUTAWAY_QUERY_CARTONS = @"
        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(case_id)) AS [Eaches]
        FROM mi_case_putaway
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class IN ('OEL_CASE_PUTAWAY_CASE_STORED', 'OEL_CASE_PUTAWAY_COMPLETE')
            AND change_uid IS NOT NULL
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $PUTAWAY_QUERY_PALLET = @"
        SELECT
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(pallet_id)) AS [Eaches]
            FROM mi_case_putaway
            WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class IN ('OEL_CASE_PUTAWAY_CASE_STORED', 'OEL_CASE_PUTAWAY_COMPLETE')
            AND change_uid IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $PUTAWAY_QUERY_EACHES }
        2 { $PUTAWAY_QUERY_TOTES }
        3 { $PUTAWAY_QUERY_CARTONS }
        4 { $PUTAWAY_QUERY_PALLET }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Cartons." -ForegroundColor Yellow;
            $PUTAWAY_QUERY_CARTONS 
        }
    }

    return @{Query = $query; UOM = $Uom}


}


<#
.SYNOPSIS
    Returns a Inbound - Receiving Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Totes is NOT supported):
      - EACHES  (UOM 1): SUM(quantity)       — individual item count
      - CARTONS (UOM 3): COUNT(DISTINCT(case_id))  — carton count
      - PALLETS (UOM 4): COUNT(DISTINCT(pallet_id))  — individual rough pallet count
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    RECEVEING_UOM enum value. Default CARTONS (3). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [RECEVEING_UOM]
.EXAMPLE
    $result = Get-ReceivingPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [RECEVEING_UOM]::CARTONS

    $result = Get-ReceivingPerformanceQuery -TargetDate "2025-06-01" -Uom ([RECEVEING_UOM]::CARTONS)
#>
function Get-ReceivingPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [RECEIVING_UOM]$Uom = [RECEIVING_UOM]::CARTONS
    )
    #todo: Set a enum of default UOM?

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Totes not available)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([RECEIVING_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            # NR-NOTE: was [RECEVEING_UOM] - that type doesn't exist (typo of RECEIVING_UOM),
            # so this threw "Unable to find type" whenever someone actually picked a UOM here.
            $Uom = [RECEIVING_UOM]$userSelection
        }
    }

    # Again, we have three different queries based on UOM selection. 

    $RECEIVING_QUERY_EACHES = @"
        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            SUM(quantity) AS [Eaches]
        FROM mi_receiving
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_RECEIVING_CASE_RECEIVED'
            AND change_uid IS NOT NULL
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $RECEIVING_QUERY_TOTES = @"
    Invalid Selection. Totes not available. How did you even get this far?
"@
    
    $RECEIVING_QUERY_CARTONS = @"
        SELECT 
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(case_id)) AS [Eaches]
        FROM mi_receiving
        WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_RECEIVING_CASE_RECEIVED'
            AND change_uid IS NOT NULL
            AND quantity IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $RECEIVING_QUERY_PALLET = @"
        SELECT
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(DISTINCT(pallet_id)) AS [Eaches]
            FROM mi_receiving
            WHERE CAST(event_time AS DATE) = '$targetDate'
            AND oel_class = 'OEL_RECEIVING_CASE_RECEIVED'
            AND change_uid IS NOT NULL
        GROUP BY change_uid, DATEPART(HOUR, event_time)
        ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $RECEIVING_QUERY_EACHES }
        3 { $RECEIVING_QUERY_CARTONS }
        4 { $RECEIVING_QUERY_PALLET }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Cartons." -ForegroundColor Yellow;
            $RECEIVING_QUERY_CARTONS 
        }
    }

    return @{Query = $query; UOM = $Uom}


}

<#
.SYNOPSIS
    Returns a Inbound - Receiving Performance SQL query based on the chosen unit of measure.
.DESCRIPTION
    Builds one of two usable query variants (Totes is NOT supported):
      - EACHES  (UOM 1): SUM(quantity)       — individual item count
      - CARTONS (UOM 3): COUNT(DISTINCT(case_id))  — carton count
      - PALLETS (UOM 4): COUNT(DISTINCT(pallet_id))  — individual rough pallet count
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    RECEVEING_UOM enum value. Default CARTONS (3). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [RECEVEING_UOM]
.EXAMPLE
    $result = Get-ReceivingPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [RECEVEING_UOM]::CARTONS

    $result = Get-ReceivingPerformanceQuery -TargetDate "2025-06-01" -Uom ([RECEVEING_UOM]::CARTONS)
#>
function Get-InboundTrammingPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [INBOUND_TRAMMING_UOM]$Uom = [INBOUND_TRAMMING_UOM]::PALLET
    )
    

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Only Pallets!)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([INBOUND_TRAMMING_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [INBOUND_TRAMMING_UOM]$userSelection
        }
    }

    # Again, we have three different queries based on UOM selection. The main complexity is deduplicating events that occur within 5 minutes for the same tote_id + sku_id, which likely represent QA checks and release at the  decant space being recorded multiple times in WCS.

    $INBOUND_TRAMMING_QUERY_EACHES = @"
    Invalid RIP!
"@

    $INBOUND_TRAMMING_QUERY_TOTES = @"
    Invalid Selection. Totes not available. How did you even get this far?
"@
    
    $INBOUND_TRAMMING_QUERY_CARTONS = @"
    Invalid? But the default is loaded here???? I have done something wrong. Maybe fix, but dont know what.
"@

    $INBOUND_TRAMMING_QUERY_PALLET = @"
    SELECT
        pickup.change_uid AS [User],
        DATEPART(HOUR, pickup.event_time) AS [Hour],
        COUNT(pickup.tm_id) AS [Eaches]
    FROM mi_tramming AS pickup
    CROSS APPLY (
        SELECT TOP 1
            event_time,
            location
        FROM mi_tramming
        WHERE tm_id = pickup.tm_id
            AND change_uid = pickup.change_uid
            AND oel_class = 'OEL_TRAMMING_PALLET_DROPPED_OFF'
            AND event_time > pickup.event_time
        ORDER BY event_time ASC
    ) AS dropoff
    WHERE pickup.oel_class = 'OEL_TRAMMING_PALLET_PICKED_UP'
        AND CAST(pickup.event_time AS DATE) = '$targetDate'
        AND pickup.location LIKE 'RL%'
        AND dropoff.location NOT LIKE 'PR%'
        AND dropoff.location NOT LIKE 'S%'
        AND NOT EXISTS (
            SELECT 1
            FROM mi_tramming AS earlier
            WHERE earlier.tm_id = pickup.tm_id
                AND earlier.change_uid = pickup.change_uid
                AND earlier.location = pickup.location
                AND earlier.oel_class = 'OEL_TRAMMING_PALLET_PICKED_UP'
                AND earlier.event_time < pickup.event_time
                AND earlier.event_time >= DATEADD(SECOND, -120, pickup.event_time)
        )
    GROUP BY pickup.change_uid, DATEPART(HOUR, pickup.event_time)
    ORDER BY pickup.change_uid, DATEPART(HOUR, pickup.event_time);
"@

    $query = switch ([int]$Uom) {
        4 { $INBOUND_TRAMMING_QUERY_PALLET }
        default { 
            #todo: Hardcoding the default is wild. Later will update.
            Write-Host "Invalid UOM. Defaulting to Cartons." -ForegroundColor Yellow;
            $INBOUND_TRAMMING_QUERY_PALLET 
        }
    }

    return @{Query = $query; UOM = $Uom}

}

<#
.SYNOPSIS
    Returns a Returns Processing SQL query based on the chosen measure.
.DESCRIPTION
    Builds one of four query variants:
      - TOTAL_QTY    (UOM 1): COUNT of returned lines (saleable + non-saleable)
      - SALEABLE     (UOM 2): COUNT of lines returned saleable
      - NON_SALEABLE (UOM 3): COUNT of lines returned non-saleable
      - DELIVERIES   (UOM 4): COUNT(DISTINCT delivery_id)
    NOTE: Channel (ECOM vs WHS/RET) is NOT split out here - totals are combined
    across both channels. TODO: add a channel filter if that split is needed later.
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    RETURNS_UOM enum value. Default TOTAL_QTY (1). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [RETURNS_UOM]
.EXAMPLE
    $result = Get-ReturnsPerformanceQuery
    $result.Query   # - T-SQL string
    $result.UOM     # - [RETURNS_UOM]::TOTAL_QTY
#>
function Get-ReturnsPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param   (
        [string]$targetDate,
        [RETURNS_UOM]$Uom = [RETURNS_UOM]::TOTAL_QTY
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (channel split not available yet)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([RETURNS_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [RETURNS_UOM]$userSelection
        }
    }

    $RETURNS_QUERY_TOTAL_QTY = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        SUM(qty_total) AS [Eaches]
    FROM (
        SELECT
            change_uid,
            event_time,
            delivery_id,
            COUNT(sku_id) AS qty_total
        FROM mi_returns_processing
        WHERE oel_class IN (
                'OEL_RETURNS_PROCESSING_ITEM_RETURNED_SALEABLE',
                'OEL_RETURNS_PROCESSING_ITEM_RETURNED_NON_SALEABLE'
            )
            AND delivery_id LIKE 'PVHHIVE%'
            AND change_uid IS NOT NULL
            AND CAST(event_time AS DATE) = '$targetDate'
        GROUP BY change_uid, event_time, delivery_id
    ) delivery_detail
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $RETURNS_QUERY_SALEABLE = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(*) AS [Eaches]
    FROM mi_returns_processing
    WHERE oel_class = 'OEL_RETURNS_PROCESSING_ITEM_RETURNED_SALEABLE'
        AND delivery_id LIKE 'PVHHIVE%'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $RETURNS_QUERY_NON_SALEABLE = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(*) AS [Eaches]
    FROM mi_returns_processing
    WHERE oel_class = 'OEL_RETURNS_PROCESSING_ITEM_RETURNED_NON_SALEABLE'
        AND delivery_id LIKE 'PVHHIVE%'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $RETURNS_QUERY_DELIVERIES = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(DISTINCT delivery_id) AS [Eaches]
    FROM mi_returns_processing
    WHERE oel_class IN (
            'OEL_RETURNS_PROCESSING_ITEM_RETURNED_SALEABLE',
            'OEL_RETURNS_PROCESSING_ITEM_RETURNED_NON_SALEABLE'
        )
        AND delivery_id LIKE 'PVHHIVE%'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $RETURNS_QUERY_TOTAL_QTY }
        2 { $RETURNS_QUERY_SALEABLE }
        3 { $RETURNS_QUERY_NON_SALEABLE }
        4 { $RETURNS_QUERY_DELIVERIES }
        default {
            Write-Host "Invalid UOM. Defaulting to Total Qty." -ForegroundColor Yellow;
            $RETURNS_QUERY_TOTAL_QTY
        }
    }

    return @{Query = $query; UOM = $Uom}
}


<#
.SYNOPSIS
    Returns the CUBI (dimensioning completed) SQL query for a single day.
.DESCRIPTION
    Single measure only - distinct SKUs dimensioned per user per hour.
.NOTES
    NR-NOTE: the source SQL this was lifted from filtered `change_uid IS NULL`,
    which would blank the [User] column entirely and break this dashboard's
    per-user breakdown. Flipped to IS NOT NULL here - verify this matches intent
    before relying on it.
    Update to the note 2026-07-09. Reason we only wanted null was adjustments done by OM will have a change_uid,
    but the original dimensioning was done by the cubi machine, which is a api interface, will have a null change_uid. 
    We are calculating how efficent users are at the cubi machine, so we only want to count the null change_uid events...
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [string] (fixed label, no UOM choice)
.EXAMPLE
    $result = Get-CubiPerformanceQuery
#>
function Get-CubiPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    $query = @"
    SELECT
        'NOT RECORDED IN WCS' AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(DISTINCT (sku_id)) AS [Eaches]
    FROM mi_sku_uom
    WHERE oel_class = 'OEL_SKU_UOM_FIELD_CHANGE'
        AND change_field = 'Dimensioning Complete'
        AND to_value = 'YES'
        AND change_uid IS NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    return @{Query = $query; UOM = "SKUs Dimensioned"}
}


<#
.SYNOPSIS
    Returns a VAS (Value Added Services) Performance SQL query based on the chosen measure.
.DESCRIPTION
    Cross-checks completed VAS activity against the originating pick task to
    pull a quantity. This is VAS at the Zone Route Area Builds one of two query variants:
      - EACHES  (UOM 1): SUM(pt.total_pick_qty)      - item qty across the VAS'd task
      - CARTONS (UOM 3): COUNT(DISTINCT v.tm_id)      - cartons that went through VAS
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    VAS_UOM enum value. Default EACHES (1). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [VAS_UOM]
.EXAMPLE
    $result = Get-VASPerformanceQuery
#>
function Get-VASPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate,
        [VAS_UOM]$Uom = [VAS_UOM]::EACHES
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (Totes not applicable)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([VAS_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [VAS_UOM]$userSelection
        }
    }

    $VAS_QUERY_EACHES = @"
    SELECT
        v.change_uid AS [User],
        DATEPART(HOUR, v.event_time) AS [Hour],
        SUM(pt.total_pick_qty) AS [Eaches]
    FROM mi_post_vas v
    INNER JOIN mi_pick_task pt ON pt.tm_id = v.tm_id
    WHERE v.oel_class = 'OEL_POST_VAS_ACTIVITY_COMPLETE'
        AND v.change_uid IS NOT NULL
        AND CAST(v.event_time AS DATE) = '$targetDate'
        AND pt.oel_class = 'OEL_PICK_TASK_STATE_CHANGE'
        AND pt.change_field = 'Pick Task State'
        AND pt.from_value = 'STARTED'
    GROUP BY v.change_uid, DATEPART(HOUR, v.event_time)
    ORDER BY v.change_uid, DATEPART(HOUR, v.event_time);
"@

    $VAS_QUERY_CARTONS = @"

    SELECT
        v.change_uid AS [User],
        DATEPART(HOUR, v.event_time) AS [Hour],
        COUNT(DISTINCT (v.tm_id)) AS [Eaches]
    FROM mi_post_vas v
    INNER JOIN mi_pick_task pt ON pt.tm_id = v.tm_id
    WHERE v.oel_class = 'OEL_POST_VAS_ACTIVITY_COMPLETE'
        AND v.change_uid IS NOT NULL
        AND CAST(v.event_time AS DATE) = '$targetDate'
        AND pt.oel_class = 'OEL_PICK_TASK_STATE_CHANGE'
        AND pt.change_field = 'Pick Task State'
        AND pt.from_value = 'STARTED'
    GROUP BY v.change_uid, DATEPART(HOUR, v.event_time)
    ORDER BY v.change_uid, DATEPART(HOUR, v.event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $VAS_QUERY_EACHES }
        3 { $VAS_QUERY_CARTONS }
        default {
            Write-Host "Invalid UOM. Defaulting to Eaches." -ForegroundColor Yellow;
            $VAS_QUERY_EACHES
        }
    }

    return @{Query = $query; UOM = $Uom}
}


<#
.SYNOPSIS
    Returns a QI Check Performance SQL query based on the chosen measure.
.DESCRIPTION
    Builds one of two query variants against the 'QI PALLET' station:
      - SKUS    (UOM 1): COUNT(DISTINCT sku_id)  - distinct SKUs checked
      - PALLETS (UOM 4): COUNT(DISTINCT tm_id)   - pallets checked
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.PARAMETER Uom
    QI_UOM enum value. Default SKUS (1). User is prompted to change.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [QI_UOM]
.EXAMPLE
    $result = Get-QIPerformanceQuery
#>
function Get-QIPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate,
        [QI_UOM]$Uom = [QI_UOM]::SKUS
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    if (-not $PSBoundParameters.ContainsKey('Uom')) {
        Write-Host "Current default UOM selection: $Uom (QI PALLET station only)" -ForegroundColor Green
        $userSelection = Read-EnumSelection -EnumType ([QI_UOM])
        if (Test-ValidEnumSelection -UserSelection $userSelection) {
            $Uom = [QI_UOM]$userSelection
        }
    }

    $QI_QUERY_SKUS = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(DISTINCT (sku_id)) AS [Eaches]
    FROM mi_qa
    WHERE oel_class = 'OEL_QA_STOCK_RELEASED'
        AND station_id = 'QI PALLET'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $QI_QUERY_PALLETS = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(DISTINCT (tm_id)) AS [Eaches]
    FROM mi_qa
    WHERE oel_class = 'OEL_QA_STOCK_RELEASED'
        AND station_id = 'QI PALLET'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    $query = switch ([int]$Uom) {
        1 { $QI_QUERY_SKUS }
        4 { $QI_QUERY_PALLETS }
        default {
            Write-Host "Invalid UOM. Defaulting to SKUs." -ForegroundColor Yellow;
            $QI_QUERY_SKUS
        }
    }

    return @{Query = $query; UOM = $Uom}
}


<#
.SYNOPSIS
    Returns the Parcel Repack Performance SQL query for a single day.
.DESCRIPTION
    Single measure only - count of repacked orders per user per hour.
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [string] (fixed label, no UOM choice)
.EXAMPLE
    $result = Get-ParcelRepackPerformanceQuery
#>
function Get-ParcelRepackPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    $query = @"
    WITH combined AS (
        SELECT
            change_uid AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(order_id) AS [Eaches]
        FROM mi_repack
        WHERE oel_class = 'OEL_REPACK_ORDER_PACKED'
            AND change_uid IS NOT NULL
            AND CAST(event_time AS DATE) = '2026-07-10'
        GROUP BY change_uid, DATEPART(HOUR, event_time)

        UNION ALL

        SELECT
            'SYSTEM' AS [User],
            DATEPART(HOUR, event_time) AS [Hour],
            COUNT(carton_id) AS [Eaches]
        FROM mi_cmc_packing
        WHERE oel_class = 'OEL_CMC_PACKING_ORDER_PACKED'
            AND change_uid IS NOT NULL
            AND CAST(event_time AS DATE) = '2026-07-10'
        GROUP BY change_uid, DATEPART(HOUR, event_time)
    )
    SELECT
        [User],
        [Hour],
        SUM([Eaches]) AS [Eaches]
    FROM combined
    GROUP BY [User], [Hour]
    ORDER BY [User], [Hour];
"@

    return @{Query = $query; UOM = "Orders Repacked"}
}


<#
.SYNOPSIS
    Returns the Despatch Performance SQL query for a single day.
.DESCRIPTION
    Single measure only - cartons despatched per user per hour.
.NOTES
    2026-08-04 update: switched source from mi_despatch_sorting (which had no
    oel_class filter, hardcoded [User] to the literal 'SYSTEM', and grouped by a
    change_uid it never selected) to mi_trailer_loading, filtered to
    OEL_TRAILER_LOADING_CARTON_LOADED, same as Get-KPIOverviewQuery's DESPATCH
    section. [User] is now the real change_uid, so this screen shows genuine
    per-user despatch activity instead of one lumped 'SYSTEM' row per hour.
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.OUTPUTS
    [hashtable] with keys: Query [string], UOM [string] (fixed label, no UOM choice)
.EXAMPLE
    $result = Get-DespatchPerformanceQuery
#>
function Get-DespatchPerformanceQuery {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [string]$targetDate
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    $query = @"
    SELECT
        change_uid AS [User],
        DATEPART(HOUR, event_time) AS [Hour],
        COUNT(DISTINCT tm_id) AS [Eaches]
    FROM mi_trailer_loading
    WHERE oel_class = 'OEL_TRAILER_LOADING_CARTON_LOADED'
        AND change_uid IS NOT NULL
        AND CAST(event_time AS DATE) = '$targetDate'
    GROUP BY change_uid, DATEPART(HOUR, event_time)
    ORDER BY change_uid, DATEPART(HOUR, event_time);
"@

    return @{Query = $query; UOM = "Cartons Despatched"}
}


<#
.SYNOPSIS
    Returns the "All Activities" overview SQL query for a single day.
.DESCRIPTION
    Adapted from the standalone multi-activity extract script (originally a rolling
    3-month window feeding a report/export) down to a single target date, so it can
    run inside the interactive dashboard without hammering the server on every load.
    Builds all activity temp tables (Receiving, Returns, CUBI, VAS, Despatch, Putaway,
    QI Check, Decant, GTP Pick, Demech Pick, Manual Pick (VNA), Manual Pick SPR, Parcel
    Repack, Replenishment) for the one
    day, then UNIONs them into a single Activity/Hour/User result set with every
    metric column (Eaches/Cartons/Pallets/Totes/Cases/Carton_CMC/Combined) - most
    activities only ever populate a subset of those columns, the rest stay NULL.
.NOTES
    Same NR-NOTE as Get-CubiPerformanceQuery: the CUBI change_uid filter was flipped
    from IS NULL to IS NOT NULL here too, for the same reason.

    2026-08-04 update: Manual Picking is now split by task source - VNA-sourced picks
    still surface under the 'MANUAL_PICK' Activity (unchanged label, so existing
    consumers of that name keep working), and PR/SPR-sourced picks now surface
    separately under a new 'MANUAL_PICK_SPR' Activity. Despatch now reads from
    mi_trailer_loading (OEL_TRAILER_LOADING_CARTON_LOADED) with the real change_uid,
    instead of mi_despatch_sorting with a hardcoded 'SYSTEM' user.
.PARAMETER TargetDate
    Date to query. Prompts user if not supplied.
.OUTPUTS
    [string] - T-SQL query. Returns columns: Activity, DATE, HOUR, USER, Eaches,
    Cartons, Pallets, Totes, Cases, Carton_CMC, Combined.
.EXAMPLE
    $sql = Get-KPIOverviewQuery -TargetDate "2025-06-01"
#>
function Get-KPIOverviewQuery {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string]$targetDate
    )

    if (-not $targetDate){$targetDate = Read-DateSelection}

    return @"
DECLARE @targetDate date = '$targetDate';
DECLARE @nextDate   date = DATEADD(DAY, 1, @targetDate);   -- ADDED: gives every section a clean upper bound

DROP TABLE IF EXISTS #receiving_base;
DROP TABLE IF EXISTS #decant_base;
DROP TABLE IF EXISTS #decant_deduped;
DROP TABLE IF EXISTS #replen_base;
DROP TABLE IF EXISTS #pick_base;
DROP TABLE IF EXISTS #manual_pick_base;
DROP TABLE IF EXISTS #manual_pick_base_spr;
DROP TABLE IF EXISTS #putaway_base;
DROP TABLE IF EXISTS #repacking_base;
DROP TABLE IF EXISTS #despatch_base;
DROP TABLE IF EXISTS #returns_base;
DROP TABLE IF EXISTS #VAS_work;
DROP TABLE IF EXISTS #VAS_base;
DROP TABLE IF EXISTS #cubi_base;
DROP TABLE IF EXISTS #qi_base;

/* Tables for purge order :? Don't touch please.*/
DROP TABLE IF EXISTS #pick_task_raw;
DROP TABLE IF EXISTS #purge_orders;
DROP TABLE IF EXISTS #pick_order;
DROP TABLE IF EXISTS #pick_task_base;
DROP TABLE IF EXISTS #pick_purge_base;

SELECT pick_task_id, oel_process_id, base_task_type
INTO #pick_task_raw
FROM (
    SELECT
        mpt.pick_task_id,
        mpt.oel_process_id,
        CASE
            WHEN mpt.oel_process_id = 'pt_build_gtp-dms' THEN 'GTP_PICK'
			WHEN mpt.oel_process_id = 'pt_build_pr' THEN 'MANUAL_PICK_SPR'
			WHEN mpt.oel_process_id = 'pt_build_vna' THEN 'MANUAL_PICK_VNA'
            WHEN mpt.oel_process_id LIKE 'pt_build_replen%' THEN 'REPLENISHMENT'
            ELSE 'MANUAL_PICK_ERR'
        END AS base_task_type,
        ROW_NUMBER() OVER (PARTITION BY mpt.pick_task_id ORDER BY mpt.event_time DESC) AS rn
    FROM mi_pick_task mpt
    WHERE mpt.change_uid IS NOT NULL
        AND mpt.oel_process_id IN (
            'pt_build_replen-from-vna', 
			'pt_build_replen-from-pr',
            'pt_build_replen-from-returns', 
			'pt_build_pr',
            'pt_build_gtp-dms', 
			'pt_build_vna'
        )
        AND mpt.event_time >= @targetDate
        AND mpt.event_time < @nextDate                      -- CHANGED: was open-ended, now one day
) x
WHERE rn = 1;
CREATE CLUSTERED INDEX IX_pick_task_raw ON #pick_task_raw (pick_task_id);

SELECT DISTINCT order_id
INTO #purge_orders
FROM mi_purge_order_processing;
CREATE CLUSTERED INDEX IX_purge_orders ON #purge_orders (order_id);

SELECT pick_task_id, order_id
INTO #pick_order
FROM (
    SELECT
        mp.pick_task_id,
        mp.order_id,
        ROW_NUMBER() OVER (PARTITION BY mp.pick_task_id ORDER BY mp.event_time DESC) AS rn
    FROM mi_pick mp
    WHERE mp.pick_task_id IN (SELECT pick_task_id FROM #pick_task_raw)
) y
WHERE rn = 1;
CREATE CLUSTERED INDEX IX_pick_order ON #pick_order (pick_task_id);

SELECT
    r.pick_task_id,
    r.oel_process_id,
    r.base_task_type + CASE WHEN po.order_id IS NOT NULL THEN '_PURGE' ELSE '' END AS task_type
INTO #pick_task_base
FROM #pick_task_raw r
LEFT JOIN #pick_order o ON o.pick_task_id = r.pick_task_id
LEFT JOIN #purge_orders po ON po.order_id = o.order_id;

-- ADDED: #pick_task_base is scanned by task_type four times below with no index at all
CREATE CLUSTERED INDEX IX_pick_task_base ON #pick_task_base (pick_task_id);
CREATE NONCLUSTERED INDEX IX_pick_task_base_type ON #pick_task_base (task_type) INCLUDE (pick_task_id);


/* SEC TABLE BUILDING */

/* SEC RECEIVING */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, quantity, case_id, pallet_id
INTO #receiving_base
FROM mi_receiving
WHERE oel_class = 'OEL_RECEIVING_CASE_RECEIVED'
    AND change_uid IS NOT NULL
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable

/* CUBI */
SELECT
    CAST(event_time AS DATE) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, sku_id
INTO #cubi_base
FROM mi_sku_uom
WHERE oel_class = 'OEL_SKU_UOM_FIELD_CHANGE'
  AND change_field = 'Dimensioning Complete'
  AND to_value = 'YES'
  AND change_uid IS NOT NULL
  AND event_time >= @targetDate AND event_time < @nextDate;     -- CHANGED: sargable

/* QI CHECKS BEEN COMPLETED */
SELECT
    CAST(event_time AS DATE) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, sku_id, tm_id
INTO #qi_base
FROM mi_qa
WHERE oel_class = 'OEL_QA_STOCK_RELEASED'
  AND station_id = 'QI PALLET'
  AND change_uid IS NOT NULL
  AND event_time >= @targetDate AND event_time < @nextDate;     -- CHANGED: sargable


/* VAS */
SELECT
    CAST(event_time AS DATE) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, tm_id
INTO #VAS_work
FROM mi_post_vas
WHERE oel_class = 'OEL_POST_VAS_ACTIVITY_COMPLETE'
  AND change_uid IS NOT NULL
  AND event_time >= @targetDate AND event_time < @nextDate;     -- CHANGED: sargable

SELECT
    v.txn_date, v.hour_num, v.change_uid,
    COUNT(DISTINCT v.tm_id) AS cartons_done,
    SUM(pt.total_pick_qty)  AS total_qty
INTO #VAS_base
FROM #VAS_work v
INNER JOIN mi_pick_task pt ON pt.tm_id = v.tm_id
WHERE pt.oel_class = 'OEL_PICK_TASK_STATE_CHANGE'
  AND pt.change_field = 'Pick Task State'
  AND pt.from_value = 'STARTED'
GROUP BY v.txn_date, v.hour_num, v.change_uid
ORDER BY v.txn_date, v.hour_num, v.change_uid;


/* RETURNS */
SELECT
    CAST(event_time AS DATE) AS txn_date,
    DATEPART(HOUR, event_time) as hour_num,
    change_uid, channel,
    COUNT(DISTINCT delivery_id) AS distinct_deliveries,
    SUM(qty_total)               AS qty_total,
    SUM(qty_saleable)            AS qty_saleable,
    SUM(qty_non_saleable)        AS qty_non_saleable
INTO #returns_base
FROM (
    SELECT
        change_uid, event_time, delivery_id,
        CASE
            WHEN order_id LIKE 'C%' OR order_id LIKE 'V%'
              OR order_id LIKE 'T%' OR order_id LIKE 'IC%'
            THEN 'ECOM' ELSE 'WHS/RET'
        END AS channel,
        COUNT(sku_id) AS qty_total,
        SUM(CASE WHEN oel_class = 'OEL_RETURNS_PROCESSING_ITEM_RETURNED_SALEABLE' THEN 1 ELSE 0 END) AS qty_saleable,
        SUM(CASE WHEN oel_class = 'OEL_RETURNS_PROCESSING_ITEM_RETURNED_NON_SALEABLE' THEN 1 ELSE 0 END) AS qty_non_saleable
    FROM mi_returns_processing
    WHERE oel_class IN (
            'OEL_RETURNS_PROCESSING_ITEM_RETURNED_SALEABLE',
            'OEL_RETURNS_PROCESSING_ITEM_RETURNED_NON_SALEABLE'
        )
        AND delivery_id LIKE 'PVHHIVE%'
        AND change_uid IS NOT NULL
        AND event_time >= @targetDate AND event_time < @nextDate   -- CHANGED: sargable
    GROUP BY
        event_time, change_uid, delivery_id,
        CASE
            WHEN order_id LIKE 'C%' OR order_id LIKE 'V%'
              OR order_id LIKE 'T%' OR order_id LIKE 'IC%'
            THEN 'ECOM' ELSE 'WHS/RET'
        END
) delivery_detail
GROUP BY CAST(event_time AS DATE), DATEPART(HOUR, event_time), change_uid, channel;   -- ADDED: trailing semicolon

/* SEC DEMECH PICKING */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, order_id, oel_class, from_value, pick_state,
    location, qty, stock_tm_id, case_qty, each_qty, pick_task_id
INTO #pick_purge_base
FROM mi_pick
WHERE change_uid IS NOT NULL
    AND oel_class = 'OEL_PICK_PICKED'
    AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'GTP_PICK_PURGE')
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: was unbounded, now one day like everything else


/* SEC VNA MANUAL PICKING */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, order_id, oel_class, from_value, pick_state,
    location, qty, stock_tm_id, case_qty, each_qty, pick_task_id
INTO #manual_pick_base
FROM mi_pick
WHERE change_uid IS NOT NULL
    AND oel_class = 'OEL_PICK_PICKED'
    AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'MANUAL_PICK_VNA')
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable

/* SEC PR MANUAL PICKING */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, order_id, oel_class, from_value, pick_state,
    location, qty, stock_tm_id, case_qty, each_qty, pick_task_id
INTO #manual_pick_base_spr
FROM mi_pick
WHERE change_uid IS NOT NULL
    AND oel_class = 'OEL_PICK_PICKED'
    AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'MANUAL_PICK_SPR')
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable

/* SEC GTP PICKING */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, order_id, oel_class, from_value, pick_state,
    location, qty, stock_tm_id, case_qty, each_qty, pick_task_id
INTO #pick_base
FROM mi_pick
WHERE change_uid IS NOT NULL
    AND oel_class = 'OEL_PICK_PICKED'
    AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'GTP_PICK')
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable


/* SEC RETRIEVAL OR REPLENISHMENT */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, case_id, pick_task_id, quantity
INTO #replen_base
FROM mi_manual_full_case_picking
WHERE oel_class = 'OEL_MANUAL_FULL_CASE_PICKING_PICK_COMPLETED'
    AND change_uid IS NOT NULL
    AND quantity IS NOT NULL
    AND pick_task_id IN (SELECT pick_task_id FROM #pick_task_base WHERE task_type = 'REPLENISHMENT')
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable


/* SEC DECANT */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, quantity, case_id, tote_id, sku_id, event_time,
    LAG(event_time) OVER (PARTITION BY tote_id, sku_id ORDER BY event_time) AS prev_event_time
INTO #decant_base
FROM mi_decant
WHERE oel_class = 'OEL_DECANT_STOCK_TOTE_COMPLETED'
    AND change_uid IS NOT NULL
    AND quantity IS NOT NULL
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable

SELECT txn_date, hour_num, change_uid, quantity, case_id, tote_id
INTO #decant_deduped
FROM #decant_base
WHERE prev_event_time IS NULL
    OR DATEDIFF(SECOND, prev_event_time, event_time) > 300;


/* SEC PACKING */
WITH combined AS (
        --This is repacking
		SELECT
			CAST(event_time AS date) AS txn_date,
            DATEPART(HOUR, event_time) AS hour_num,
            change_uid AS [User],
            order_id
        FROM mi_repack
        WHERE oel_class = 'OEL_REPACK_ORDER_PACKED'
            AND change_uid IS NOT NULL
			AND event_time >= @targetDate AND event_time < @nextDate
		--OLD DEBUGGROUP BY change_uid, DATEPART(HOUR, event_time)

        UNION ALL

		--This is CMC
        SELECT
			CAST(event_time AS date) AS txn_date,
            DATEPART(HOUR, event_time) AS hour_num,
            'SYSTEM' AS [User],
            carton_id AS order_id
        FROM mi_cmc_packing
        WHERE oel_class = 'OEL_CMC_PACKING_ORDER_PACKED'
            AND change_uid IS NOT NULL
			AND event_time >= @targetDate AND event_time < @nextDate
            --AND CAST(event_time AS DATE) = '2026-07-10' --OLD for debugging lol
        --GROUP BY change_uid, DATEPART(HOUR, event_time)
    )
    SELECT
        txn_date,
        hour_num,
        [User] AS change_uid,
        COUNT(order_id) AS Eaches
	INTO #repacking_base
    FROM combined
	GROUP BY txn_date, hour_num, [User]
    ORDER BY txn_date, hour_num, [User];



/* SEC PUTAWAY */
SELECT
    CAST(event_time AS date) AS txn_date,
    DATEPART(HOUR, event_time) AS hour_num,
    change_uid, quantity, case_id, pallet_id, oel_class
INTO #putaway_base
FROM mi_case_putaway
WHERE oel_class IN ('OEL_CASE_PUTAWAY_CASE_STORED', 'OEL_CASE_PUTAWAY_COMPLETE')
    AND change_uid IS NOT NULL
    AND event_time >= @targetDate AND event_time < @nextDate;   -- CHANGED: sargable


/* SEC DESPATCH */
/*** NATHANIEL THIS NEEDS TO BE UPDATED FOR EACHES ***/
SELECT 
	CAST(event_time AS DATE) as txn_date,
	DATEPART(HOUR, event_time) AS hour_num,
	change_uid,
	COUNT(DISTINCT tm_id) AS CARTONS_DESPATCHED
INTO #despatch_base
FROM mi_trailer_loading
WHERE 
	1=1
	AND oel_class = 'OEL_TRAILER_LOADING_CARTON_LOADED'
	AND event_time >= @targetDate AND event_time < @nextDate      -- CHANGED: sargable
GROUP BY CAST(event_time AS date), DATEPART(HOUR, event_time), change_uid

/* SEC INDEX CREATION */
CREATE INDEX IX_receiving_base ON #receiving_base (txn_date, hour_num, change_uid) INCLUDE (quantity, case_id, pallet_id);
CREATE INDEX IX_putaway_base ON #putaway_base (txn_date, hour_num, change_uid, oel_class) INCLUDE (quantity, case_id, pallet_id);
CREATE INDEX IX_decant_deduped ON #decant_deduped (txn_date, hour_num, change_uid) INCLUDE (quantity, case_id, tote_id);
CREATE INDEX IX_pick_base_group ON #pick_base (txn_date, hour_num, change_uid, oel_class) INCLUDE (qty, stock_tm_id, each_qty);
CREATE INDEX IX_replen_base ON #replen_base (txn_date, hour_num, change_uid) INCLUDE (case_id, pick_task_id, quantity);
CREATE INDEX IX_manual_pick_base ON #manual_pick_base (oel_class, txn_date, hour_num, change_uid) INCLUDE (case_qty, pick_task_id);
CREATE INDEX IX_despatch_base ON #despatch_base (txn_date, hour_num, change_uid) INCLUDE (CARTONS_DESPATCHED);   -- CHANGED: order_id doesn't exist on this table; index the column actually used


/* SEC PRESENTING DATA */
WITH combined AS (
    SELECT 'RECEIVING' AS Activity, txn_date, hour_num, change_uid AS [USER],
        SUM(CAST(quantity AS decimal(18,2))) AS Eaches,
        CAST(COUNT(DISTINCT case_id) AS decimal(18,2)) AS Cartons,
        CAST(COUNT(DISTINCT pallet_id) AS decimal(18,2)) AS Pallets,
        CAST(NULL AS decimal(18,2)) AS Totes, CAST(NULL AS decimal(18,2)) AS Cases,
        CAST(NULL AS decimal(18,2)) AS Carton_CMC, CAST(NULL AS decimal(18,2)) AS Combined
    FROM #receiving_base WHERE quantity IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'RETURNS', txn_date, hour_num, change_uid,
        SUM(CAST(qty_total AS decimal(18,2))), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #returns_base WHERE qty_total IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'CUBI', txn_date, hour_num, change_uid,
        CAST(COUNT(DISTINCT sku_id) AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #cubi_base
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'VAS', txn_date, hour_num, change_uid,
        SUM(CAST(total_qty AS decimal(18,2))), SUM(CAST(cartons_done AS decimal(18,2))), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #VAS_base WHERE total_qty IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'DESPATCH', txn_date, hour_num, change_uid,
        CAST(NULL AS decimal(18,2)), CAST(SUM(CARTONS_DESPATCHED) AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #despatch_base
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'PUTAWAY', txn_date, hour_num, change_uid,
        SUM(CASE WHEN oel_class = 'OEL_CASE_PUTAWAY_CASE_STORED' THEN CAST(quantity AS decimal(18,2)) END),
        CAST(COUNT(DISTINCT CASE WHEN oel_class = 'OEL_CASE_PUTAWAY_CASE_STORED' AND quantity IS NOT NULL THEN case_id END) AS decimal(18,2)),
        CAST(COUNT(DISTINCT CASE WHEN oel_class = 'OEL_CASE_PUTAWAY_COMPLETE' AND pallet_id IS NOT NULL THEN pallet_id END) AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #putaway_base
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'QI_CHECK', txn_date, hour_num, change_uid,
        CAST(COUNT(DISTINCT sku_id) AS decimal(18,2)), CAST(COUNT(DISTINCT tm_id) AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #qi_base
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'DECANT', txn_date, hour_num, change_uid,
        SUM(CAST(quantity AS decimal(18,2))), CAST(COUNT(DISTINCT case_id) AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(COUNT(tote_id) AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #decant_deduped
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'GTP_PICK', txn_date, hour_num, change_uid,
        SUM(CAST(each_qty AS decimal(18,2))), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #pick_base WHERE oel_class = 'OEL_PICK_PICKED' AND qty IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'DEMECH_PICK', txn_date, hour_num, change_uid,
        SUM(CAST(each_qty AS decimal(18,2))), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #pick_purge_base WHERE oel_class = 'OEL_PICK_PICKED' AND qty IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'MANUAL_PICK', txn_date, hour_num, change_uid,
        CAST(NULL AS decimal(18,2)), SUM(CAST(case_qty AS decimal(18,2))), CAST(COUNT(DISTINCT pick_task_id) AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #manual_pick_base WHERE oel_class = 'OEL_PICK_PICKED' AND qty IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

	UNION ALL
    SELECT 'MANUAL_PICK_SPR', txn_date, hour_num, change_uid,
        CAST(NULL AS decimal(18,2)), SUM(CAST(case_qty AS decimal(18,2))), CAST(COUNT(DISTINCT pick_task_id) AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #manual_pick_base_spr WHERE oel_class = 'OEL_PICK_PICKED' AND qty IS NOT NULL
    GROUP BY txn_date, hour_num, change_uid

    UNION ALL
	SELECT 'PARCEL_REPACK', txn_date, hour_num, change_uid,
		CAST(NULL AS decimal(18,2)), CAST(SUM(Eaches) AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
		CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
	FROM #repacking_base
	GROUP BY txn_date, hour_num, change_uid

    UNION ALL
    SELECT 'REPLENISHMENT', txn_date, hour_num, change_uid,
        SUM(CAST(quantity AS decimal(18,2))), CAST(COUNT(DISTINCT case_id) AS decimal(18,2)), CAST(COUNT(DISTINCT pick_task_id) AS decimal(18,2)),
        CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2))
    FROM #replen_base
    GROUP BY txn_date, hour_num, change_uid
)
SELECT
    Activity, txn_date AS [DATE],
    'HH' + RIGHT('0' + CAST(hour_num AS varchar(2)), 2) AS [HOUR],
    [USER], Eaches, Cartons, Pallets, Totes, Cases, Carton_CMC, Combined
FROM combined
ORDER BY [DATE], Activity, [USER], [HOUR];
"@
}

#endregion









#region dashboardCallerFunctions


#These are the "screens", they call function above to display.
#Share common boring pattern but helper functions.
#Means less copy and paste!!!!!

<#
.SYNOPSIS
    Displays the Fill Percentage report.
.DESCRIPTION
    Runs the fill percentage query and shows results as a simple Format-Table.
    No refresh loop — single-shot display.
.EXAMPLE
    Invoke-FillPercentage
#>
function Invoke-FillPercentage {
    [CmdletBinding()]
    param()

    Clear-Host
    Write-Host "Raw Fill Data Query" -ForegroundColor Green

    try {
        $data = Invoke-SqlQueryDirect -Query (Get-FillPercentageQuery)
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


<#
.SYNOPSIS
    Displays an hourly performance dashboard with a manual refresh prompt.
.DESCRIPTION
    Generic dashboard runner used by all the KPI screens. Queries data, converts it,
    renders the pivot table, then waits for the user to choose refresh or exit.
.PARAMETER Title
    The banner title shown at the top (e.g. "DECANT PERFORMANCE").
.PARAMETER QueryResult
    A hashtable with keys 'Query' (SQL string) and 'UOM' (enum value),
    as returned by Get-DecantPerformanceQuery or Get-GTPPickingQuery.
.PARAMETER Thresholds
    Colour threshold hashtable with 'High' and 'Medium' keys.
.EXAMPLE
    $qr = Get-DecantPerformanceQuery
    $cfg = Get-DashboardConfig
    Show-HourlyDashboard -Title "DECANT PERFORMANCE" `
                         -QueryResult $qr `
                         -Thresholds $cfg.DecantThresholds
#>
function Show-HourlyDashboard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [hashtable]$QueryResult,

        [hashtable]$Thresholds     = @{}
    )

    $continueRunning = $true
    while ($continueRunning) {
        Clear-Host
        Write-Host ("=" * 66) -ForegroundColor Cyan
        Write-Host "                  $Title - HOURLY BREAKDOWN" -ForegroundColor Cyan
        Write-Host ("=" * 66) -ForegroundColor Cyan
        Write-Host ""
        Write-Host "UOM DISPLAYED BELOW IS: $($QueryResult.UOM)" -ForegroundColor Red

        try {
            $data = Invoke-SqlQueryDirect -Query $QueryResult.Query
            if ($data.Rows.Count -eq 0) {
                Write-Host "No data found for the selected date." -ForegroundColor Yellow
                Pause
                return
            }

            $results = ConvertTo-PivotData -DataTable $data -PropertyMap @{
                User   = "User"
                Hour   = "Hour"
                Eaches = "Eaches"
            }

            Show-PivotTable -Data $results `
                -RowProperty "User" `
                -ColumnProperty "Hour" `
                -ValueProperty "Eaches" `
                -ColumnEnumOverride ([HOURS_DAY]) `
                -ColorThresholds $Thresholds

            Write-Host ""
            Write-Host "Query Ran at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Green
        }
        catch {
            Write-Host "Error retrieving data: $_" -ForegroundColor Red
            Pause
        }

        Write-Host ""
        $continueRunning = Wait-UserAction
    }
}


<#
.SYNOPSIS
    Launches the Decant Performance dashboard.
.EXAMPLE
    Invoke-DecantPerformance
#>
function Invoke-DecantPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-DecantPerformanceQuery
    Show-HourlyDashboard -Title "DECANT PERFORMANCE" -QueryResult $qr -Thresholds $cfg.DecantThresholds
}


<#
.SYNOPSIS
    Launches the GTP Picking Performance dashboard.
.EXAMPLE
    Invoke-GTPPickingPerformance
#>
function Invoke-GTPPickingPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-GTPPickingQuery
    Show-HourlyDashboard -Title "GTP PICKING PERFORMANCE" -QueryResult $qr -Thresholds $cfg.GTPThresholds
}

<#
.SYNOPSIS
    Launches the Replenishment Performance dashboard.
.EXAMPLE
    Invoke-ReplenishmentPerformance
#>
function Invoke-ReplenishmentPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-ReplenishmentPerformanceQuery
    Show-HourlyDashboard -Title "REPLENISHMENT PERFORMANCE" -QueryResult $qr -Thresholds $cfg.ReplenishmentThresholds
}


<#
.SYNOPSIS
    Displays Consumable Usage for the last 7 days.
.EXAMPLE
    Invoke-ConsumableUsage
#>
function Invoke-ConsumableUsage {
    [CmdletBinding()]
    param()

    Clear-Host
    Write-Host "Consumable Usage Query" -ForegroundColor Green
    Write-Host "Showing count of cartons used in the last 7 days by carton type." -ForegroundColor Green
    $fromDate = (Get-Date).AddDays(-7)
    $toDate   = Get-Date
    Write-Host "Date Range: $($fromDate.ToShortDateString()) - $($toDate.ToShortDateString())" -ForegroundColor Green
    Write-Host "WARNING: THIS QUERY IS NOT SEARCHING MI_DU THEREFORE MIGHT BE MISSING ALOT" -ForegroundColor Red

    try {
        $data    = Invoke-SqlQueryDirect -Query (Get-ConsumableUsageQuery)
        $results = ConvertTo-PivotData -DataTable $data -PropertyMap @{
            Date      = "Date"
            CartonType = "CartonType"
            Eaches    = "CartonCount"
        }

        Show-PivotTable -Data $results `
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


<#
.SYNOPSIS
    Displays Eaches Per Day by pick state.
.EXAMPLE
    Invoke-EachesPerDay
#>
function Invoke-EachesPerDay {
    [CmdletBinding()]
    param()

    Clear-Host
    Write-Host "Eaches Per Day Query" -ForegroundColor Green

    try {
        $data    = Invoke-SqlQueryDirect -Query (Get-EachesPerDayQuery)
        $results = ConvertTo-PivotData -DataTable $data -PropertyMap @{
            DateDay = "DateDay"
            State   = "State"
            Eaches  = "Eaches"
        }

        Show-PivotTable -Data $results `
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
.SYNOPSIS
    Displays through put of GTP locations by hour.
.EXAMPLE
    Invoke-GTPUtilisation
#>

function Invoke-GTPUtilisation {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-GTPUtilisationQuery
    Show-HourlyDashboard -Title "GTP UTILISATION" -QueryResult $qr -Thresholds $cfg.GTPUTILISATIONThresholds

}

function  Invoke-PutawayPerformanceQuery {
    [CmdletBinding()]
    param ()
    
    $cfg = Get-DashboardConfig
    $qr  = Get-PutawayPerformanceQuery
    Show-HourlyDashboard -Title "PUTAWAY PERFORMANCE" -QueryResult $qr -Thresholds $cfg.PutawayThresholds

}

Function Invoke-ReceivingPerformanceQuery {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr = Get-ReceivingPerformanceQuery
    Show-HourlyDashboard -Title "RECEIVING PERFORMANCE" -QueryResult $qr -Thresholds $cfg.ReceivingThresholds
}

Function Invoke-InboundTrammingPerformanceQuery {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr = Get-InboundTrammingPerformanceQuery
    Show-HourlyDashboard -Title "INBOUND-TRAMMING PERFORMANCE" -QueryResult $qr -Thresholds $cfg.InboundTrammingThresholds

}


<#
.SYNOPSIS
    Launches the Returns Processing dashboard.
.EXAMPLE
    Invoke-ReturnsPerformance
#>
function Invoke-ReturnsPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-ReturnsPerformanceQuery
    Show-HourlyDashboard -Title "RETURNS PROCESSING" -QueryResult $qr -Thresholds $cfg.ReturnsThresholds
}


<#
.SYNOPSIS
    Launches the CUBI (Dimensioning) dashboard.
.EXAMPLE
    Invoke-CubiPerformance
#>
function Invoke-CubiPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-CubiPerformanceQuery
    Show-HourlyDashboard -Title "CUBI (DIMENSIONING)" -QueryResult $qr -Thresholds $cfg.CubiThresholds
}


<#
.SYNOPSIS
    Launches the VAS (Value Added Services) Performance dashboard.
.EXAMPLE
    Invoke-VASPerformance
#>
function Invoke-VASPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-VASPerformanceQuery
    Show-HourlyDashboard -Title "VAS PERFORMANCE" -QueryResult $qr -Thresholds $cfg.VASThresholds
}


<#
.SYNOPSIS
    Launches the QI Check Performance dashboard.
.EXAMPLE
    Invoke-QIPerformance
#>
function Invoke-QIPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-QIPerformanceQuery
    Show-HourlyDashboard -Title "QI CHECK PERFORMANCE" -QueryResult $qr -Thresholds $cfg.QIThresholds
}


<#
.SYNOPSIS
    Launches the Parcel Repack Performance dashboard.
.EXAMPLE
    Invoke-ParcelRepackPerformance
#>
function Invoke-ParcelRepackPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-ParcelRepackPerformanceQuery
    Show-HourlyDashboard -Title "PARCEL REPACK PERFORMANCE" -QueryResult $qr -Thresholds $cfg.ParcelRepackThresholds
}


<#
.SYNOPSIS
    Launches the Despatch Performance dashboard.
.EXAMPLE
    Invoke-DespatchPerformance
#>
function Invoke-DespatchPerformance {
    [CmdletBinding()]
    param()

    $cfg = Get-DashboardConfig
    $qr  = Get-DespatchPerformanceQuery
    Show-HourlyDashboard -Title "DESPATCH PERFORMANCE" -QueryResult $qr -Thresholds $cfg.DespatchThresholds
}


<#
.SYNOPSIS
    Displays a single-day snapshot across every KPI activity at once.
.DESCRIPTION
    Runs the combined multi-activity query (builds all temp tables for one day,
    UNIONs them together) then shows a per-activity totals summary: active users
    and summed Eaches/Cartons/Pallets/Totes/Cases for the whole day. Optionally
    also dumps the raw hour-by-hour detail if asked.
.NOTES
    This is heavier than the single-activity screens since it builds temp tables
    for every activity in one go - fine for a single day, don't be tempted to
    widen the date range without re-adding indexes/considering server load.
.EXAMPLE
    Invoke-KPIOverview
#>
function Invoke-KPIOverview {
    [CmdletBinding()]
    param()

    Clear-Host
    Write-Host "All Activities Overview - single day snapshot" -ForegroundColor Green
    Write-Host "Builds every activity's data for one day then summarises. Might take a moment." -ForegroundColor Yellow

    $targetDate = Read-DateSelection
    $sql = Get-KPIOverviewQuery -TargetDate $targetDate

    try {
        $data = Invoke-SqlQueryDirect -Query $sql
        if ($data.Rows.Count -eq 0) {
            Write-Host "No activity found for $targetDate." -ForegroundColor Yellow
            Pause
            return
        }

        # NR-NOTE: SQL NULLs come back from a DataTable as [DBNull], not $null - Measure-Object
        # throws on that ("is not numeric"), so convert first. See ConvertTo-CleanObjects.
        $rows = ConvertTo-CleanObjects -DataTable $data

        Write-Host ""
        Write-Host "Summary for $targetDate - totals per activity across the whole day:" -ForegroundColor Cyan

        $summary = $rows | Group-Object Activity | ForEach-Object {
            $group = $_.Group
            [PSCustomObject]@{
                Activity    = $_.Name
                ActiveUsers = ($group | Select-Object -ExpandProperty USER -Unique).Count
                Eaches      = ($group | Where-Object { $null -ne $_.Eaches }  | Measure-Object -Property Eaches -Sum).Sum
                Cartons     = ($group | Where-Object { $null -ne $_.Cartons } | Measure-Object -Property Cartons -Sum).Sum
                Pallets     = ($group | Where-Object { $null -ne $_.Pallets } | Measure-Object -Property Pallets -Sum).Sum
                Totes       = ($group | Where-Object { $null -ne $_.Totes }   | Measure-Object -Property Totes -Sum).Sum
                Cases       = ($group | Where-Object { $null -ne $_.Cases }   | Measure-Object -Property Cases -Sum).Sum
            }
        } | Sort-Object Activity

        $summary | Format-Table -AutoSize

        $showDetail = Read-Host "Show raw hour-by-hour detail too? (y/N)"
        if ($showDetail -match '^[Yy]') {
            $rows | Sort-Object Activity, HOUR, USER |
                Format-Table Activity, HOUR, USER, Eaches, Cartons, Pallets, Totes, Cases -AutoSize
        }

        Pause
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Pause
    }
}


#Place holders Screens not yet implemented, to avoid user confusion and errors if they click something we haven't built yet.
function Invoke-InventoryQueryMenu { Write-Host "Inventory Query Menu not yet implemented" -ForegroundColor Yellow; Pause }
function Invoke-Troubleshoot       { Write-Host "Troubleshoot not yet implemented" -ForegroundColor Yellow; Pause }
function Invoke-Extras             { Write-Host "Extras not yet implemented" -ForegroundColor Yellow; Pause }

function Show-HelpMenu {
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

#endregion










#region menuNavigationScreens

<#
.SYNOPSIS
    Displays the KPI sub-menu.
.DESCRIPTION
    Loops until the user presses B to go back. Routes to the various KPI dashboards.
.EXAMPLE
    Show-KPIsMenu
#>
function Show-KPIsMenu {
    [CmdletBinding()]
    param()

    while ($true) {
        Clear-Host
        Write-Host ""
        Write-Host "       ______________________________________________________________"
        Write-Host "        GREEN = WORKING " -ForegroundColor Green -NoNewline
        Write-Host ":: YELLOW = SOON! " -ForegroundColor Yellow -NoNewline
        Write-Host ":: RED = NOT SET UP " -ForegroundColor Red
        Write-Host "                 KPIs Menu:"
        Write-Host ""
        Write-Host "             [1-4] Fulfilment Performance KPIs" -ForegroundColor Green
        Write-Host ""
        Write-Host "             [1]  Decant Performance"   -ForegroundColor Green
        Write-Host "             [2]  GTP Picking Performance" -ForegroundColor Green
        Write-Host "             [3]  Replenishment Performance" -ForegroundColor Green
        Write-Host "             [4]  VAS Performance" -ForegroundColor Green
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [5-8] Inbound Performance KPIs" -ForegroundColor Green
        Write-Host ""
        Write-Host "             [5]  Receiving Performance" -ForegroundColor Green
        Write-Host "             [6]  PutAway Performance" -ForegroundColor Green
        Write-Host "             [7]  Inbound Tramming Performance" -ForegroundColor Green
        Write-Host "             [8]  QI Check Performance" -ForegroundColor Green
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [9]  Inventory Turnover" -ForegroundColor Red
        Write-Host "             [10] Dock-to-Stock Time" -ForegroundColor Red
        Write-Host "             [11] Labor Productivity" -ForegroundColor Red
        Write-Host "             [12] Returns Processing" -ForegroundColor Green
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [13-15] Additional Performance KPIs" -ForegroundColor Green
        Write-Host ""
        Write-Host "             [13] CUBI (Dimensioning) Performance" -ForegroundColor Green
        Write-Host "             [14] Parcel Repack Performance" -ForegroundColor Green
        Write-Host "             [15] Despatch Performance" -ForegroundColor Green
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [O]  All Activities Overview (single day snapshot)" -ForegroundColor Green
        Write-Host "             [B] Back to Main Menu"
        Write-Host "       ______________________________________________________________"
        Write-Host ""

        $choice = Read-Host "Choose a KPI option"

        switch ($choice.ToUpper()) {
            "1"  { Invoke-DecantPerformance }
            "2"  { Invoke-GTPPickingPerformance }
            "3"  { Invoke-ReplenishmentPerformance }
            "4"  { Invoke-VASPerformance }
            #Space for me
            "5"  { Invoke-ReceivingPerformanceQuery}
            "6"  { Invoke-PutawayPerformanceQuery}
            "7"  { Invoke-InboundTrammingPerformanceQuery }
            "8"  { Invoke-QIPerformance }
            "9"  { Write-Host "not yet implemented" -ForegroundColor Yellow; Pause }
            "10" { Write-Host "not yet implemented" -ForegroundColor Yellow; Pause }
            "11" { Write-Host "not yet implemented" -ForegroundColor Yellow; Pause }
            "12" { Invoke-ReturnsPerformance }
            "13" { Invoke-CubiPerformance }
            "14" { Invoke-ParcelRepackPerformance }
            "15" { Invoke-DespatchPerformance }
            "O"  { Invoke-KPIOverview }
            "B"  { return }
            default {
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep 1
            }
        }
    }
}


<#
.SYNOPSIS
    Displays the main menu and routes user selections.
.DESCRIPTION
    Entry point loop. Runs until the user selects 0 (Exit).
.EXAMPLE
    Show-MainMenu
#>
function Show-MainMenu {
    [CmdletBinding()]
    param()

    while ($true) {
        Clear-Host

        Write-Host ""
        Write-Host "       ______________________________________________________________"
        Write-Host "        GREEN = WORKING " -ForegroundColor Green -NoNewline
        Write-Host ":: YELLOW = SOON! " -ForegroundColor Yellow -NoNewline
        Write-Host ":: RED = NOT SET UP " -ForegroundColor Red
        Write-Host "                 Logs Methods:"
        Write-Host ""
        Write-Host "             [1] Fill Percentage                          - AUKC01" -ForegroundColor Green
        Write-Host "             [2] Consumable Usage (WIP)                   - AUKC01" -ForegroundColor Green
        Write-Host "             [3] Operation KPIs Menu                      - AUKC01" -ForegroundColor Green
        Write-Host "             [4] Picks By Day                             - AUKC01" -ForegroundColor Green
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [5] GTP Utilisation" -ForegroundColor Green
        Write-Host "             [6] Random Selection 6" -ForegroundColor Red
        Write-Host "             [7] Random Selection 7" -ForegroundColor Red
        Write-Host "             __________________________________________________"
        Write-Host ""
        Write-Host "             [8] Troubleshoot" -ForegroundColor Red
        Write-Host "             [E] Extras" -ForegroundColor Red
        Write-Host "             [H] Help" -ForegroundColor Green
        Write-Host "             [0] Exit" -ForegroundColor Green
        Write-Host "       ______________________________________________________________"
        Write-Host ""

        $choice = Read-Host "Choose a menu option"

        switch ($choice.ToUpper()) {
            "1" { Invoke-FillPercentage }
            "2" { Invoke-ConsumableUsage }
            "3" { Show-KPIsMenu }
            "4" { Invoke-EachesPerDay }
            "5" { Invoke-GTPUtilisation }
            "6" { Write-Host "Its red... why did you select!"; Pause }
            "7" { Write-Host "Its red... why did you select!"; Pause }
            "8" { Invoke-Troubleshoot }
            "E" { Invoke-Extras }
            "H" { Show-HelpMenu }
            "0" { exit }
            default {
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep 1
            }
        }
    }
}

#endregion

#region updatenotice

function Show-UpdateNotice {
    Clear-Host

    Write-Host "=============================================================" -ForegroundColor Red
    Write-Host "                    WCS CHECKER UPDATE" -ForegroundColor Red
    Write-Host "=============================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Updated: " -NoNewline
    Write-Host "04-08-2026" -ForegroundColor Yellow
    Write-Host "Author : " -NoNewline
    Write-Host "Nathaniel Ritchie" -ForegroundColor Green
    Write-Host ""
    Write-Host "This release introduces updated SQL queries and reporting logic."
    Write-Host "For the most accurate results, please ensure you are familiar"
    Write-Host "with the new calculations and what they now remove."
    Write-Host ""
    Write-Host "This is expected to be the final standalone release before the"
    Write-Host "tool is merged into the WCS Labour Tracker workbook."
    Write-Host ""
    Write-Host "Thank you for your continued feedback and support."
    Write-Host "-Nathaniel Ritchie"

    for ($i = 10; $i -gt 0; $i--) {
        Write-Host "`rLaunching in $i second(s)... " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }

    Clear-Host
}



#endregion



#region startUp
# Script entry point

<#
.SYNOPSIS
    Launches the interactive console dashboard (the original CLI entry point).
.DESCRIPTION
    Pulled out into its own function - see the guard below - so this whole file
    can also be dot-sourced (e.g. by WCS-Checker-GUI.ps1) purely to reuse its
    config/query-building/SQL-execution functions, without it immediately
    clearing the host and blocking on Show-MainMenu's Read-Host loop.
.EXAMPLE
    Start-WCSCheckerConsole
#>
function Start-WCSCheckerConsole {
    [CmdletBinding()]
    param()

    Clear-Host
    $Host.UI.RawUI.WindowTitle = "NR WCS Performance Dashboard - AUKC01"

    # Force TLS 1.2+
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

    # Ensure prerequisites
    Assert-SqlModule

    # Set console size
    try {
        $size = $Host.UI.RawUI.WindowSize
        $size.Width  = 150
        $size.Height = 50
        $Host.UI.RawUI.WindowSize = $size
    }
    catch {
        # Non-fatal — some terminals don't support resize
    }

    # Go
    try {
        Show-UpdateNotice
        Show-MainMenu
    }
    catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Pause
    }

}

# NR-NOTE: only auto-launch the console menu when this script is run directly (.\WCS-Checker.ps1).
# When it's dot-sourced instead (". .\WCS-Checker.ps1", which the GUI companion script does to
# reuse all the functions above), $MyInvocation.InvocationName is "." - skip auto-launch in that case.
if ($MyInvocation.InvocationName -ne '.') {
    Start-WCSCheckerConsole
}

#endregion