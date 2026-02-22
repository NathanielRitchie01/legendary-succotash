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
#>


#region decEnums

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
    #Totes not supported. - 1-to-3 picking makes tote counting unreliable
    CARTONS = 3
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
    [hashtable] with keys: SQLServer, SQLDatabase, DefaultRefreshSeconds,
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

    return @{
        SQLServer             = "SQLDBAUP010"
        SQLDatabase           = "prodmis"
        DefaultRefreshSeconds = 7000
        DecantThresholds      = @{ High = 100; Medium = 50 }
        GTPThresholds         = @{ High = 200; Medium = 100 }
    }
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
        return $table
    }
    catch {
        Write-Host "Database connection error: $_" -ForegroundColor Red
        throw
    }
    finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
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
        return (Get-Date -Format "yyyy-MM-dd")
    }

    try {
        return [DateTime]::Parse($dateInput).ToString("yyyy-MM-dd")
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

    if ([string]::IsNullOrWhiteSpace($userInput)) { return -1 }

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

    if ($Thresholds.Count -eq 0) { return "White" }

    if ($Value -ge $Thresholds.High)   { return "Green"  }
    if ($Value -ge $Thresholds.Medium) { return "Yellow" }
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
    Runs a refresh countdown loop, returning $true to continue or $false to exit.
.DESCRIPTION
    Displays a live countdown timer. If the user presses Enter, Esc, or Q during the
    countdown, returns $false (stop). Otherwise returns $true (refresh the dashboard).
.PARAMETER Seconds
    Number of seconds to count down.
.OUTPUTS
    [bool] — $true = refresh again, $false = user wants to exit.
.EXAMPLE
    if (-not (Wait-RefreshCountdown -Seconds 120)) { return }
#>
function Wait-RefreshCountdown {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [int]$Seconds = 7000
    )

    $secondsRemaining = $Seconds
    while ($secondsRemaining -gt 0) {
        if ($host.UI.RawUI.KeyAvailable) {
            $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($key.VirtualKeyCode -eq 13 -or $key.VirtualKeyCode -eq 27 -or $key.VirtualKeyCode -eq 81) {
                Clear-Host
                Write-Host "Returning to menu..." -ForegroundColor Yellow
                return $false
            }
        }

        $minutes = [math]::Floor($secondsRemaining / 60)
        $secs    = $secondsRemaining % 60
        Write-Host ("`rNext refresh in: {0:D2}:{1:D2}  (Press Enter/Esc/Q to exit)" -f $minutes, $secs) -NoNewline -ForegroundColor Cyan
        Start-Sleep -Seconds 1
        $secondsRemaining--
    }

    Write-Host "`n`rRefreshing data..." -ForegroundColor Yellow
    return $true
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



#endregion









#region sqlQueriesBuilder

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

    Write-Host "Current default UOM selection: $Uom" -ForegroundColor Green
    $userSelection = Read-EnumSelection -EnumType ([DECANT_UOM])

    Write-Host "The current default selection: $Uom" -ForegroundColor Green
    $userSelection = userUOMSelection -enumType ([DECANT_UOM])
    if (Test-ValidEnumSelection -UserSelection $userSelection) {
        $Uom = [DECANT_UOM]$userSelection
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
     Write-Host "Current default UOM: $Uom (Totes is not available)" -ForegroundColor Green
    $userSelection = Read-EnumSelection -EnumType ([GTP_UOM])
    if (Test-ValidEnumSelection -UserSelection $userSelection) {
        $Uom = [GTP_UOM]$userSelection
    }

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

    $query = switch ([int]$Uom) {
        1 {  $GTP_QUERY_EACHES }
        2 {  $GTP_QUERY_TOTES }
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
        DATEPART(DAY,$filter) AS [DateDay],
        SUM(each_qty) AS [Eaches]
    FROM x_pick
    GROUP BY pick_state, DATEPART(DAY,$filter);
"@

}

<#
.SYNOPSIS
    Returns the GTP Utilisation SQL query (last 3 days).
.OUTPUTS
    [string] — T-SQL query.
#>
function Get-GTPUtilisationQuery {
    [CmdletBinding()]
    [OutputType([string])]
    param()

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
    Displays an auto-refreshing hourly performance dashboard.
.DESCRIPTION
    Generic dashboard runner used by both Decant and GTP Picking screens.
    Queries data, converts it, renders the pivot table, then enters a
    countdown loop.  Eliminates the duplicated refresh logic.
.PARAMETER Title
    The banner title shown at the top (e.g. "DECANT PERFORMANCE").
.PARAMETER QueryResult
    A hashtable with keys 'Query' (SQL string) and 'UOM' (enum value),
    as returned by Get-DecantPerformanceQuery or Get-GTPPickingQuery.
.PARAMETER Thresholds
    Colour threshold hashtable with 'High' and 'Medium' keys.
.PARAMETER RefreshSeconds
    Seconds between auto-refreshes. Default from config.
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

        [hashtable]$Thresholds     = @{},
        [int]$RefreshSeconds       = (Get-DashboardConfig).DefaultRefreshSeconds
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
            Write-Host "Press Enter to exit." -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error retrieving data: $_" -ForegroundColor Red
            Pause
        }

        Write-Host ""
        $continueRunning = Wait-RefreshCountdown -Seconds $RefreshSeconds
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
            "1"  { Invoke-DecantPerformance }
            "2"  { Invoke-GTPPickingPerformance }
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
            "1" { Invoke-FillPercentage }
            "2" { Invoke-ConsumableUsage }
            "3" { Show-KPIsMenu }
            "4" { Invoke-EachesPerDay }
            "5" { Write-Host "Option 5 not implemented"; Pause }
            "6" { Write-Host "Option 6 not implemented"; Pause }
            "7" { Write-Host "Option 7 not implemented"; Pause }
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



#region startUp
# Script entry point

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
Show-MainMenu

#endregion