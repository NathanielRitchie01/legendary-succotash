Clear-Host
$Host.UI.RawUI.WindowTitle = "Nathaniel Ritchie RandomScript"

# Force TLS 1.2 or higher (TLS 1.0 and 1.1 are deprecated)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

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
        Write-Host "             [4] Again Another                            - AUKC01"
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
            "4" { Write-Host "Option 4 not implemented"; Pause }
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
        [string]$targetDate = "2026-02-09" #(Get-Date -Format "yyyy-MM-dd")  # Default to today's date
    )
   
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
        [string]$query = (DecantPerformanceQuery)
    )

    #Unsure if want another function to obtain DecantKPISettings but all welllll.
    $Decant_high = 100
    $Decant_medium = 50

    #Unsure if want another function to obtain OverAll RefreshSettings but all welllll.
    $refreshInterval = 900  # 15 minutes in seconds
    $continueRunning = $true

   

    while($continueRunning) {
        Clear-Host
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host "                  DECANT PERFORMANCE - HOURLY BREAKDOWN" -ForegroundColor Cyan
        Write-Host "==================================================================" -ForegroundColor Cyan
        Write-Host ""
       
        #Possible to get user date by calling UserSelectDate but for now just defaulting to today for ease of use
        #link is not working lol - reason is due to the fact I am doing a string passing and need to edit said string.
       
        #$targetDate = Get-Date -Format "yyyy-MM-dd"
        $targetDate = "2026-02-09"
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

            # Create pivot table (Users as rows, Hours as columns)
            $users = $results | Select-Object -ExpandProperty User -Unique | Sort-Object
            $hours = 0..23  # All possible hours in a day

            # Build the pivot table
            Write-Host ("User".PadRight(15)) -NoNewline -ForegroundColor Cyan
            foreach ($hour in $hours) {
                Write-Host ("{0:D2}h" -f $hour).PadLeft(6) -NoNewline -ForegroundColor Cyan
            }
            Write-Host ("  Total".PadLeft(8)) -ForegroundColor Cyan
            Write-Host ("-" * 165) -ForegroundColor Gray

            foreach ($user in $users) {
                $userTotal = 0
                Write-Host ($user.PadRight(15)) -NoNewline

                foreach ($hour in $hours) {
                    $eaches = ($results | Where-Object { $_.User -eq $user -and $_.Hour -eq $hour }).Eaches
                   
                    if ($null -eq $eaches -or $eaches -eq 0) {
                        Write-Host "     -" -NoNewline -ForegroundColor DarkGray
                    } else {
                        $userTotal += $eaches
                        # Color coding for performance
                        if ($eaches -ge $Decant_high) {
                            Write-Host ("{0,6}" -f $eaches) -NoNewline -ForegroundColor Green
                        } elseif ($eaches -ge $Decant_medium) {
                            Write-Host ("{0,6}" -f $eaches) -NoNewline -ForegroundColor Yellow
                        } else {
                            Write-Host ("{0,6}" -f $eaches) -NoNewline -ForegroundColor Red
                        }
                    }
                }
                Write-Host ("{0,8}" -f $userTotal) -ForegroundColor Cyan
            }


           
            # Hourly totals row
            Write-Host ("-" * 165) -ForegroundColor Gray
            Write-Host ("HOURLY TOTAL".PadRight(15)) -NoNewline -ForegroundColor Cyan
            $grandTotal = 0
            foreach ($hour in $hours) {
                $hourTotal = ($results | Where-Object { $_.Hour -eq $hour } | Measure-Object -Property Eaches -Sum).Sum
                if ($hourTotal -gt 0) {
                    Write-Host ("{0,6}" -f $hourTotal) -NoNewline -ForegroundColor Cyan
                    $grandTotal += $hourTotal
                } else {
                    Write-Host "     -" -NoNewline -ForegroundColor DarkGray
                }
            }
            Write-Host ("{0,8}" -f $grandTotal) -ForegroundColor Green

            Write-Host ""
            Write-Host "==================================================================" -ForegroundColor Cyan
            Write-Host "Color Legend: " -NoNewline
            Write-Host "Green >= ?SGJN eaches/hr  " -NoNewline -ForegroundColor Green
            Write-Host "Yellow >= DSLFKJS eaches/hr  " -NoNewline -ForegroundColor Yellow
            Write-Host "White < SDL eaches/hr" -ForegroundColor White
            Write-Host "==================================================================" -ForegroundColor Cyan
            Write-Host ""


            # Last updated timestamp
            $lastUpdated = Get-Date -Format "HH:mm:ss"
            Write-Host ""
            Write-Host "Last Updated: $lastUpdated" -ForegroundColor Green
           
            Pause
        }
        catch {
            Write-Host "Error retrieving decant performance data: $_" -ForegroundColor Red
            Pause
        }

        # This all below for refreshing could quite literally be considered an overall performance function somewhere else lol..... This code is getting messy but it works for now and I want to get the other KPIs in there before I refactor the whole thing.

        # Countdown timer with key detection
        Write-Host ""
        Write-Host "Next refresh in: " -NoNewline
       
        $secondsRemaining = $refreshInterval
        $keyPressed = $false

        while ($secondsRemaining -gt 0 -and -not $keyPressed) {
            # Check if a key has been pressed
            if ($host.UI.RawUI.KeyAvailable) {
                $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
               
                # Check for Enter or Escape or Q key
                if ($key.VirtualKeyCode -eq 13 -or $key.VirtualKeyCode -eq 27 -or $key.VirtualKeyCode -eq 81) {  # 13=Enter, 27=Esc, 81=Q
                    $keyPressed = $true
                    $continueRunning = $false
                    Write-Host "`n`nReturning to KPI menu..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                    break
                }
            }

            # Display countdown
            $minutes = [math]::Floor($secondsRemaining / 60)
            $seconds = $secondsRemaining % 60
           
            # Move cursor to beginning of line and overwrite
            $cursorPosition = $host.UI.RawUI.CursorPosition
            $cursorPosition.X = 17  # Position after "Next refresh in: "
            $host.UI.RawUI.CursorPosition = $cursorPosition
           
            Write-Host ("{0:D2}:{1:D2}" -f $minutes, $seconds) -NoNewline -ForegroundColor Cyan
           
            Start-Sleep -Seconds 1
            $secondsRemaining--
        }

        # If we exited due to timeout (not key press), continue the loop
        if (-not $keyPressed) {
            Write-Host "`n`nRefreshing data..." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
   

        #this god damn bracket is for the while loop.
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
