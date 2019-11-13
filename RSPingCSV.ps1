Templated from Chrissy LeMaire
# https://blog.netnerds.net/2016/12/runspaces-simplified/
# BLOCK 1: Create and open runspace pool, setup runspaces array with min and max threads


$synchash = [hashtable]::Synchronized(@{ })
#$pool = [RunspaceFactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS + 10)
$pool = [RunspaceFactory]::CreateRunspacePool(1, 20000)
$pool.ApartmentState = "MTA"
$pool.Open()
$runspaces = @()
$Global:Synchash = [hashtable]::Synchronized(@{ })

$Global:Synchash.time = (Get-Date -format "yy/MM/dd HH:mm")
$Global:Synchash.RSACount = 0   #Used to step count the threads.
$Global:Synchash.RSBCount = 0
$Global:Synchash.RSCCount = 0
$Global:Synchash.outReport = @{ }
$results = @()
$Global:Synchash.i = 0
$ExportPath = $env:APPDATA + "\" + "RSPRINGER" + "\" + "Export.csv"
$ErrorActionPreference = 'SilentlyContinue'

Write-Host "This script will export the .xlsx you provie into a csv. It will then run through the IPs. Once finished it will export to a CSV file. This file will be independed to the .xlsx file`r`n"


$TPath = Test-Path $ExportPath
if ($TPath -eq $false) {
    New-Item -Path ($env:APPDATA + "\" + "RSPinger") -ItemType Directory
    $Path = Read-Host "Full path of spreadsheet."
    $File = $Path.Replace('"', '')

    $excel = New-Object -ComObject excel.application
    $excel.DisplayAlerts = $false

    $WorkBook = $excel.Workbooks.Open($File)
    #    $Global:Synchash.WorkSheet = $Global:Synchash.excel.Worksheets.Item("Printer Tracker")
    #    $Global:Synchash.WorkSheet.SaveAs($ExportPath, 6)
    $WorkBook.SaveAs($ExportPath, 6)
    $WorkBook.Close()
    $excel.Quit()
}


Start-Sleep -Seconds 2 #Buffer between excel closing and grabbing csv.

(Get-Date).DateTime
$CSV = Import-CSV -Path $ExportPath
Remove-Item -Path $ExportPath -Force -ErrorAction SilentlyContinue

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$Stopwatch.Start()


# BLOCK 2: Create reusable scriptblock. This is the workhorse of the runspace. Think of it as a function.
$scriptblock = {
    Param ($IP, $Global:Synchash)
    $Global:Synchash.RSACount++
    try {

        $IPPING = ($IP.'IP Address').Trim()
        $i = 0
        while ($i -le 5) {
            $ThisPing = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($IPPING)
            $i++
            if ($ThisPing.Result.Status -eq 'Success') { break }
        }
    }
    catch {
        $ThisPing.Result.Status = "Failed to Ping"               # This is for SendPingAsync
    }

    $Global:Synchash.RSBCount++

    $T = [string]$Global:Synchash.time
    $Global:Synchash.outReport = ($IP | Select-Object *, @{Name = $T ; Expression = { $ThisPing.Result.Status } })

    $Global:Synchash.RSCCount++
    Return $Global:Synchash.outReport
}


# BLOCK 3: Create runspace and add to runspace pool

foreach ( $IP In $CSV) {

    if ($IP.'IP Address' -lt 1) {
        break
    }



    $runspace = [PowerShell]::Create()
    $Global:Synchash.i++
    $null = $runspace.AddScript($scriptblock)
    $null = $runspace.AddArgument($IP)
    $null = $runspace.AddArgument($Global:Synchash)
    $runspace.RunspacePool = $pool

    # BLOCK 4: Add runspace to runspaces collection and "start" it
    $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }

}


# BLOCK 5: Wait for runspaces to finish
while ($runspaces.Status.IsCompleted -notcontains $true) {
    Start-Sleep -Milliseconds 250
}

# BLOCK 6: Clean up
foreach ($runspace in $runspaces ) {
    # EndInvoke method retrieves the results of the asynchronous call
    $results += $runspace.Pipe.EndInvoke($runspace.Status)
    $runspace.Pipe.Dispose()
}

$pool.Close()
$pool.Dispose()

$Stopwatch.Elapsed
$Stopwatch.stop()

$results | Export-Csv -path $ExportPath -NoTypeInformation -Append -Force

.$ExportPath