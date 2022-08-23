###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 20-04-2020
# COMMENT : This script is gathering logs
#           from Kaspersky Security Center
#           about Microsoft SQL database
#
###########################################################

$checkCounter = 0
$siteStopped = $true
$timeStamp = Get-Date -Format dd-MM-yyyy-HH-mm-ss
$kscServerList = '<server name>', '<server name>'


Get-IISSite -Name 'sql-dashboard' | Stop-IISSite -Confirm:$false
do {
    if ((Get-IISSite -Name 'sql-dashboard').State -like 'Started') {
        $checkCounter++
    }

    if ($checkCounter -eq 12) {
        $siteStopped = $false
        break
    }
    Start-Sleep -Seconds 5
} until ((Get-IISSite -Name 'sql-dashboard').State -like 'Stopped')

if ($siteStopped -eq $true) {
    foreach ($kscServer in $kscServerList) {
        $kscServerCC = $kscServer.Split('-')[0]
        
        if (!(Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\sql-report\logs' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1} -ErrorAction SilentlyContinue)) {
            #TODO Make condition that data are unavailable
        } else {
            $latestSqlLog = Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\sql-report\logs' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1}
            $sqlLogName = $latestSqlLog.Name.Split('-')[0] + $latestSqlLog.Name.Split('-')[1] + "-data.csv"
            if (Test-Path -Path "$env:ProgramData\udreports\$kscServerCC\sql-data\$sqlLogName") {
                Get-Item -Path "$env:ProgramData\udreports\$kscServerCC\sql-data\$sqlLogName" | Remove-Item -Force -Confirm:$false
            } else {
                New-Item -Path "$env:ProgramData\udreports\$kscServerCC\sql-data" -Name $sqlLogName -ItemType File
            }
            Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-Content -Path "C:\ProgramData\sql-report\logs\$using:latestSqlLog"} | Out-File -FilePath "$env:ProgramData\udreports\$kscServerCC\sql-data\$sqlLogName" -Append
        }
        
        $kscServerCC = $null
        $latestSqlLog = $null
    }

    Get-IISSite -Name 'sql-dashboard' | Start-IISSite
} elseif ($siteStopped -eq $false) {
    "$timeStamp - IIS site could not be stopped" | Out-File -FilePath "$env:ProgramData\udreports\sqlreport\sqlreport_error_log.txt" -Append
} else {
    "$timeStamp - IIS site could not be stopped" | Out-File -FilePath "$env:ProgramData\udreports\sqlreport\sqlreport_error_log.txt" -Append
}