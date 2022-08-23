###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 15-04-2020
# COMMENT : This script is gathering logs
#           from Kaspersky Security Center
#           for PDC Windows image maintenance (wim) jobs
#
###########################################################

# Connection to database
$sqlServerInfo = "<server name>"
$sqlInstanceInfo = Get-SqlInstance -Path "SQLSERVER:\SQL\<server name>\DEFAULT"
$sqlDatabaseInfo = "WindowsImageMaintenance"

# Defining variables
$kscServerList = '<server name>', '<server name>'
$filteredMessagesArray = New-Object -TypeName "System.Collections.ArrayList"
$wimReportData = "$env:ProgramData\udreports\wimreport\pdc-wim-data.csv"
Get-Item -Path $wimReportData | Remove-Item -Force -Confirm:$false
[regex]$regexLogDate = "\d\d\d\d-\d\d-\d\d"
[regex]$regexLogTime = "\d\d:\d\d:\d\d"

foreach ($kscServer in $kscServerList) {
    $kscServerCC = $kscServer.Split('-')[0]
    
    # Windows Server 2016 Standard image log
    if (!(Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\vmware-templates-maintenance\win2k16' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1} -ErrorAction SilentlyContinue)) {
        $messageObject = [PSCustomObject]@{
            Country = $kscServerCC.ToUpper()
            Date = Get-Date -Format yyyy-mm-dd
            Time = Get-Date -Format HH:mm:ss
            OS = $kscServer
            Level = "ERROR"
            Message = "Could not connect or get data"
        }
        $filteredMessagesArray.Add($messageObject)
    } else {
        $latestLogFile2016 = Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\vmware-templates-maintenance\win2k16' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1}
        if (Test-Path -Path "$env:ProgramData\udreports\$kscServerCC\win2k16\$latestLogFile2016") {
            Get-Item -Path "$env:ProgramData\udreports\$kscServerCC\win2k16\$latestLogFile2016" | Remove-Item -Force -Confirm:$false
        } else {
            New-Item -Path "$env:ProgramData\udreports\$kscServerCC\win2k16" -Name $latestLogFile2016.Name -ItemType File
        }
        Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-Content -Path "C:\ProgramData\vmware-templates-maintenance\win2k16\$using:latestLogFile2016"} | Out-File -FilePath "$env:ProgramData\udreports\$kscServerCC\win2k16\$latestLogFile2016"
        
        $latestLogFileContent2016 = Get-Content -Path "$env:ProgramData\udreports\$kscServerCC\win2k16\$latestLogFile2016"
        # Writing messages to database
        foreach ($line in $latestLogFileContent2016) {
            $logCountry = $kscServerCC
            $logDate = $regexLogDate.Match($line).Value
            $logTime = $regexLogTime.Match($line).Value
            $logOs = "Win2k16"
            $logLevel = $line.Split('-')[3].Trim()
            [int]$dashCounter = 0
            [int]$dashPosition = 0
            for ($i = 0; $i -lt $line.Length; $i++) {
                if ($line[$i] -eq '-') {
                    $dashCounter++
                    if ($dashCounter -eq 4) {
                        $dashPosition = $i
                        break
                    }
                }
            }
            $logMessage = $line.Substring($dashPosition + 2, $line.Length - ($dashPosition + 2))

            $messageQuery = "
		    INSERT INTO wim_pdc (
                log_country,
                log_date,
                log_time,
                log_os,
			    log_level,
			    log_message
		    ) VALUES (
                '$logCountry',
                '$logDate',
                '$logTime',
                '$logOs',
			    '$logLevel',
			    '$logMessage'
		    );
            "

            Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $messageQuery

            if (($line -match "WARNING") -or ($line -match "ERROR")) {
                $messageObject = [PSCustomObject]@{
                    Country = $logCountry.ToUpper()
                    Date = $logDate
                    Time = $logTime
                    OS = $logOs
                    Level = $logLevel
                    Message = $logMessage
                }
                $filteredMessagesArray.Add($messageObject)
            }
        }
    }

    # Windows Server 2012R2 Standard image log
    if (!(Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\vmware-templates-maintenance\win2k12' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1} -ErrorAction SilentlyContinue)) {
        $messageObject = [PSCustomObject]@{
            Country = $kscServerCC.ToUpper()
            Date = Get-Date -Format yyyy-mm-dd
            Time = Get-Date -Format HH:mm:ss
            OS = $kscServer
            Level = "ERROR"
            Message = "Could not connect or get data"
        }
        $filteredMessagesArray.Add($messageObject)
    } else {
        $latestLogFile2012 = Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-ChildItem -Path 'C:\ProgramData\vmware-templates-maintenance\win2k12' | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1}
        if (Test-Path -Path "$env:ProgramData\udreports\$kscServerCC\win2k12\$latestLogFile2012") {
            Get-Item -Path "$env:ProgramData\udreports\$kscServerCC\win2k12\$latestLogFile2012" | Remove-Item -Force -Confirm:$false
        } else {
            New-Item -Path "$env:ProgramData\udreports\$kscServerCC\win2k12" -Name $latestLogFile2012.Name -ItemType File        
        }
        Invoke-Command -ComputerName $kscServer -ScriptBlock {Get-Content -Path "C:\ProgramData\vmware-templates-maintenance\win2k12\$using:latestLogFile2012"} | Out-File -FilePath "$env:ProgramData\udreports\$kscServerCC\win2k12\$latestLogFile2012"
        
        $latestLogFileContent2012 = Get-Content -Path "$env:ProgramData\udreports\$kscServerCC\win2k12\$latestLogFile2012"
        # Writing messages to database
        foreach ($line in $latestLogFileContent2012) {
            $logCountry = $kscServerCC
            $logDate = $regexLogDate.Match($line).Value
            $logTime = $regexLogTime.Match($line).Value
            $logOs = "Win2k12"
            $logLevel = $line.Split('-')[3].Trim()
            [int]$dashCounter = 0
            [int]$dashPosition = 0
            for ($i = 0; $i -lt $line.Length; $i++) {
                if ($line[$i] -eq '-') {
                    $dashCounter++
                    if ($dashCounter -eq 4) {
                        $dashPosition = $i
                        break
                    }
                }
            }
            $logMessage = $line.Substring($dashPosition + 2, $line.Length - ($dashPosition + 2))

            $messageQuery = "
		    INSERT INTO wim_pdc (
                log_country,
                log_date,
                log_time,
                log_os,
			    log_level,
			    log_message
		    ) VALUES (
                '$logCountry',
                '$logDate',
                '$logTime',
                '$logOs',
			    '$logLevel',
			    '$logMessage'
		    );
            "

            Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $messageQuery

            if (($line -match "WARNING") -or ($line -match "ERROR")) {
                $messageObject = [PSCustomObject]@{
                    Country = $logCountry.ToUpper()
                    Date = $logDate
                    Time = $logTime
                    OS = $logOs
                    Level = $logLevel
                    Message = $logMessage
                }
                $filteredMessagesArray.Add($messageObject)
            }
        }
        
    }
    
    $kscServerCC = $null
    $latestLogFile2016 = $null
    $latestLogFileContent2016 = $null
    $latestLogFile2012 = $null
    $latestLogFileContent2012 = $null
}

$filteredMessagesArray | Export-Csv -Path $wimReportData -NoTypeInformation