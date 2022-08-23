###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 16-04-2020
# COMMENT : This script is gathering log
#           from server for DRC 
#           Windows image maintenance (wim) jobs
#
###########################################################

# Connection to database
$sqlServerInfo = "<server name>"
$sqlInstanceInfo = Get-SqlInstance -Path "SQLSERVER:\SQL\<server name>\DEFAULT"
$sqlDatabaseInfo = "WindowsImageMaintenance"

# Defining variables
$messagesArray = New-Object -TypeName "System.Collections.ArrayList"
$wimReportData = "$env:ProgramData\udreports\wimreport\drc-wim-data.csv"
Get-Item -Path $wimReportData | Remove-Item -Force -Confirm:$false
[regex]$regexLogDate = "\d\d\d\d-\d\d-\d\d"
[regex]$regexLogTime = "\d\d:\d\d:\d\d"

if (!(Get-ChildItem -Path "$env:ProgramData\vmware-templates-maintenance-drc\Logs" -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1)) {
    exit
} else {
    $latestLogFileContent = Get-ChildItem -Path "$env:ProgramData\vmware-templates-maintenance-drc\Logs" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1 | Get-Content
    
    foreach ($line in $latestLogFileContent) {
        $logDate = $regexLogDate.Match($line).Value
    	$logTime = $regexLogTime.Match($line).Value
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
		INSERT INTO wim_drc (
			log_date,
			log_time,
			log_level,
			log_message
		) VALUES (
			'$logDate',
			'$logTime',
			'$logLevel',
			'$logMessage'
		);
        "
        
        Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $messageQuery
        
        $messageObject = [PSCustomObject]@{
            Date = $logDate
            Time = $logTime
            Level = $logLevel
            Message = $logMessage
        }
        $messagesArray.Add($messageObject)
    }
}

$messagesArray | Export-Csv -Path $wimReportData -NoTypeInformation