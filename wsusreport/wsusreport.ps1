###########################################################
# AUTHOR  : Filip Vagner
# EMAIL   : filip.vagner@hotmail.com
# DATE    : 07-02-2020
# COMMENT : This script loads data about
#           Windows image status on WSUS
#           used for deplaoyment in each country
#
###########################################################

# Defining variables
$wsusServerList = '<server name>', '<server name>'
$wsusClientsListObject = New-Object -TypeName "System.Collections.ArrayList"
$wsusServersListObject = New-Object -TypeName "System.Collections.ArrayList"
$wsusReportError = "$env:ProgramData\udreports\wsusreport\wsusreport-error.csv"
Get-Item -Path $wsusReportError | Remove-Item -Force -Confirm:$false
$wsusReportData = "$env:ProgramData\udreports\wsusreport\wsusreport-data.csv"
Get-Item -Path $wsusReportData | Remove-Item -Force -Confirm:$false

foreach ($wsusServer in $wsusServerList) {
    
    switch ($wsusServer) {
        '<server name>' { $wsusClientList = '<client name>', '<client name>', '<client name>', '<client name>'; break }
        '<server name>' { $wsusClientList = '<client name>', '<client name>', '<client name>', '<client name>', '<client name>'; break }
        Default  {Write-Warning "WSUS Server $wsusServer could not be contacted"; break }
    }

    if (!(Test-NetConnection -ComputerName $wsusServer -Port 8531 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).TcpTestSucceeded) {
        $wsusServerInfoObject = [PSCustomObject]@{
            Time = Get-Date -Format dd-MM-yyyy-HH-mm-ss
            WSUSserver = $wsusServer
            Level = "ERROR"
            Message = "Could not connect to WSUS server"
        }
        $wsusServersListObject.Add($wsusServerInfoObject)
        continue
    } else {
        $wsusServerConnection = Get-wsusServer -Name $wsusServer -PortNumber 8531 -UseSsl
    }
    
    foreach ($wsusClient in $wsusClientList) {
        $wsusClientInfo = Get-WsusComputer -UpdateServer $wsusServerConnection -NameIncludes $wsusClient
        if (($wsusClientInfo -like 'No computers available.') -or ($null -eq $wsusClientInfo)) {
            $wsusClientInfoObject = [PSCustomObject]@{
                Computer = $wsusClient
                IPAddress = $null
                LastSyncTime = $null
                LastReportedTime = $null
                LastSyncResult = 'Not found'
            }
            $wsusClientsListObject.Add($wsusClientInfoObject)
        } else {
            $wsusClientInfoObject = [PSCustomObject]@{
                Computer = $wsusClientInfo.FullDomainName
                IPAddress = $wsusClientInfo.IPAddress
                LastReportedTime = $wsusClientInfo.LastReportedStatusTime
                LastSyncTime = $wsusClientInfo.LastSyncTime
                LastSyncResult = $wsusClientInfo.LastSyncResult
            }
            $wsusClientsListObject.Add($wsusClientInfoObject)   
        }
    }

    $wsusServer = $null
    $wsusServerConnection = $null
    $wsusClientList = $null
}

$wsusServersListObject | Export-Csv -Path $wsusReportError -NoTypeInformation
$wsusClientsListObject | Export-Csv -Path $wsusReportData -NoTypeInformation