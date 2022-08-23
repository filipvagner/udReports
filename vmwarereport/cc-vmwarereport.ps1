# Data files
$commonDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-common-data.csv"
$pdcDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-pdc-data.csv"
$pdcHostDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-pdc-host-data.csv"
$pdcDatastoreDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-pdc-datastore-data.csv"
$pdcSnapshotDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-pdc-vm-snapshots.csv"
$drcDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-drc-data.csv"
$drcHostDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-drc-host-data.csv"
$drcDatastoreDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-drc-datastore-data.csv"
$drcSnapshotDataPath = "$env:ProgramData\udreports\cc\vmware-data\cc-drc-vm-snapshots.csv"

# Removing previsouly created data files
Get-Item -Path $commonDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $pdcDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $pdcHostDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $pdcDatastoreDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $pdcSnapshotDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $drcDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $drcHostDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $drcDatastoreDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $drcSnapshotDataPath | Remove-Item -Force -Confirm:$false

# Connection to vCenter
$pdcVcenterServer = "<server name>"
$drcVcenterServer = "<server name>"
$UserNameToAccessVcenter = '<user name>@<domain>'
$EncryptedPasswordToAccessVcenter = Get-Content -Path "$env:ProgramData\udreports\eid.txt" | ConvertTo-SecureString
$CredentialsToAccessVcenter = New-Object -TypeName System.Management.Automation.PSCredential($UserNameToAccessVcenter, $EncryptedPasswordToAccessVcenter)
Connect-VIServer -Server $pdcVcenterServer, $drcVcenterServer -Credential $CredentialsToAccessVcenter -ErrorAction Stop

# Common variables
$commonData = [PSCustomObject]@{
    countryName = "<country name>"
    dateReportCreated = Get-Date -Format dd-MM-yyyy
    timeReportCreated = Get-Date -Format HH:mm:ss
}
$commonData | Export-Csv -Path $commonDataPath -NoTypeInformation

# PDC variables
## Gathering overview data
$pdcHostCluster = "<cluster name>"
$a = 0
$b = 0
Get-Cluster -Name $pdcHostCluster | Get-VMHost | ForEach-Object {$a = $a + $_.MemoryTotalGB}
Get-Cluster -Name $pdcHostCluster | Get-VMHost | ForEach-Object {$b = $b + $_.MemoryUsageGB}
$pdcHostClusterUsagePercent =  [System.Math]::Round(($b * 100) / $a, 2)
$pdcDsCluster = "<cluster name>"
[String]$pdcDSClusterUsagePercent = (Get-DatastoreCluster -Name $pdcDsCluster | Select-Object @{Name='UsedPercents'; Expression = {[System.Math]::Round((((($_.CapacityGB)-($_.FreespaceGB))*100)/$_.CapacityGB),2)}}).UsedPercents
$pdcData = [PSCustomObject]@{
    pdcVcenterServer = $pdcVcenterServer
    pdcHostCluster = $pdcHostCluster
    pdcHostClusterUsagePercent = $pdcHostClusterUsagePercent
    pdcDsCluster = $pdcDsCluster
    pdcDSClusterUsagePercent = $pdcDSClusterUsagePercent
}
$pdcData | Export-Csv -Path $pdcDataPath -NoTypeInformation

## Gathering host data
$pdcVmHostArray = New-Object -TypeName "System.Collections.ArrayList"
Get-Cluster -Name $pdcHostCluster | Get-VMHost | ForEach-Object {
    $pdcVmHostObj = [PSCustomObject]@{
        VMHostName = $_.Name;
        VMHostCpuUsage = [System.Math]::Round($_.CpuUsageMhz/1000, 0);
        VMHostCpuTotal = [System.Math]::Round($_.CpuTotalMhz/1000, 0);
        VMHostMemoryUsage = [System.Math]::Round($_.MemoryUsageGB, 0);
        VMHostMemoryTotal = [System.Math]::Round($_.MemoryTotalGB, 0);
    }
    $pdcVmHostArray.Add($pdcVmHostObj)
}
$pdcVmHostArray | Export-Csv -Path $pdcHostDataPath -NoTypeInformation

## Gathering datastore data
$pdcDsArray = New-Object -TypeName "System.Collections.ArrayList"
Get-DatastoreCluster -Name $pdcDsCluster | Get-Datastore | ForEach-Object {
    $pdcDsObj = [PSCustomObject]@{
        DSName = $_.Name;
        DSSpaceUsed = [System.Math]::Round(($_.CapacityGB/1024)-($_.FreeSpaceGB/1024), 2);
        DSSpaceLeft = [System.Math]::Round($_.FreeSpaceGB/1024, 2);
    }
    $pdcDsArray.Add($pdcDsObj)
}
$pdcDsArray | Export-Csv -Path $pdcDatastoreDataPath -NoTypeInformation

## Gathering snapshot data
$pdcVmsList = Get-Cluster -Name $pdcHostCluster | Get-VM
Get-VM $pdcVmsList | Get-Snapshot | Select-Object VM, Name, Created, Description, @{Name='SnapshotSizeMB'; Expression={($_.SizeMB) -as [int]}} | Export-Csv -Path $pdcSnapshotDataPath -NoTypeInformation

# DRC variables
## Gathering overview data
$drcHostCluster = "<cluster name>"
$a = 0
$b = 0
Get-Cluster -Name $drcHostCluster | Get-VMHost | ForEach-Object {$a = $a + $_.MemoryTotalGB}
Get-Cluster -Name $drcHostCluster | Get-VMHost | ForEach-Object {$b = $b + $_.MemoryUsageGB}
$drcHostClusterUsagePercent =  [System.Math]::Round(($b * 100) / $a, 2)
$drcDsCluster = "<cluster name>"
[String]$drcDSClusterUsagePercent = (Get-DatastoreCluster -Name $drcDsCluster | Select-Object @{Name ='UsedPercents'; Expression = {[System.Math]::Round((((($_.CapacityGB)-($_.FreespaceGB))*100)/$_.CapacityGB),2)}}).UsedPercents
$drcData = [PSCustomObject]@{
    drcVcenterServer = $drcVcenterServer
    drcHostCluster = $drcHostCluster
    drcHostClusterUsagePercent = $drcHostClusterUsagePercent
    drcDsCluster = $drcDsCluster
    drcDSClusterUsagePercent = $drcDSClusterUsagePercent
}
$drcData | Export-Csv -Path $drcDataPath -NoTypeInformation

## Gathering host data
$drcVmHostArray = New-Object -TypeName "System.Collections.ArrayList"
Get-Cluster -Name $drcHostCluster | Get-VMHost | ForEach-Object {
    $drcVmHostObj = [PSCustomObject]@{
        VMHostName = $_.Name;
        VMHostCpuUsage = [System.Math]::Round($_.CpuUsageMhz/1000, 0);
        VMHostCpuTotal = [System.Math]::Round($_.CpuTotalMhz/1000, 0);
        VMHostMemoryUsage = [System.Math]::Round($_.MemoryUsageGB, 0);
        VMHostMemoryTotal = [System.Math]::Round($_.MemoryTotalGB, 0);
    }
    $drcVmHostArray.Add($drcVmHostObj)
}
$drcVmHostArray | Export-Csv -Path $drcHostDataPath -NoTypeInformation

## Gathering datastore data
$drcDsArray = New-Object -TypeName "System.Collections.ArrayList"
Get-DatastoreCluster -Name $drcDsCluster | Get-Datastore | ForEach-Object {
    $drcDsObj = [PSCustomObject]@{
        DSName = $_.Name;
        DSSpaceUsed = [System.Math]::Round(($_.CapacityGB/1024)-($_.FreeSpaceGB/1024), 2);
        DSSpaceLeft = [System.Math]::Round($_.FreeSpaceGB/1024, 2);
    }
    $drcDsArray.Add($drcDsObj)
}
$drcDsArray | Export-Csv -Path $drcDatastoreDataPath -NoTypeInformation

## Gathering host data
$drcVmsList = Get-Cluster -Name $drcHostCluster | Get-VM
Get-VM $drcVmsList | Get-Snapshot | Select-Object VM, Name, Created, Description, @{Name='SnapshotSizeMB'; Expression={($_.SizeMB) -as [int]}} | Export-Csv -Path $drcSnapshotDataPath -NoTypeInformation

# Disconnect from vCenter servers
Disconnect-VIServer -Server $pdcVcenterServer, $drcVcenterServer -Confirm:$false