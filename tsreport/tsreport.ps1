# Stop ts-dashboard site before loading data
Stop-IISSite -Name 'ts-dashboard' -Confirm:$false

# Data files
$tsUsersDataPath = "$env:ProgramData\udreports\tsreport\ts-users-data.csv"
$tsGlobalDataPath = "$env:ProgramData\udreports\tsreport\ts-global-data.csv"
$tsCountryDataPath = "$env:ProgramData\udreports\tsreport\ts-country-data.csv"

# Removing previsouly created data files
Get-Item -Path $tsUsersDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $tsGlobalDataPath | Remove-Item -Force -Confirm:$false
Get-Item -Path $tsCountryDataPath | Remove-Item -Force -Confirm:$false

# Connection to database
$sqlServerInfo = "<server name>"
$sqlInstanceInfo = Get-SqlInstance -Path "SQLSERVER:\SQL\<server name>\DEFAULT"
$sqlDatabaseInfo = "TSLicensesReport"

# Common variables
$tsLicenseServer = "<server name>"
$tsLicensesList = Get-CimInstance -ClassName 'Win32_TSIssuedLicense' -ComputerName $tsLicenseServer
$tsLicensesPackList = Get-CimInstance -ClassName 'Win32_TSLicenseKeyPack' -ComputerName $tsLicenseServer
$tsReportedOn = Get-Date -Format yyyy-MM-dd

# Gathering users data
$tsUserDataArray = New-Object -TypeName "System.Collections.ArrayList"
[int]$tsLicenseVersion = 0000
$tsLicensesList | ForEach-Object {
    $tsUser = $_.sIssuedToUser.Split('\')[1]
    $tsIssuedOn = $_.IssueDate.ToString("yyyy-MM-dd")
    $tsExpireOn = $_.ExpirationDate.ToString("yyyy-MM-dd")
    $tsLicenseStatus = $_.LicenseStatus
    if ($_.KeyPackId -eq 3) {
        $tsLicenseVersion = 2012
    } elseif ($_.KeyPackId -eq 5) {
        $tsLicenseVersion = 2008
    } else {
        $tsLicenseVersion = 0000
    }
    $tsUserCountry = (Get-ADUser -Identity $_.sIssuedToUser.Split('\')[1]).DistinguishedName.Split(',')[(-1) + 3].Remove(0, 3)

    $tsLicenseObj = [PSCustomObject]@{
        User = $tsUser
        IssuedOn = $tsIssuedOn
        ExpireOn = $tsExpireOn
        LicenseStatus = $tsLicenseStatus
        LicenseVersion = $tsLicenseVersion
        Country = $tsUserCountry
        ReportedOn = $tsReportedOn
    }
    $tsUserDataArray.Add($tsLicenseObj)

    $userDataQuery = "
	INSERT INTO ts_users (
		user_name,
	    issued_on,
	    expire_on,
        license_status,
        license_version,
	    country,
	    reported_on
	) VALUES (
		'$tsUser',
        '$tsIssuedOn',
        '$tsExpireOn',
        '$tsLicenseStatus',
        '$tsLicenseVersion',
        '$tsUserCountry',
        '$tsReportedOn'
	);
    "

    Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $userDataQuery
    
}
$tsUserDataArray | Export-Csv -Path $tsUsersDataPath -NoTypeInformation

# Gathering global data
$tsGlobalDataArray = New-Object -TypeName "System.Collections.ArrayList"
$tsLicensesPackList | ForEach-Object {
    if ($_.KeyPackId -ne 2) {
        $tsAvailableLicNum = $_.AvailableLicenses
        $tsIssuedLicNum = $_.IssuedLicenses
        $tsTotalLicNum = $_.TotalLicenses
        $tsKeyPackId = $_.KeyPackId
        $tsProductVersion = $_.ProductVersion

        $tsGlobalDataObj = [PSCustomObject]@{
            AvailableLicenses = $tsAvailableLicNum
            IssuedLicenses = $tsIssuedLicNum
            TotalLicenses = $tsTotalLicNum
            KeyPackId = $tsKeyPackId
            ProductVersion = $tsProductVersion
            ReportedOn = $tsReportedOn
        }
        $tsGlobalDataArray.Add($tsGlobalDataObj)
    
        $globalDataQuery = "
        INSERT INTO ts_global_stats (
            available_lic_num,
            issued_lic_num,
            total_lic_num,
            key_pack_id,
            product_version,
            reported_on
        ) VALUES (
            '$tsAvailableLicNum',
            '$tsIssuedLicNum',
            '$tsTotalLicNum',
            '$tsKeyPackId',
            '$tsProductVersion',
            '$tsReportedOn'
        );
        "

        Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $globalDataQuery
    }
    
}
$tsGlobalDataArray | Export-Csv -Path $tsGlobalDataPath -NoTypeInformation

# Gathering country data
$tsCountryDataArray = New-Object -TypeName "System.Collections.ArrayList"
$tsCcCountryList = $tsUserDataArray | Group-Object -Property Country | Select-Object -ExpandProperty Name

foreach ($tsCcCountry in $tsCcCountryList) {
    $tsCcIssuedLicNum = ($tsUserDataArray | Group-Object -Property Country | Where-Object {$_.Name -like $tsCcCountry}).Count
    $tsCcLicenseVersionWe = ($tsUserDataArray | ForEach-Object {$_ | Where-Object {($_.Country -like $tsCcCountry) -and ($_.LicenseVersion -eq '2008')}}).Count
    $tsCcLicenseVersionWt = ($tsUserDataArray | ForEach-Object {$_ | Where-Object {($_.Country -like $tsCcCountry) -and ($_.LicenseVersion -eq '2012')}}).Count

    $tsCcCountryObj = [PSCustomObject]@{
        Country = $tsCcCountry
        IssuedLicNum = $tsCcIssuedLicNum
        LicenseVersionWe = $tsCcLicenseVersionWe
        LicenseVersionWt = $tsCcLicenseVersionWt
        ReportedOn = $tsReportedOn
    }
    $tsCountryDataArray.Add($tsCcCountryObj)

    $countryDataQuery = "
    INSERT INTO ts_country_stats (
        country,
	    number_of_issued_licenses,
	    license_version_2008,
	    license_version_2012,
	    reported_on
    ) VALUES (
        '$tsCcCountry',
        '$tsCcIssuedLicNum',
        '$tsCcLicenseVersionWe',
        '$tsCcLicenseVersionWt',
        '$tsReportedOn'
    );
    "

    Invoke-Sqlcmd -HostName $sqlServerInfo -ServerInstance $sqlInstanceInfo -Database $sqlDatabaseInfo -Query $countryDataQuery
}
$tsCountryDataArray | Export-Csv -Path $tsCountryDataPath -NoTypeInformation

# Start ts-dashboard site after loading data
Start-IISSite -Name 'ts-dashboard'