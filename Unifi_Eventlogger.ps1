param(
	[string]$server = '<UniFi Controller hostname/ip>',
	[string]$port = '<UniFi Controller port>',
	[string]$site = '<UniFi Controller siteid>',
	[string]$username = '<UniFi Controller Username>',
	[string]$password = '<UniFi Controller Password>',
	[switch]$debug = $false
)

## Time converter Unix-to-Readable:
Function Convert-FromUnixDate ($UnixDate) {
   [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
}

## Check if the Event-Log source does exist, if not, make it.
if (! [System.Diagnostics.Eventlog]::SourceExists("Unifi-Logger")) {
    New-EventLog –LogName Application –Source “Unifi-Logger”
}

#Ignore SSL Errors
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}  

#Define supported Protocols
[System.Net.ServicePointManager]::SecurityProtocol = @("Tls12","Tls11","Tls","Ssl3")


# Confirm Powershell Version.
if ($PSVersionTable.PSVersion.Major -lt 3) {
	Write-Output "<prtg>"
	Write-Output "<error>1</error>"
	Write-Output "<text>Powershell Version is $($PSVersionTable.PSVersion.Major) Requires at least 3. </text>"
	Write-Output "</prtg>"
	Exit
}

# Create $controller and $credential using multiple variables/parameters.
[string]$controller = "https://$($server):$($port)"
[string]$credential = "`{`"username`":`"$username`",`"password`":`"$password`"`}"

# Start debug timer
$queryMeasurement = [System.Diagnostics.Stopwatch]::StartNew()

# Perform the authentication and store the token to myWebSession
try {
$null = Invoke-Restmethod -Uri "$controller/api/login" -method post -body $credential -ContentType "application/json; charset=utf-8"  -SessionVariable myWebSession
}catch{
    	Write-Output ""
	    Write-Output "Error: 1 // Can not login..."
	    Write-Output "API Query Failed: $($_.Exception.Message)"
	Exit
}

try {
$getsites = Invoke-Restmethod -Uri "$controller/api/self/sites" -WebSession $myWebSession
}catch{
    	Write-Output ""
	    Write-Output "Error: 2 // Can not get Sites"
	    Write-Output "API Query Failed: $($_.Exception.Message)"
	Exit
}
foreach ($entry in $getsites.data) {
    $sitename = $entry.name
    $sitedesc = $entry.desc
    Write-Host "#########################################################################"
    Write-Host "#########################################################################"
    Write-Host "###"
    Write-Host "###   Site: $sitedesc"
    Write-Host "###   $controller/manage/site/$sitename/dashboard "
    Write-Host "###"
    Write-Host "#########################################################################"
    Write-Host "#########################################################################"

    #Query API providing token from first query.
    try {
        $jsonresultat = Invoke-Restmethod -Uri "$controller/api/s/$sitename/stat/device/" -WebSession $myWebSession
    }catch{
    	Write-Output ""
	    Write-Output "Error: 3 // Can not get Site information for '$sitedesc'"
	    Write-Output "API Query Failed: $($_.Exception.Message)"
	    Exit
    }
    # Load File from Debug Log
    # $jsonresultatFile = Get-Content '.\unifi_sensor2017-15-02-05-42-24_log.json'
    # $jsonresultat = $jsonresultatFile | ConvertFrom-Json

    # Stop debug timer
    $queryMeasurement.Stop()


    # Iterate jsonresultat and count the number of AP's. 
    #   $_.state -eq "1" = Connected 
    #   $_.type -like "uap" = Access Point
    #   $_.type -like "ugw" = Gateways
    #   $_.type -like "usw" = Switches


    ## Check if all Access-Points are connected, if not generate eventlog.
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -eq "uap" })) { 
        $devicename = $entry.name
        Write-Host "Unifi Access Point Connected: $devicename" -ForegroundColor Green
    }
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -eq "uap"})) {
        $devicename = $entry.name
        $lastseen = Convert-FromUnixDate $entry.last_seen
        Write-Host "Unifi Access Point Disconnected: $devicename" -ForegroundColor red
        Write-EventLog –LogName Application –Source "Unifi-Logger” –EntryType Error –EventID 1 –Message “Unifi Access Point Disconnected from the Unifi Controller: $devicename Last seen on: $lastseen”
    }

    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -eq "ugw"})) { 
        $devicename = $entry.name
        Write-Host "Unifi Gateway Connected: $devicename" -ForegroundColor Green
    }
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -eq "ugw"})) {
        $devicename = $entry.name
        $lastseen = Convert-FromUnixDate $entry.last_seen
        Write-Host "Unifi Gateway Disconnected: $devicename" -ForegroundColor red
        Write-EventLog –LogName Application –Source "Unifi-Logger” –EntryType Error –EventID 1 –Message “Unifi Gateway Disconnected from the Unifi Controller: $devicename Last seen on: $lastseen”
    }


    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -eq "usw"})) { 
        $devicename = $entry.name
        Write-Host "Unifi Switch Connected: $devicename" -ForegroundColor Green
    }
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -eq "usw"})) {
        $devicename = $entry.name
        $lastseen = Convert-FromUnixDate $entry.last_seen
        Write-Host "Unifi Switch Disconnected: $devicename" -ForegroundColor red
        Write-EventLog –LogName Application –Source "Unifi-Logger” –EntryType Error –EventID 1 –Message “Unifi Switch Disconnected from the Unifi Controller: $devicename Last seen on: $lastseen”
    }

    Write-Host ""

    ## Connected counter

    $uapCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -like "uap"})){
    	$uapCount ++
    }

    $ugwCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -like "ugw"})){
    	$ugwCount ++
    }

    $uswCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "1" -and $_.type -like "usw"})){
    	$uswCount ++
    }


    ## Disconnect count
    $uapDCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -like "uap"})){
    	$uapDCount ++
    }
    
    $ugwDCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -like "ugw"})){
    	$ugwDCount ++
    }

    $uswDCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.state -eq "0" -and $_.type -like "usw"})){
    	$uswDCount ++
    }


    $uapUpgradeable = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.type -like "uap" -and $_.upgradable -eq "true"})){
    	$uapUpgradeable ++
    }


    $ugwUpgradeable = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.type -like "ugw" -and $_.upgradable -eq "true"})){
    	$ugwUpgradeable ++
    }

    $uswUpgradeable = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.type -like "usw" -and $_.upgradable -eq "true"})){
    	$uswUpgradeable ++
    }

    $userCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.type -like "uap"})){
    	$userCount += $entry.'num_sta'
    }

    $guestCount = 0
    Foreach ($entry in ($jsonresultat.data | where-object { $_.type -like "uap"})){
    	$guestCount += $entry.'guest-num_sta'
    }

    #Write Results
    Write-Host "Connected devices:" -ForegroundColor red
    Write-Host "Access Points Connected: $($uapCount)"
    Write-Host "Gateways Connected: $($ugwCount)"
    Write-Host "Switches Connected: $($uswCount)"
    Write-Host ""
    Write-Host "Disconnected devices:" -ForegroundColor red
    Write-Host "Access Points Disconnected: $($uapDCount)"
    Write-Host "Gateways Disconnected: $($ugwDCount)"
    Write-Host "Switches Disconnected: $($uswDCount)"
    Write-Host ""
    Write-Host "Upgradable devices:" -ForegroundColor red
    Write-Host "Access Points Upgradeable: $($uapUpgradeable)"
    Write-Host "Gateways Upgradeable: $($ugwUpgradeable)"
    Write-Host "Switches Upgradeable: $($uswUpgradeable)"
    Write-Host ""
    Write-Host "Clients (Total): $($userCount)" -ForegroundColor Green
    Write-Host "Guests: $($guestCount)" -ForegroundColor Green

    Write-Host ""
    Write-Host ""

    # Write JSON file to disk when -debug is set. For troubleshooting only.
    if ($debug){
    	[string]$logPath = ((Get-ItemProperty -Path "hklm:SOFTWARE\Wow6432Node\Paessler\PRTG Network Monitor\Server\Core" -Name "Datapath").DataPath) + "Logs (Sensors)\"
    	$timeStamp = (Get-Date -format yyyy-dd-MM-hh-mm-ss)

	   $json = $jsonresultat | ConvertTo-Json
	   $json | Out-File $logPath"unifi_sensor$($timeStamp)_log.json"
    }
}