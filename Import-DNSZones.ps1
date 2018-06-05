## Import every domain, that has been exported to an specific location into MS-DNS Server.

$WebClient = New-Object System.Net.WebClient
$Path = "C:\dns\dnszones.txt"
Write-Host "Downloading DNS Zone file '$Path'" -ForegroundColor Green
$Url = "https://www.<domain.tld>/dnszones.txt"
$WebClient.DownloadFile( $url, $path )

Write-Host "Get DNS Zones from file '$path'" -ForegroundColor Green
$dnszones = Get-Content $path
$PrimaryDNSServer = <MasterSerer-IPAddress>

foreach ($zone in $dnszones) {
    #Check for Zone Existence
    $ZoneFound = $false
    Try {
        if ((get-dnsserverzone $zone | select ZoneName) -ne $null) { 
            $ZoneFound = $true
        }
    }
    Catch {
        $ZoneFound = $false
    }
    Finally {
        If ($ZoneFound) { 
            Write-Host "Zone '$zone' exists"  -ForegroundColor Green
        }
        If (!($ZoneFound)) {
            Write-Host "Zone '$zone' does not exist" -ForegroundColor Red
        }
        if (!($ZoneFound)) {
           Try {
                #Create the zone
                Write-Host "Attempting to create zone '$zone'." -ForegroundColor Green
                $ZoneFound = $true
                Add-DnsServerSecondaryZone -Name "$zone" -ZoneFile "$zone.dns" -MasterServers $PrimaryDNSServer
            } Catch {
                Write-Host "Could not create zone '$zone'." -ForegroundColor Red
                Write-Host "An error occurred creating the DNS zone for '$zone'." -ForegroundColor Red
                $ErrorOccurred = $true
                $ZoneFound = $false
            }
            Finally {
                If ($ZoneFound) {
                    Write-Host "Zone '$zone' has been created." -ForegroundColor Green
                }
                If (!($ZoneFound)) {
                    Write-Host "Zone '$zone' had not been created." -ForegroundColor Red
                }
            }
        }
    }
}