
$path = 'C:\dns\'
## Preparation
if (!(Test-Path $path)) {
    New-Item $path -ItemType Directory
    Write-Host "Created folder '$path'" -ForegroundColor Red
} else {
   Write-Host "Folder '$path' existed" -ForegroundColor Green
}

if (!(Test-Path $path\dnszones.txt)) {
   New-Item $path\dnszones.txt -ItemType File
   Write-Host "Created file '$path\dnszone.txt'"  -ForegroundColor Red
} else { 
    Clear-Content $path\dnszones.txt
    Write-Host "Emptied file '$path\dnszones.txt'" -ForegroundColor Green
}

Write-Host "Getting DNS Zones" -ForegroundColor Green
$zones = (Get-DnsServerZone).ZoneName

Write-Host "Writing zones to '$path\dnszones.txt'" -ForegroundColor Green
foreach ($zone in $zones) {
    $contains = '0.in-addr.arpa','127.in-addr.arpa','255.in-addr.arpa','TrustAnchors'
    if ($contains -contains $zone){ } else {
        $zone | Out-File $path\dnszones.txt -Append
    }
}