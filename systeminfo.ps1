## System information
$PCinfo = get-wmiobject win32_computersystem
## CPU information
$cpuinfo = Get-WmiObject Win32_Processor -Property "Name",“systemname”,”maxclockspeed”,”addressWidth”,“numberOfCores”, “NumberOfLogicalProcessors”
## GPU information
$gpuinfos = Get-WmiObject win32_VideoController

## Windows information
$winfo = get-wmiobject win32_operatingsystem

## TPM Information
$tpm = Get-WmiObject -Namespace ROOT\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm
$tpmstatus = $tpm.IsEnabled_InitialValue

if ($tpmstatus -eq $true) {
    $tpm_status = "Enabled"
} else { 
    $tpm_status = "Disabled"
}

$geheugen = get-wmiobject win32_computersystem | ForEach-Object {[math]::truncate((get-wmiobject win32_computersystem).TotalPhysicalMemory / 1000MB)}
Write-Host "################################################################"
Write-Host "##                        Systeem Info                        ##"
Write-Host "################################################################"
Write-Host "#"
Write-Host "# Manufacturer:" $PCinfo.Manufacturer
Write-Host "# Model:" $PCinfo.Model 
Write-Host "#"
Write-Host "# CPU Name:" $cpuinfo.Name
Write-Host "# CPU Cores:" $cpuinfo.NumberOfLogicalProcessors
Write-Host "#"
Write-Host "# Memory:" $geheugen"GB"
Write-Host "#"
Write-Host "# Windows ver:" $winfo.Caption
Write-Host "# Windows Build:" $winfo.BuildNumber
Write-Host "# Windows Type:" $PCinfo.SystemType
Write-Host "#"
Write-Host "# Trusted Platform Module:" $tpm_status
Write-Host "#"
Write-Host "################################################################"
Write-Host "##                         Video Info                         ##"
Write-Host "################################################################"

foreach ($gpuinfo in $gpuinfos){
    $gpumem = [math]::truncate($gpuinfo.AdapterRAM / 1000MB)
    Write-Host "#"
    Write-Host "#" $gpuinfo.DeviceID":" $gpuinfo.VideoProcessor
    Write-Host "# Video Memory:" $gpumem"GB"
    if (!$gpuinfo.CurrentHorizontalResolution) { Write-Host "# Status: Not in use." }
    if ($gpuinfo.CurrentHorizontalResolution) { Write-Host "# Resolution:" $gpuinfo.CurrentHorizontalResolution"x"$gpuinfo.CurrentVerticalResolution }
    Write-Host "#"
} 
Write-Host "################################################################"
Write-Host "##                         Owner Info                         ##"
Write-Host "################################################################"
Write-Host "#"
Write-Host "# Domain:" $PCinfo.Domain
Write-Host "# Name:" $PCinfo.Name
Write-Host "# Owner:" $PCinfo.PrimaryOwnerName
Write-Host "# Username:" $PCinfo.UserName
Write-Host "#"
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""

$Drives = Get-WMIObject Win32_LogicalDisk
foreach ($Drive in $Drives) {
    $driveid = $Drive.DeviceID
    
    if($Drive.DriveType -eq '5') {
    $return1 = 'Disabled'
    } else {
    $ProtectionState = Get-WmiObject -Namespace ROOT\CIMV2\Security\Microsoftvolumeencryption -Class Win32_encryptablevolume -Filter "DriveLetter = '$driveid'"
    
    switch ($ProtectionState.GetProtectionStatus().protectionStatus){
        ('0'){$return1 = 'Disabled'}
        ('1'){$return1 = 'Enabled'}
        ('2'){$return1 = 'Unknown'}
        default {$return1 = 'Unknown'}
    }#EndSwitch
    }

    switch ($Drive.DriveType){
        ('3'){$return = 'Hard Drive'}
        ('5'){$return = 'Optic Drive'}
        ('2'){$return = 'USB Device'}
        default {$return = 'Unknown'}
    }
    $size = [math]::truncate($Drive.Size / 1000MB)
    $freespace = [math]::truncate($Drive.FreeSpace / 1000MB)
    Write-Host "################################################################"
    Write-Host "##                        Device ID:" $Drive.DeviceID "                      ##"
    Write-Host "################################################################"
    Write-Host "#"
    Write-Host "# Device ID:" $Drive.DeviceID
    Write-Host "# Device Type:" $return
    Write-Host "# Volume Name:" $Drive.VolumeName
    Write-Host "# Size:" $size"GB"
    Write-Host "# Free Space:" $freespace"GB"
    Write-Host "# Bitlocker:" $return1
    Write-Host "#"

}
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
