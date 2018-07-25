#####################################################################################
##   Do not change things under this line if you do not know what you are doing!!  ##
#####################################################################################
## Loading client config.

## Time calculations
Function Get-UnixTimestamp($normalDate) {
    $DateTime = Get-Date #or any other command to get DateTime object
    ([DateTimeOffset]$normalDate).ToUnixTimeSeconds()
}

Function Get-NormalDate ($UnixDate) {
   [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
}


$currentMonth = Get-Date -UFormat "%m"
#$currentMonth = $currentMonth-1

$dayoftheMonth = Get-Date -UFormat "%d"

$now = Get-Date -UFormat "%m/%d/%Y %H:%M:%S"
$now_unixtime = Get-UnixTimestamp -normalDate $now
$European_date = Get-Date (Get-NormalDate -UnixDate $now_unixtime) -UFormat "%d/%m/%Y"



## Loading Configuration file.
try {
    . ("$ScriptDirectory\db-credentials.ps1")
}
catch {
    Write-Output "Error while loading db credentials" 
    exit 10
}
try {
    . ("$ScriptDirectory\config.ps1")
}
catch {
    Write-Output "Error while loading supporting PowerShell Scripts" 
    exit 10
}


if($O365)
{
if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host "Module exists"
} else {
    Write-Host "Module does not exist"
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force 
    Find-PackageProvider -Name "nuget" | Import-PackageProvider
    Install-Module MSOnline -Confirm:$False
}
}


## Switch debugging (switch found in config)
if ($debug) {
    $WarningPreference = 'Continue'
    $DebugPreference = 'Continue'
    $ProgressPreference = 'Continue'
    $VerbosePreference = 'Continue'
} else {
    $WarningPreference = 'SilentlyContinue'
    $DebugPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $VerbosePreference = 'SilentlyContinue'
}

## Timestamp function:
if ($timestamps) {
    filter timestamp {"$(Get-Date -Format G): $_"}
} else {
    filter timestamp {"$_"}
}

## Loading mysql import data.
try {
    . ("$ScriptDirectory\importMySQL.Data.ps1")
}
catch {
    Write-Output "Error while loading supporting PowerShell Scripts" 
    exit 10
}


## Create PSCredentials
Function Create-PSCredential() { 
    $klant = get-content "$O365Klanten\$O365Tenant.txt"    
    $username = $klant[0]
    $password = $klant[1]
    if((($password | Measure-Object -Character).Characters) -lt 30) {
        $password = $password | ConvertTo-SecureString -AsPlainText -Force 
        $secureStringText = $password | ConvertFrom-SecureString
        $klantbestand = "$O365Tenant.txt"
        $sh_file = "$O365Klanten\$O365Tenant.txt"
        set-Content $sh_File $username
        add-Content $sh_File $secureStringText
        $password = $secureStringText
    }
    $SecurePassword = ConvertTo-SecureString $password -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($username, $SecurePassword) 
    Return $Credential
} 


## Run MySQL Queries:
Function Run-MySQLQuery {
	Param(
        [Parameter(
            Mandatory = $true,
            ParameterSetName = '',
            ValueFromPipeline = $true)]
            [string]$query,   
		[Parameter(
            Mandatory = $true,
            ParameterSetName = '',
            ValueFromPipeline = $true)]
            [string]$connectionString
        )
	Begin {
		Write-Verbose "Starting Begin Section"		
    }
	Process {
		Write-Verbose "Starting Process Section"
		try {
			# load MySQL driver and create connection
			Write-Verbose "Create Database Connection"
			$connection = New-Object MySql.Data.MySqlClient.MySqlConnection
			$connection.ConnectionString = $ConnectionString
            
			Write-Verbose "Open Database Connection"
           
			$connection.Open()
			
			# Run MySQL Querys
			Write-Verbose "Run MySQL Querys"
			$command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
			$dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
			$dataSet = New-Object System.Data.DataSet
			$recordCount = $dataAdapter.Fill($dataSet, "data")
			$dataSet.Tables["data"]
            ## $data = $dataSet.Tables["data"]
		}		
		catch {
			Write-Output "Could not run MySQL Query" $Error[0]	
		}	
		Finally {
			Write-Verbose "Close Connection"
			$connection.Close()
		}
    }
	End {
		Write-Verbose "Starting End Section"
	}
}