$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\config.ps1")
}
catch {
    Write-Host "Error while loading supporting PowerShell Scripts" 
    exit 10
}

#####################################################################################
##   Do not change things under this line if you do not know what you are doing!!  ##
#####################################################################################




#[void][system.reflection.Assembly]::LoadFrom("C:\milo\scripts\MySQL.Data.dll")
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
			# You could also could use a direct Link to the DLL File
			# $mySQLDataDLL = "C:\scripts\mysql\MySQL.Data.dll"
			# [void][system.reflection.Assembly]::LoadFrom($mySQLDataDLL)
			[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
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
			Write-Host "Could not run MySQL Query" $Error[0]	
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

Function ImportADUsers() {
    $CheckCount = 0
    foreach ($DomainGroupFunction in $DomainGroups.Keys) {
        $DomainGroupName = $DomainGroups.$DomainGroupFunction
        if ($DomainGroupName -like "OU=*") {
            $domainusers = Get-ADUser -SearchBase “$DomainGroupName,$BaseOU” -Filter * -ResultSetSize 5000 -Properties *
        } else {
            $domainusers = Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | Get-Aduser -Properties *
        }
        $Users = 0
        foreach ($domainuser in $domainusers) {
            $ad_AccountName = $domainuser.SamAccountName
            $ad_DisplayName = $domainuser.DisplayName
            $ad_Department = $domainuser.Department
            $ad_City = $domainuser.City
            $ad_Created = $domainuser.Created
            $ad_UPN = $domainuser.UserPrincipalName
            $ad_Company = $domainuser.Company
            if (!($DefaultAccounts -contains $ad_AccountName)) {
                $Users++
                if ($CheckCount -eq "") {
                    Write-Host "Checking for new users" -ForegroundColor Green
                    $CheckCount++
                }
                ## Check if the user is in the DB
                $data = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * from users where UPN='$ad_UPN';"
                if ($data -eq $null) {
                    Write-Host "Adding $ad_UPN to the database" -ForegroundColor Red
                    ## Write-Host "$ad_accountName, $ad_Displayname, $ad_Company, $ad_Department, $ad_City, $ad_upn, $ad_Created"
                    ## Add to database 
                    run-MySQLQuery -ConnectionString $ConnectionString -Query "INSERT INTO users (AccountName, DisplayName, Company, Department, City, UPN, Created) VALUES ('$ad_AccountName', '$ad_DisplayName', '$ad_Company', '$ad_Department', '$ad_City', '$ad_UPN', '$ad_Created');"
                }
            }
        }
        Write-Host "$DomainGroupFunction ($DomainGroupName) has $Users users."
        if ($DomainGroupFunction -eq "MailBox" -Or $DomainGroupFunction -eq "OU=Mailbox" ) {
            $MailboxCount = $Users
        }
        if ($DomainGroupFunction -eq "RDS") {
            $RDSCount = $Users
        }
        if ($DomainGroupFunction -eq "Office") {
            $OfficeCount = $Users
        }
        if ($DomainGroupFunction -eq "ADUsers") {
            $ADUserCount = $Users
        }
    }
    $now = Get-Date -UFormat "%m/%d/%Y %H:%M:%S"
    Write-Host "Mailbox Users: $MailboxCount, Office Users: $OfficeCount, RDS Users: $RDSCount, AD Users: $ADUserCount" -ForegroundColor Magenta
    run-MySQLQuery -ConnectionString $ConnectionString -Query "INSERT INTO user_counter (Date, Company, ADUsers, OfficeUsers, RDSUsers, MailboxUsers) VALUES ('$now','$Company','$ADUserCount','$OfficeCount','$RDSCount','$MailboxCount');"
}
Function CheckSQLUsers() {
    $CheckCount = 0
    foreach ($DomainGroup in $DomainGroups){
        if ($CheckCount -eq "") {
            Write-Host "Checking for deleted users." -ForegroundColor Green
            $CheckCount++
        }
        $TotalUserCounter = 0
        $data = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * FROM users WHERE deleted IS NULL;"
        foreach($dataitem in $data) {
            $id = $dataitem[0]
            $AccountName = $dataitem[1]
            $DisplayName = $dataitem[2]
            $Company = $dataitem[3]
            $Department = $dataitem[4]
            $City = $dataitem[5]
            $UPN = $dataitem[6]
            $Created = $dataitem[7]
            if (!($DefaultAccounts -contains $AccountName)) {
                ## Write-Host "$id, $AccountName, $DisplayName, $Company, $Department, $City, $UPN, $Created"
                if (Get-ADUser -Filter {UserPrincipalName -eq $UPN}) {
                    if ($CheckCount -eq 1) {
                        Write-Host "Counting all users." -ForegroundColor Green
                        $CheckCount++
                    }
                    ## Write-Host "This user Still exists" -ForegroundColor green
                    $TotalUserCounter++
                } else { 
                    Write-Host "This user has been deleted" -ForegroundColor red
                    $now = Get-Date -UFormat "%m/%d/%Y %H:%M:%S"
                    run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE users SET deleted='$now' WHERE UPN='$UPN';"
                }
            }
        }
    }
}



ImportADUsers
CheckSQLUsers
