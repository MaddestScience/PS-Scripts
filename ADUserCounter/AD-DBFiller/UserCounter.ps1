##############################################################
##############################################################
##                                                          ##
##                  Script initialization.                  ##
##              Loads all needed information.               ##
##                                                          ##
##############################################################
##############################################################
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\Initialize-script.ps1")
}
catch {
    Write-Output "Error while loading supporting PowerShell Scripts" 
    exit 10
}

##############################################################
##
##  Do not change, unless you know what you are doing.

## Import New AD Users
Function ImportADUsers() {
    $ExcludedCounter = 0
    $CheckCount = 0
    $MailboxCount = 0
    $Users = 0


    ## Excluding user list generation from excludegroup.
    $ExcludedUserList = @()
    foreach ($user in (Get-ADUser -SearchBase “OU=Users,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true)  -AND ($_.memberof -like "*$ExcludeGroup*")}).SamAccountName) {
        $ExcludedUserList += "$user"
    }
    foreach ($ExcludeOU in $ExcludeOUs) {
        ## Generating Excluded User List from OUs
        foreach ($user in (Get-ADUser -SearchBase “$ExcludeOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).SamAccountName) {
            $ExcludedUserList += "$user"
        }
        Write-Host "Excluded OU: $ExcludeOU"
    }
    $ExcludedUserCounter = $ExcludedUserList.Count
    $ExcludedUserCounter = [int]$ExcludedUserCounter

    foreach ($DomainGroupFunction in $DomainGroups.Keys) {
        $DomainGroupName = $DomainGroups.$DomainGroupFunction
        if ($DomainGroupFunction -like "ADUsers") {
            if ($DomainGroupName -like "OU=*") {
                $MailUserCounter = (Get-AdUser -SearchBase "$DomainGroupName,$BaseOU" -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND ($_.msExchMailboxGuid) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
                $ADUsers = Get-ADUser -SearchBase “$DomainGroupName,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}
                $ADUserCounter = ($ADUsers).Count

            } else {
                $MailUserCounter = (Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | Get-ADUser -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND ($_.msExchMailboxGuid) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
                $ADUsers = Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | Get-Aduser -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}
                $ADUserCounter = ($ADUsers).Count
            }
        }
        if ($DomainGroupFunction -like "Office") {
            if ($DomainGroupName -like "OU=*") {
                $OfficeUserCounter = (Get-ADUser -SearchBase “$DomainGroupName,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
            } else {
                $OfficeUserCounter = (Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | Get-Aduser -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
            }
        }
        if ($DomainGroupFunction -eq "RDS"){ 
            if ($DomainGroupName -like "OU=*") {
                $RDSUserCounter = (Get-ADUser -SearchBase “$DomainGroupName,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
            } else {
                $RDSUserCounter = (Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | Get-Aduser -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).Count
            }
        }
    }
    if ($CheckCount -eq "") {
        Write-Output "Checking for new users"  | timestamp
        $CheckCount++
    }

    foreach ($ADUser in $ADUsers) {
        $ad_AccountName = $ADUser.SamAccountName
        $ad_DisplayName = $ADUser.DisplayName
        $ad_Department = $ADUser.Department
        $ad_City = $ADUser.City
        $ad_Created = Get-UnixTimestamp -normaldate $ADUser.Created
        if ($ADUser.LastLogonDate) {
            $ad_LastLogonDate = Get-UnixTimestamp -normaldate $ADUser.LastLogonDate
        } else { $ad_LastLogonDate = '' }
        $ad_UPN = $ADUser.UserPrincipalName
        $ad_Company = $ADUser.Company
        $ad_Exch = $ADUser.msExchMailboxGuid
        ## Check if the user is in the DB
        $data = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * from users where UPN='$ad_UPN';"
        if ($data -eq $null) {
            Write-Output "Adding $ad_UPN to the database"  | timestamp
            ## Add to database 
            run-MySQLQuery -ConnectionString $ConnectionString -Query "INSERT INTO users (AccountName, DisplayName, Company, Department, City, UPN, Created) VALUES ('$ad_AccountName', '$ad_DisplayName', '$Company', '$ad_Department', '$ad_City', '$ad_UPN', '$ad_Created');"
        } elseif ($ad_LastLogonDate) {
            Write-Output "$ad_UPN Still exists, adding LastLogin" | timestamp
            run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE users SET LastLogin='$ad_LastLogonDate' WHERE UPN='$ad_UPN';"
            
        }
    }
    Write-Output "Excluded Users: $ExcludedUserCounter" | timestamp
    Write-Output "Exchange Users: $MailUserCounter" | timestamp
    Write-Output "Office Users: $OfficeUserCounter" | timestamp
    Write-Output "RDS Users: $RDSUserCounter" | timestamp
    Write-Output "AD Users: $ADUserCounter" | timestamp
    run-MySQLQuery -ConnectionString $ConnectionString -Query "INSERT INTO user_counter (Date, Month, Company, ADUsers, OfficeUsers, RDSUsers, MailboxUsers, ExcludedUsers) VALUES ('$now_unixtime', '$currentMonth','$Company','$ADUserCounter','$OfficeUserCounter','$RDSUserCounter','$MailUserCounter', '$ExcludedUserCounter');"
}


## Check for deleted users, and add remove date.
Function CheckSQLUsers() {
    $CheckCount = 0
    $ExcludeCounter = 0

    ## Excluding user list generation from excludegroup.
    $ExcludedUserList = @()
    foreach ($user in (Get-ADUser -SearchBase “OU=Users,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true)  -AND ($_.memberof -like "*$ExcludeGroup*")}).SamAccountName) {
        $ExcludedUserList += "$user"
    }
    foreach ($ExcludeOU in $ExcludeOUs) {
        ## Generating Excluded User List from OUs
        foreach ($user in (Get-ADUser -SearchBase “$ExcludeOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true) -AND (!($_.SamAccountName -in $ExcludedUserList))}).SamAccountName) {
            $ExcludedUserList += "$user"
        }
        Write-Host "Excluded OU: $ExcludeOU"
    }

    foreach ($DomainGroupFunction in $DomainGroups.Keys) {
        $ExcludedCounter = 0
        $DomainGroupName = $DomainGroups.$DomainGroupFunction
        if ($CheckCount -eq "") {
            Write-Output "Checking for deleted users, this may take a while..." | timestamp
            $CheckCount++
        }
        if ($DomainGroupName -like "OU=*") {
            $domainusers = Get-ADUser -SearchBase “$DomainGroupName,$BaseOU” -Filter * -ResultSetSize 5000 -Properties * | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true)}
            Write-Output "Checking users in $DomainGroupName..."  | timestamp
        } else {
            $domainusers = Get-ADGroupMember -Identity "$DomainGroupName" -Recursive | %{Get-ADUser -Identity $_.distinguishedName -Properties * } | Where-Object { ($_.ObjectClass -eq "user") -AND ($_.enabled -eq $true)}
            Write-Output "Checking users in $DomainGroupName..."  | timestamp
        }
        $TotalUserCounter = 0
        $data = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * FROM users WHERE deleted IS NULL AND Company='$Company';"
        foreach($dataitem in $data) {
            $id = $dataitem[0]
            $AccountName = $dataitem[1]
            $DisplayName = $dataitem[2]
            $Company = $dataitem[3]
            $Department = $dataitem[4]
            $City = $dataitem[5]
            $UPN = $dataitem[6]
            $Created = $dataitem[7]
            if ($AccountName -in $ExcludedUserList) {
                run-MySQLQuery -ConnectionString $ConnectionString -Query "DELETE FROM `users` WHERE AccountName='$AccountName' AND Company='$Company';"
                Write-Output "Removed excluded user '$AccountName' from the database." | timestamp
            } elseif ($DefaultAccounts -contains $AccountName) {
                ## DO NOTHING! :P
            } else {
                if ($DomainGroupFunction -eq "ADUsers"){ 
                    if ($domainusers.UserPrincipalName -contains $UPN) {
                        # Write-Output "This user Still exists" -ForegroundColor green

                        $TotalUserCounter++
                    } else { 
                        Write-Output "User '$AccountName' has been deleted..." | timestamp
                        run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE users SET deleted='$now_unixtime' WHERE UPN='$UPN';"
                    }
                }
            }
        }
    }
    Write-Output "Total amount of users:  $TotalUserCounter" | timestamp
}

Function O365Count() {
    $licenties = @()
    Import-Module MSonline
    $Credential = Create-PSCredential
    Connect-MsolService -Credential $Credential -AzureEnvironment AzureCloud
    $licenses = Get-MsolAccountSku
    foreach ($license in $licenses) {
        $license_sku = $license.AccountSkuId
        $license_active = $license.ActiveUnits
        $license_cunits = $license.ConsumedUnits
        if ($license_sku) {
            $license_sku = $license_sku.Split(':')[-1]
            Write-Host $license_sku $license_active $license_cunits
            $licenties += $license_sku
            New-Variable -Name "$($license_sku)_O365" -Value $license_active -Force
        }
    }
    foreach ($licentie in $licenties) {
        Write-Host "Variable '$((Get-Variable -Name "$($licentie)_O365").Name)' have a value of $((Get-Variable -Name "$($licentie)_O365").Value)"
        if ($((Get-Variable -Name "$($licentie)_O365").Name) -in (run-MySQLQuery -ConnectionString $ConnectionString -Query "SHOW columns FROM user_counter").field) {
            Write-Host "$((Get-Variable -Name "$($licentie)_O365").Name) license is in the database"
        } else {
            Write-Host "$((Get-Variable -Name "$($licentie)_O365").Name) license is not in the database"
            try {
                Run-MySQLQuery -connectionString $ConnectionString -query "ALTER TABLE `user_counter` ADD COLUMN $((Get-Variable -Name "$($licentie)_O365").Name) INT(10) NULL DEFAULT '0' AFTER `ExcludedUsers`;"
            } catch {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host "Failed adding $((Get-Variable -Name "$($licentie)_O365").Name) to the database."
                Write-host "Failed with error: $ErrorMessage : $FailedItem"
                Break
            } finally {
                Write-Host "Added $((Get-Variable -Name "$($licentie)_O365").Name) to the database."
            }
        }
        run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE user_counter SET $((Get-Variable -Name "$($licentie)_O365").Name)='$((Get-Variable -Name "$($licentie)_O365").Value)'  WHERE Company='$Company' AND Date='$now_unixtime';"
    }

    $O365Users = Get-MSolUser
    $userlist = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * FROM users WHERE Company='$Company' Order By Created;"
    foreach($userrow in $userlist) {
        ## User information from rows.
        $id = $userrow[0]
        $UPN = $userrow[6]
        $O365lic = $userrow[9]
        foreach ($O365User in $O365Users) {
            $O365UPN = $O365User.UserPrincipalName
            $O365License = ($O365User.Licenses).AccountSkuId
            if (($UPN -eq $O365UPN) -AND $O365License) {
                if ($O365License -like "*:*") {
                    $O365Lic = $O365License
                    $O365Lic = $O365Lic.Split(':')[-1]
                    run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE users SET O365Lic='$O365Lic'  WHERE UPN='$UPN' AND Company='$Company';"
                } else {
                    $O365Lic = ''
                    run-MySQLQuery -ConnectionString $ConnectionString -Query "UPDATE users SET O365Lic='$O365Lic'  WHERE UPN='$UPN' AND Company='$Company';"
                }
            }
        }
    }

}




if ($ImportADUsers) {
    ImportADUsers
}
if($O365) {
    O365Count
}
if ($CheckSQLUsers) {
    CheckSQLUsers
}

Write-Output "Done..." | timestamp