#####################################################################################
##   Do not change things under this line if you do not know what you are doing!!  ##
#####################################################################################
## Loading client config.
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\Initialize-script.ps1")
}
catch {
    Write-Output "Error while loading supporting PowerShell Scripts" 
    exit 10
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
        }
    }

    foreach ($licentie in $licenties) {
        New-Variable -Name "$($licentie)_O365" -Value $license_active -Force
    }
    foreach ($licentie in $licenties) {
        Write-Host "Variable '$((Get-Variable -Name "$($licentie)_O365").Name)' have a value of $((Get-Variable -Name "$($licentie)_O365").Value)"
        if ($((Get-Variable -Name "$($licentie)_O365").Name) -in (run-MySQLQuery -ConnectionString $ConnectionString -Query "SHOW columns FROM user_counter").field) {
            Write-Host "$((Get-Variable -Name "$($licentie)_O365").Name) license is in the database"
        } else {
            Write-Host "$((Get-Variable -Name "$($licentie)_O365").Name) license is not in the database"
            try {
                Run-MySQLQuery -connectionString $ConnectionString -query "ALTER TABLE `mwt_test`.`user_counter` ADD COLUMN $((Get-Variable -Name "$($licentie)_O365").Name) INT(10) NULL DEFAULT '0' AFTER `ExcludedUsers`;"
            } catch {
                Write-Host "Failed adding Write-Host $((Get-Variable -Name "$($licentie)_O365").Name) to the database."
            } finally {
                Write-Host "We've tried adding $((Get-Variable -Name "$($licentie)_O365").Name) to the database."
            }
        }
        run-MySQLQuery -ConnectionString $ConnectionString -Query "INSERT INTO user_counter (Date, Month, Company, ADUsers, OfficeUsers, RDSUsers, MailboxUsers, ExcludedUsers) VALUES ('$now_unixtime', '$currentMonth','$Company','000','000','000','000', '000');"
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
O365Count