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
    Write-Output "Error while loading the initialization script." 
    exit 10
}

##############################################################
#
# Do not change, unles you know what you are doing...

Function SendMail ($Company) {
    if ($currentMonth -eq 1) {
        $lastMonth = 12
        if ($debug) { 
            $lastmonth = 1
        }
    } else {
        $lastMonth = ($currentMonth - 1)
        if ($debug) { 
            $lastmonth = $lastmonth+1
        }
    }
    
    Write-Output "Setting up Mail..." | timestamp

    ## Begin mail
    $body = "<head><style> #customers { font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif; font-size:12px; border-collapse: collapse; width: 100%; } #customers td, #customers th { border: 1px solid #ddd; padding: 8px; } #customers tr:nth-child(even){background-color: #f2f2f2;} #customers tr:hover {background-color: #ddd;} #customers th { padding-top: 12px; padding-bottom: 12px; text-align: left; background-color: #142167; color: white; }</style></head>"

    $rows = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * from user_counter where company='$Company' AND Month=$lastMonth ORDER By Date ASC;"

    
    $licenties = @()
    $grablicenties = run-MySQLQuery -ConnectionString $ConnectionString -Query "SHOW columns FROM user_counter LIKE '%_O365'"
    foreach ($grabbedlicentie in $grablicenties) {
        ## setting up variables for highest and lowest counter.
        $licentie = $grabbedlicentie.Field
        if ($licentie -ne "RIGHTSMANAGEMENT_ADHOC_O365") {
            $licenties += $grabbedlicentie.Field
            $rowchecker = (run-MySQLQuery -ConnectionString $ConnectionString -Query "SELECT $licentie FROM user_counter WHERE company='$Company' AND Month='$lastMonth' ORDER By Date ASC;")
            foreach ($rowcheck in $rowchecker) { 
                [int]$countcheck = [int]$countcheck + [int]$rowcheck[0]
            }
            New-Variable -Name "$($licentie)_check" -Value $countcheck -Force
            New-Variable -Name "$($licentie)_High" -Value 0 -Force
            New-Variable -Name "$($licentie)_Low" -Value 1000 -Force
        }
    }
    ## Zeroing some variables used to calculate highest numbers
    $ADUsersHigh = 0
    $OfficeUsersHigh = 0
    $RDSUsersHigh = 0
    $MailboxUsersHigh = 0
    $EXCLUDEDHigh = 0
    
    ## thousanding some variables used to calculate lowest numbers
    $ADUsersLow = 1000
    $OfficeUsersLow = 1000
    $RDSUsersLow = 1000
    $MailboxUsersLow = 1000
    $EXCLUDEDLow = 1000
    foreach ($row in $rows) {
        ## Rowcounter
        $Rowcounter++
        
        ## Check user amount per row.
        $Date = Get-Date (Get-NormalDate -UnixDate $row[1]) -UFormat "%d-%m-%Y"
        $ADUsers = [int]$row[4]
        $OfficeUsers = [int]$row[5]
        $RDSUsers = [int]$row[6]
        $MailboxUsers = [int]$row[7]
        $EXCLUDED = [int]$row[8]

        ## Getting Lowest numbers
        if ($ADUsers -lt $ADUsersLow) {
            $ADUsersLow = $ADUsers
        }
        if ($OfficeUsers -lt $OfficeUsersLow) {
            $OfficeUsersLow = $OfficeUsers
        }
        if ($RDSUsers -lt $RDSUsersLow) {
            $RDSUsersLow = $RDSUsers
        }
        if ($MailboxUsers -lt $MailboxUsersLow) {
            $MailboxUsersLow = $MailboxUsers
        }
        if ($EXCLUDED -lt $EXCLUDEDLow) {
            $EXCLUDEDLow = $EXCLUDED
        }
        ## Getting Highest numbers
        if ($ADUsers -gt $ADUsersHigh) {
            $ADUsersHigh = $ADUsers
        }
        if ($OfficeUsers -gt $OfficeUsersHigh) {
            $OfficeUsersHigh = $OfficeUsers
        }
        if ($RDSUsers -gt $RDSUsersHigh) {
            $RDSUsersHigh = $RDSUsers
        }
        if ($MailboxUsers -gt $MailboxUsersHigh) {
            $MailboxUsersHigh = $MailboxUsers
        }
        if ($EXCLUDED -gt $EXCLUDEDHigh) {
            $EXCLUDEDHigh = $EXCLUDED
        }
    }

    ## Start table
    $body += "<center><img src='https://ictteamwork.nl/assets/uploads/2017/03/logo_icttw2.png'></center><br><br>"
    $body += "<center><font face='tahoma' color='#003399' size='4'><strong>Licentie rapportage $Company - Maand Nr: $lastMonth</strong></font></center><br><br>"
    $body += "<table id='customers'>"
    
    ## Set first row.
    $body += "<tr>"
    $body += "<th>Datum</th>"
    $body += "<th>ActiveDirectory<br>Gebruikers</th>"
    if($OfficeUsersHigh -gt 0) { $body += "<th>MS-Office<br>Gebruikers</th>" }
    if($RDSUsersHigh -gt 0) { $body += "<th>Remote Desktop<br>Gebruikers</th>" }
    if($MailboxUsersHigh -gt 0) { $body += "<th>Mailbox<br>Gebruikers</th>" }
    
    ## Make all columns for the titlebar in the table with all available licenses
    foreach ($licentie in $licenties) {
        if($((Get-Variable -Name "$($licentie)_check").Value) -gt 0) {
            $licentiename = $licentie.TrimEnd('_O365')
            $licentiename = $licentiename.replace('_','')
            $licentiename = $licentiename.TrimStart('O365')
            $body += "<th>O365 $licentiename<br>Gebruikers</th>"
        }
    }
    $body += "<th>Uitgezonderd<br>van telling</th>"
    $body += "</tr>"


    ## Get all usercount information
    $data = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * from user_counter where company='$Company' AND Month=$lastMonth ORDER By Date ASC;"
    foreach($dataitem in $data) {
        ## Rowcounter
        $Rowcounter++
        
        ## Check user amount per row.
        $Date = Get-Date (Get-NormalDate -UnixDate $dataitem[1]) -UFormat "%d-%m-%Y"
        $rowID = $dataitem[0]
        $functionaldate = $dataitem[1]
        $ADUsers = [int]$dataitem[4]
        $OfficeUsers = [int]$dataitem[5]
        $RDSUsers = [int]$dataitem[6]
        $MailboxUsers = [int]$dataitem[7]
        $EXCLUDED = [int]$dataitem[12]

        ## Make table rows for mail
        $body += "<tr>"
        $body += "<td>$Date</td>"
        $body += "<td>$ADUsers</td>"
        if($OfficeUsers -gt 0) { $body += "<td>$OfficeUsers</td>" }
        if($RDSUsers -gt 0) { $body += "<td>$RDSUsers</td>" }
        if($MailboxUsers -gt 0) { $body += "<td>$MailboxUsers</td>" }

        $avaragerowcounter++
        foreach ($licentie in $licenties) {
            $endtotal = 0
            if($((Get-Variable -Name "$($licentie)_check").Value) -gt 0) {
                $O365LicentieCount = (run-MySQLQuery -ConnectionString $ConnectionString -Query "SELECT $licentie FROM user_counter WHERE iduser_counter='$rowID'")

                foreach ($licensecount in $O365LicentieCount) {
                    New-Variable -Name "$($licentie)" -Value $licensecount[0] -Force
                    $body += "<td>$((Get-Variable -Name "$($licentie)").Value)</td>"
    
                    if ($((Get-Variable -Name "$($licentie)").Value) -gt $((Get-Variable -Name "$($licentie)_High").Value)) {
                        Set-Variable -Name "$($licentie)_High" -Value "$((Get-Variable -Name "$($licentie)").Value)" -Force
                    }
                    if ($((Get-Variable -Name "$($licentie)").Value) -lt $((Get-Variable -Name "$($licentie)_Low").Value)) {
                        Set-Variable -Name "$($licentie)_Low" -Value "$((Get-Variable -Name "$($licentie)").Value)" -Force
                    }
                }
            }
        }
        $body += "<td>$EXCLUDED</td>"
        $body += "</tr>"
    }

    ## Lowest totals.
    $body += "<tr><th colspan='1'>Laagste deze maand:</th><th>$ADUsersLow</th>"
    if($OfficeUsersHigh -gt 0) { $body += "<th> $OfficeUsersLow</th>" }
    if($RDSUsersHigh -gt 0) { $body += "<th>$RDSUsersLow</th>" }
    if($MailboxUsersHigh -gt 0) { $body += "<th>$MailboxUsersLow</th>" }
    foreach ($licentie in $licenties) {
        if($((Get-Variable -Name "$($licentie)_check").Value) -gt 0) {
            $Lowest = $((Get-Variable -Name "$($licentie)_Low").Value)
            $body += "<th>$Lowest</th>"
        }
    }
    $body += "<th>$EXCLUDEDLow</th></tr>"

    ## Highest total
    $body += "<tr><th colspan='1'>Hoogste deze maand:</th><th>$ADUsersHigh</th>"
    if($OfficeUsersHigh -gt 0) { $body += "<th> $OfficeUsersHigh</th>" }
    if($RDSUsersHigh -gt 0) { $body += "<th>$RDSUsersHigh</th>" }
    if($MailboxUsersHigh -gt 0) { $body += "<th>$MailboxUsersHigh</th>" }
    foreach ($licentie in $licenties) {
        if($((Get-Variable -Name "$($licentie)_check").Value) -gt 0) {
            $Highest = $((Get-Variable -Name "$($licentie)_High").Value)
            $body += "<th>$Highest</th>"
        }
    }
    $body += "<th>$EXCLUDEDHigh</th></tr>"


    ## End Table
    $body +="</table>"
    $body +="<br><br>"

    ## Start users table
    $body += "<table id='customers'>"
    $body += "<tr><td colspan='8' height='25' align='center'><font face='tahoma' color='#003399' size='4'><strong>Gebruikerslijst $Company</strong></font></td></tr>"

    ## Set first table row
    $body += "<tr>"
    $body += "<th>Nr.</th>"
    $body += "<th>Naam<br>Medewerker</th>"
    $body += "<th>Login<br>(UPN)</th>"
    $body += "<th>Afdeling/<br>Vestiging</th>"
    $body += "<th>Office365<br>Licentie</th>"
    $body += "<th>Account<br>Aangemaakt</th>"
    $body += "<th>Laatst<br>Ingelogd</th>"
    $body += "<th>Verwijderd</th>"
    $body += "</tr>"

    ## Get all users
    $userlist = run-MySQLQuery -ConnectionString $ConnectionString -Query "Select * FROM users WHERE Company='$Company' Order By deleted ASC, Created DESC;"
    foreach($userrow in $userlist) {
        ## User information from rows.
        $id = $userrow[0]
        $RowUserCounter++
        $AccountName = $userrow[1]
        $DisplayName = $userrow[2]
        $Department = $userrow[4]
        $UPN = $userrow[6]
        $Created = Get-Date (Get-NormalDate -UnixDate $userrow[7]) -UFormat "%d-%m-%Y"
        if (![string]::IsNullOrEmpty($userrow[8])) {
            $Removed = Get-Date (Get-NormalDate -UnixDate $userrow[8]) -UFormat "%d-%m-%Y"
        } else {
            $Removed = $userrow[8]
        }
        $O365lic = $userrow[9]
        if (![string]::IsNullOrEmpty($O365lic)) {
            $O365lic = $O365lic.replace('_',' ')
            $O365lic = $O365lic.TrimStart('O365')
        }
        if (![string]::IsNullOrEmpty($userrow[10])) {
            $LastLogin = Get-Date (Get-NormalDate -UnixDate $userrow[10]) -UFormat "%d-%m-%Y"
        } else {
            $LastLogin = "Never"
        }
        ## Make table rows for each user.
        $body += "<tr>"
        if (![string]::IsNullOrEmpty($removed)) { 
            $body += "<td style='color: red'>$RowUserCounter</td>"
            $body += "<td style='color: red'>$DisplayName</td>"
            $body += "<td style='color: red'>$UPN</td>"
            $body += "<td style='color: red'>$Department</td>"
            $body += "<td style='color: red'>$O365lic</td>"
            $body += "<td style='color: red'>$Created</td>"
            $body += "<td style='color: red'>$LastLogin</td>"
            $body += "<td style='color: red'>$Removed</td>" 
        } else { 
            $body += "<td>$RowUserCounter</td>"
            $body += "<td>$DisplayName</td>" 
            $body += "<td>$UPN</td>"
            $body += "<td>$Department</td>"
            $body += "<td>$O365lic</td>"
            $body += "<td>$Created</td>" 
            $body += "<td>$LastLogin</td>"
            $body += "<td>$Removed</td>" 
        }
        $body += "</tr>"
    }
    ## End Table
    $body +="</table>"


    #### Now send the email using \> Send-MailMessage  
    Try {
        $subject = "Licentie rapportage $company (Maand $lastMonth)"

        foreach($t in $to)
            {

                send-MailMessage -SmtpServer $smtp -To $t -From $from -Subject $subject -Body $body -BodyAsHtml -Priority Normal 
                Write-Output "Sending email to '$t' for company '$company'..." | timestamp
            
            }
        
    } Catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Output "Sending mail failed... $ErrorMessage : $FailedItem" | timestamp
    }

    ########### End of Script################ 
}


foreach($c in $companies)
{
        SendMail -Company $c
    
}