###############################################
##   Configuration of the database storage.  ##
###############################################
## Switches 
[switch]$ExcludeUsers = $true
[switch]$ImportADUsers = $true
[switch]$CheckSQLUsers = $true
[switch]$O365 = $true
[switch]$timestamps = $true
[switch]$debug = $false


##################
# Variables

## Database Information
$ConnectionString = "$DatabaseHost;$DatabaseUser;$DatabasePwd;$DatabaseDB;"

# Company Name
$Company = "company"

# OUs and Groups where users are stored/found.
$BaseOU = "OU=company,DC=int,DC=contoso,DC=com"
$DomainGroups = @{
    "ADUsers"="OU=Users"
    "Office"="RG-APP-MSOffice"
    "RDS"="RG-RDSSH"
}

# Excluding accounts, group or OUs
$DefaultAccounts = "Administrator","DefaultAccount","krbtgt"
$ExcludeGroup = "UC-Exclude"
$ExcludeOUs = "OU=GPO Test,OU=company,OU=Users,OU=company,DC=int,DC=contoso,DC=com","OU=External Admins,OU=company,DC=int,DC=contoso,DC=com"

## Office 365 variables.
$O365Klanten = "$ScriptDirectory"
$O365Tenant = "O365"