###############################################
##   Configuration of the database storage.  ##
###############################################

$DatabaseHost = "server=localhost"
$DatabaseUser = "Uid=root"
$DatabasePwd = "Pwd=password"
$DatabaseDB = "database=ActiveDirectory"
$DatabaseCert = "CertificateFile=C:\mysqlcert\mysqlclient.pfx"
$DatabaseCertPwd = "CertificatePassword=password"

# Connection string
$ConnectionString = "$DatabaseHost;$DatabaseUser;$DatabasePwd;$DatabaseDB;$DatabaseCert;$DatabaseCertPwd;"


## Configuration parameters
# Default administration / service accounts.
$DefaultAccounts = "Administrator","DefaultAccount","krbtgt"

# OUs and Groups where users are stored/found.
$BaseOU = "OU=Users,OU=TestBV,dc=testrealm,dc=local"
$DomainGroups = @{
    "ADUsers"="Domain Users"
    "Mailbox"="OU=Mailbox"
    "Office"="OU=Office"
    "RDS" = "RG-RDGW"
}

# Company Name
$Company = "TestBV"