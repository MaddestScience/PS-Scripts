###############################################
##   Configuration of the database storage.  ##
###############################################
## Switches for LicenseMailer
[switch]$debug = $false

##################
# Variables

## Database Information
$DatabaseHost = 'server=localhost'
$DatabaseUser = 'Uid=<Username>'
$DatabasePwd = 'Pwd=<Password>'
$DatabaseDB = 'database=<database>'
$ConnectionString = "$DatabaseHost;$DatabaseUser;$DatabasePwd;$DatabaseDB;"

# Company Name
$Companies =@("<company1>","<company2>","<company3")

## Mail Variables
$smtp = "<mailserver: smtp.example.com>" 
$to = @("Name <e@mail.com>")
$from = "Example LLC <noreply@example.com>" 
$maillogo = "<urltomaillogo>"
