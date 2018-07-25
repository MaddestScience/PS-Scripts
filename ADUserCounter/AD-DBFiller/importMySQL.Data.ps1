# Define the assembly we want to load - a random reference assembly from SDK 3.0
$PshName = "MySql.Data"
$Pshfile = "$ScriptDirectory\MySQL.Data.dll"

Try {
    $MySQLPsh = [System.Reflection.Assembly]::LoadWithPartialName("$PshName")
    $MySQLPshref = $MySQLPsh.GetName()
    $MySQLPshName = $MySQLPshref.Name
    $MySQLPshVer =  $MySQLPshref.Version
    Write-Output "1. Assembly: $MySQLPshName with version number: $MySQLPshVer Loaded... " | timestamp
    $FailedLWPN = $false
}
Catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Output "1. Assembly $PshName not loaded with: $ErrorMessage, $FailedItem" | timestamp
    $FailedLWPN = $true
}

Try {
    $MySQLPsh = [System.Reflection.Assembly]::LoadFile("$Pshfile")
    $MySQLPshref = $MySQLPsh.GetName()
    $MySQLPshName = $MySQLPshref.Name
    $MySQLPshVer =  $MySQLPshref.Version
    Write-Output "2. Assembly: $MySQLPshName with version number: $MySQLPshVer Loaded... " | timestamp
    $FailedLF = $false
}
Catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Output "2. Assembly $PshName not loaded with: $ErrorMessage, $FailedItem" | timestamp
    $FailedLF = $true
}
Finally {
    if ($FailedLWPN -eq $false -or $FailedLF -eq $false) {
        Write-Output "Succeded importing $PshName. Continuing script..." | timestamp
    } else { 
        Write-Output "Failed: $Failed, importing $PshName failed, script aborting." | timestamp
        exit
    }
}