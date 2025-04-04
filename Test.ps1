$serverName = "PATRICK\ARONIUM"
$databaseName = "aroniumdatabase"
$query = "UPDATE dbo.ApplicationProperty SET Value = '9e9f149dfd4247e7a8ee76a7b28c42b6' WHERE Name = 'Account.RefreshToken';"

Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query
