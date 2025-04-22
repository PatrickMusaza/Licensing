$serverName = "PATRICK\ARONIUM"
$databaseName = "aroniumdatabase"
$query = "SELECT Value FROM dbo.ApplicationProperty WHERE Name='Account.RefreshToken'"

Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query -TrustServerCertificate