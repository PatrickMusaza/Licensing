$serverName = "PATRICK\ARONIUM"
$databaseName = "aroniumdatabase"
$query = "SELECT Name FROM [dbo].[Company]"

Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query -TrustServerCertificate