<#
.SYNOPSIS
Manages license tokens for the company matching the database record.
#>

# Configuration
$licenseFilePath = "C:\Users\user\AppData\Local\Aronium"
$googleSheetId = "1wyBe0Wfb7L0EWovcFsE24Mypu8XOqE6HhdCim0BdV7c"
$googleSheetName = "DB_Details"
$googleApiKey = "AIzaSyDyNVtUV7gFlNiKZCLwgUZWUGtOOxxPUlI"
$today = Get-Date

# SQL Server Configuration - Hardcoded as requested
$serverName = "PATRICK\ARONIUM"
$databaseName = "aroniumdatabase"

# Function to get data from Google Sheet
function Get-GoogleSheetData {
    $url = "https://sheets.googleapis.com/v4/spreadsheets/$googleSheetId/values/$googleSheetName" + "?key=$googleApiKey"
    
    try {
        Write-Host "Accessing Google Sheet..."
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        
        if (-not $response.values) {
            throw "No data returned from Google Sheet"
        }
        
        $headers = $response.values[0]
        $dataRows = $response.values | Select-Object -Skip 1

        $result = @()
        foreach ($row in $dataRows) {
            $obj = @{}
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $obj[$headers[$i]] = $row[$i]
            }
            $result += $obj
        }
        return $result
    }
    catch {
        Write-Error "Failed to fetch Google Sheet data: $_"
        return $null
    }
}

# Function to get company name from database
function Get-DatabaseCompanyName {
    try {
        $query = "SELECT Name FROM dbo.Company"
        $result = Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query -TrustServerCertificate
        return $result.Name
    }
    catch {
        Write-Error "Failed to get company name from database: $_"
        return $null
    }   
}
        
function Get-DatabaseToken {
    try {
        $query = "SELECT Value FROM dbo.ApplicationProperty WHERE Name='Account.RefreshToken'"
        $result = Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query -TrustServerCertificate
        return $result.Value
    }
    catch {
        Write-Error "Failed to get token from database: $_"
        return $null
    }
}

# Function to update refresh token in database
function Update-DatabaseToken {
    param (
        [string]$newToken
    )
    try {
        $query = "UPDATE dbo.ApplicationProperty SET Value='$newToken' WHERE Name='Account.RefreshToken'"
        Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query -TrustServerCertificate
        return $true
    }
    catch {
        Write-Error "Failed to update token in database: $_"
        return $false
    }
}

# Improved date parsing function
function Parse-ExcelDate {
    param (
        [string]$dateString
    )
    
    if ([string]::IsNullOrWhiteSpace($dateString)) {
        Write-Error "Empty date string"
        return $null
    }
    
    try {
        $dateString = $dateString.Trim()
        $formats = @('MM/dd/yyyy', 'M/d/yyyy', 'dd/MM/yyyy', 'yyyy-MM-dd')
        
        foreach ($format in $formats) {
            try {
                $parsedDate = [datetime]::ParseExact($dateString, $format, $null)
                return $parsedDate
            }
            catch { continue }
        }
        
        if ([datetime]::TryParse($dateString, [ref]$parsedDate)) {
            return $parsedDate
        }
        
        Write-Error "Could not parse date string: '$dateString'"
        return $null
    }
    catch {
        Write-Error "Date parsing error: $_"
        return $null
    }
}

# Function to check internet connectivity
function Test-InternetConnection {
    try {
        return Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction Stop
    }
    catch {
        return $false
    }
}

# Main script execution
try {
    Write-Host "Starting license management process..."
    Write-Host "Current date: $($today.ToString('MM/dd/yyyy'))"
    Write-Host "Connecting to SQL Server: $serverName, Database: $databaseName"

    # Get data from Google Sheet
    $sheetData = Get-GoogleSheetData
    if (-not $sheetData) {
        throw "No data retrieved from Google Sheet"
    }

    # Get company name from database
    Write-Host "Querying database for company name..."
    $dbCompanyName = Get-DatabaseCompanyName
    if (-not $dbCompanyName) {
        throw "Could not retrieve company name from database"
    }

    Write-Host "Database company name: $dbCompanyName"

    # Find matching row in sheet data
    $matchingRow = $sheetData | Where-Object { $_.'Company Name' -eq $dbCompanyName }
    if (-not $matchingRow) {
        throw "No matching company found in Google Sheet for '$dbCompanyName'"
    }

    Write-Host "`nProcessing matching company: $($matchingRow.'Company Name') (ID: $($matchingRow.ID))"
    
    # Parse expiration date
    $expirationDate = Parse-ExcelDate -dateString $matchingRow.'Expiration Date'
    if (-not $expirationDate) {
        throw "Invalid expiration date format"
    }

    # Parse boolean paid status
    $paid = $false
    if (-not [string]::IsNullOrWhiteSpace($matchingRow.Paid)) {
        $paid = [bool]::Parse($matchingRow.Paid)
    }

    Write-Host "Expiration: $($expirationDate.ToString('MM/dd/yyyy')) | Paid: $paid"
    
    # Check if license is still valid
    if ($expirationDate -gt $today) {
        Write-Host "License is expired - proceeding with token management"
        
        # Check internet connectivity
        $internetReachable = Test-InternetConnection
        
        if ($internetReachable) {
            Write-Host "Internet connection available"
            
            if ($paid) {
                Write-Host "Account is paid - checking token status"
                
                # Get current token from database
                $currentToken = Get-DatabaseToken
                if (-not $currentToken) {
                    throw "Could not retrieve current token"
                }

                $updatedToken = $matchingRow.'Updated Token'
                Write-Host "Current Token: $currentToken"
                Write-Host "Updated Token: $updatedToken"

                if ($currentToken -eq $updatedToken) {
                    Write-Host "Tokens match - no update needed"
                }
                else {
                    Write-Host "Tokens don't match - updating..."
                    
                    # Update database with token from Excel
                    if (Update-DatabaseToken -newToken $updatedToken) {
                        Write-Host "Database token updated successfully"
                    }
                }
            }
            else {
                Write-Host "Account is not paid - executing custom logic"
                
                # Get current token from database
                $currentToken = Get-DatabaseToken
                if (-not $currentToken) {
                    throw "Could not retrieve current token"
                }

                # Create new token by concatenating
                $newToken = $currentToken + "A_Updated"
                
                $licenseFile = Join-Path -Path $licenseFilePath -ChildPath "aronium.lic"
                if (Test-Path $licenseFile) {
                    Write-Host "Removing license file"
                    Remove-Item -Path $licenseFile -Force
                }
                
                # Update database with new token
                if (Update-DatabaseToken -newToken $newToken) {
                    Write-Host "Database token updated with suffix"
                    Write-Host "NOTE: Would update Google Sheet with new token if write permissions were configured"
                }
            }
        }
        else {
            Write-Host "No internet connection - managing license file"
            $licenseFile = Join-Path -Path $licenseFilePath -ChildPath "aronium.lic"
            if (Test-Path $licenseFile) {
                Write-Host "Removing license file"
                Remove-Item -Path $licenseFile -Force
            }
        }
    }
    else {
        Write-Host "License is valid - no action needed"
    }

    Write-Host "`nLicense management process completed"
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Host "Script execution finished"
}