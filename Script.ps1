<#
    .SYNOPSIS
    Manages license tokens based on Google Sheet data and database information.
    #>

# Configuration
$licenseFilePath = "C:\Users\user\AppData\Local\Aronium"
$googleSheetId = "1wyBe0Wfb7L0EWovcFsE24Mypu8XOqE6HhdCim0BdV7c"
$googleSheetName = "DB_Details" 
$googleApiKey = "AIzaSyDyNVtUV7gFlNiKZCLwgUZWUGtOOxxPUlI"
$today = Get-Date

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
        # Remove any whitespace
        $dateString = $dateString.Trim()
            
        # Try parsing with different formats
        $formats = @('MM/dd/yyyy', 'M/d/yyyy', 'dd/MM/yyyy', 'yyyy-MM-dd')
            
        foreach ($format in $formats) {
            try {
                $parsedDate = [datetime]::ParseExact($dateString, $format, $null)
                return $parsedDate
            }
            catch {
                continue
            }
        }
            
        # Try culture-invariant parsing
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
        $connection = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction Stop
        return $connection
    }
    catch {
        return $false
    }
}

# Main script execution
try {
    Write-Host "Starting license management process..."
    Write-Host "Current date: $($today.ToString('MM/dd/yyyy'))"

    # Get data from Google Sheet
    $sheetData = Get-GoogleSheetData
    if (-not $sheetData) {
        throw "No data retrieved from Google Sheet"
    }

    # Process each row in the sheet data
    foreach ($row in $sheetData) {
        try {
            if (-not $row.'Company Name') {
                Write-Host "Skipping row with empty company name"
                continue
            }

            Write-Host "`nProcessing company: $($row.'Company Name')"
                
            # Parse expiration date
            $expirationDate = Parse-ExcelDate -dateString $row.'Expiration Date'
            if (-not $expirationDate) {
                Write-Host "Invalid expiration date format, skipping..."
                continue
            }

            # Parse boolean paid status
            $paid = $false
            if (-not [string]::IsNullOrWhiteSpace($row.Paid)) {
                $paid = [bool]::Parse($row.Paid)
            }

            $companyName = $row.'Company Name'
            $serverName = $row.'Server Name'
            $databaseName = $row.'Database Name'

            Write-Host "Expiration: $($expirationDate.ToString('MM/dd/yyyy')) | Paid: $paid"
                
            # Check if license is still valid
            if ($expirationDate -gt $today) {
                Write-Host "License is valid"
                    
                # Check internet connectivity
                $internetReachable = Test-InternetConnection
                    
                if ($internetReachable) {
                    Write-Host "Internet connection available"
                        
                    if ($paid) {
                        Write-Host "Account is paid - checking token status"
                        # Add your token update logic here
                    }
                    else {
                        Write-Host "Account is not paid - executing custom logic"
                        # Add your custom logic for unpaid accounts here
                    }
                }
                else {
                    Write-Host "No internet connection - managing license file"
                    $licenseFile = Join-Path -Path $licenseFilePath -ChildPath "license.lic"
                    if (Test-Path $licenseFile) {
                        Write-Host "Removing license file"
                        Remove-Item -Path $licenseFile -Force
                    }
                }
            }
            else {
                Write-Host "License has expired"
            }
        }
        catch {
            Write-Error "Error processing company $($row.'Company Name'): $_"
            continue
        }
    }

    Write-Host "`nLicense management process completed"
}
catch {
    Write-Error "An error occurred in the main script execution: $_"
}