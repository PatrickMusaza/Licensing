<#
.SYNOPSIS
Checks license status and manages tokens based on Google Sheet data.

.DESCRIPTION
This script:
1. Connects to a Google Sheet to get license information
2. Compares expiration dates with current date
3. Manages refresh tokens based on payment status and internet connectivity
4. Updates files or runs scripts as needed
#>

# Import required modules
Import-Module SqlServer

# Google Sheets API setup
$script:googleSheetId = "1wyBe0Wfb7L0EWovcFsE24Mypu8XOqE6HhdCim0BdV7c"
$script:googleSheetName = "DB_Details" 
$script:googleApiKey = "AIzaSyDyNVtUV7gFlNiKZCLwgUZWUGtOOxxPUlI" 

# Function to get data from Google Sheet
function Get-GoogleSheetData {
    $url = "https://sheets.googleapis.com/v4/spreadsheets/$googleSheetId/values/$googleSheetName" + "?key=$googleApiKey"
    
    try {
        Write-Host "Attempting to access Google Sheet at: $url"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        
        if (-not $response.values) {
            throw "No data returned from Google Sheet - check if sheet is empty"
        }
        
        $headers = $response.values[0]
        $dataRows = $response.values | Select-Object -Skip 1

        Write-Host "Data retrieved successfully from Google Sheet: $dataRows"

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
        Write-Host "Detailed error information:"
        Write-Host "Status Code: $($_.Exception.Response.StatusCode.value__)"
        Write-Host "Status Description: $($_.Exception.Response.StatusDescription)"
        Write-Host "Response Content: $($_.ErrorDetails.Message)"
        throw "Failed to fetch Google Sheet data: $_"
    }
}

# Function to update the refresh token
function Update-RefreshToken {
    param (
        [string]$refreshToken,
        [string]$filePath
    )
    
    # Your concatenation logic here (modify as needed)
    $updatedToken = $refreshToken + "A_updated"
    
    # Update the file with the new token
    try {
        if ($filePath -like "*.lic") {
            # Update license file
            $content = Get-Content -Path $filePath -Raw
            $updatedContent = $content -replace 'Account\.RefreshToken=.*', "Account.RefreshToken=$updatedToken"
            Set-Content -Path $filePath -Value $updatedContent -Force
        }
        else {
            # For other file types, just update the content
            Set-Content -Path $filePath -Value "Account.RefreshToken=$updatedToken" -Force
        }
        
        Write-Host "Token updated in $filePath"
        return $updatedToken
    }
    catch {
        Write-Error "Failed to update token in file: $_"
        return $null
    }
}

# Main script logic
try {
    # Get current date
    $today = Get-Date

    Write-Host $today
    
    # Get data from Google Sheet
    $sheetData = Get-GoogleSheetData
    if (-not $sheetData) {
        throw "No data retrieved from Google Sheet"
    }

    # Get company name from SQL (assuming first row has server info)
    $serverName = $sheetData[0].'Server Name'
    $databaseName = $sheetData[0].'Database Name'
    $companyName = (Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query).Value
    $query = "SELECT Name FROM dbo.Company"

    # Find matching row in Google Sheet
    $matchingRow = $sheetData | Where-Object { $_."Company Name" -eq $companyName }
    
    if ($matchingRow) {
        $dateExpired = [datetime]::ParseExact($matchingRow.'Expiration Date', 'MM/dd/yyyy', $null)
        $paid = [bool]::Parse($matchingRow.Paid)
        $updatedToken = $matchingRow.'Updated Token'
        $filePath = $matchingRow.File
        
        if ($dateExpired -gt $today) {
            # Check internet connectivity
            $internetReachable = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
            
            if ($internetReachable) {
                if ($paid) {
                    # Get current refresh token from SQL
                    $query = "SELECT Value FROM dbo.ApplicationProperty WHERE Name='Account.RefreshToken'"
                    $refreshToken = (Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $query).Value
                    
                    if ($refreshToken -eq $updatedToken) {
                        Write-Host "Tokens match, continuing..."
                    }
                    else {
                        # Update token
                        $newToken = Update-RefreshToken -refreshToken $refreshToken -filePath $filePath
                        if ($newToken) {
                            # Update the token in SQL if needed
                            $updateQuery = "UPDATE dbo.ApplicationProperty SET Value='$newToken' WHERE Name='Account.RefreshToken'"
                            Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Query $updateQuery
                        }
                    }
                }
                else {
                    # Run the script in the file
                    if (Test-Path $filePath) {
                        Write-Host "Running script at $filePath"
                        & $filePath
                    }
                    else {
                        Write-Warning "Script file not found at $filePath"
                    }
                }
            }
            else {
                # Delete file
                if (Test-Path $filePath) {
                    Remove-Item -Path $filePath -Force
                    Write-Host "File deleted due to no internet connection"
                }
            }
        }
        else {
            Write-Host "License not expired, continuing..."
        }
    }
    else {
        Write-Host "No matching company found in Google Sheet, continuing..."
    }
}
catch {
    Write-Error "An error occurred: $_"
    # Additional error handling if needed
    # exit 1
}