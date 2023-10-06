<#
.SYNOPSIS
This PowerShell script retrieves BSSIDs from Meraki access points and exports them to a CSV file. The BSSIDs are used for E911 location tracking on wireless access points in the Microsoft Teams LIS database.

.DESCRIPTION
This script uses the Meraki Dashboard API to retrieve BSSIDs for access points within an organization, and outputs them in a XLSX Excel file. The script requires the Meraki Dashboard API key and the organization ID. The script prompts the user to select an organization and enter a filter to search for specific access points.

The BSSIDs retrieved from the access points are formatted for use in the Microsoft Teams LIS database, which is used for E911 location tracking. The script outputs the BSSIDs to an Excel file, which can then be imported into the Teams LIS database.

.NOTES
- Organization ID and Meraki API Key are required to use this script.
- This script was developed for use with Microsoft Teams LIS database for E911 location tracking on wireless access points in conjunction with Meraki access points.
- This script requires the ImportExcel module to be installed. To install it, run "Install-Module -Name ImportExcel" in a PowerShell prompt.

#>

# Install the ImportExcel module if it is not already installed
if (-not(Get-Module -Name ImportExcel -ListAvailable)) {
    Write-Host "ImportExcel module not found. run ""Install-Module -Name ImportExcel"" " -ForegroundColor Red
}

# Define a function to get the headers required for making API calls
function Get-MerakiHeaders {
    # Get the API key from the MerakiBSSIDMapper class
    $APIKey = [MerakiBSSIDMapper]::APIKey()
    
    # Create a hashtable to hold the headers
    $headers = @{}
    $headers.Add("X-Cisco-Meraki-API-Key", $apiKey)  # Add the API key header
    $headers.Add("Content-Type", "application/json")  # Add the content type header
    $headers.Add("Accept", "*/*")  # Add the accept header
    
    return $headers  # Return the headers
}

# Set the base URL for the API
$BaseAPIURL = "https://api.meraki.com/api/v1"

# Define a function to get the networks for a given organization ID
function Get-MerakiNetworks {
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$OrgID
    )
    $EndpointURL = $BaseAPIURL + "/organizations/$OrgID/networks"  # Construct the API endpoint URL
    $headers = Get-MerakiHeaders  # Get the headers
    
    return (Invoke-RestMethod -Uri $EndpointURL -Method GET -Headers $headers -TimeoutSec 300)  # Invoke the API and return the response
}

# Define a function to get the organizations
function Get-MerakiOrganizations {  
    $EndpointURL = $BaseAPIURL + "/organizations"  # Construct the API endpoint URL
    $headers = Get-MerakiHeaders  # Get the headers
    
    return (Invoke-RestMethod -Uri $EndpointURL -Method GET -Headers $headers -TimeoutSec 300)  # Invoke the API and return the response
}

# Define a function to get the devices for a given organization ID and filter
function Get-MerakiOrganizationDevices {
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$OrgID,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)][string]$Filter
    )
    $params = @{"name" = "$Filter"; }  # Create a hashtable to hold the query parameters
    $EndpointURL = $BaseAPIURL + "/organizations/$OrgID/devices"  # Construct the API endpoint URL
    $headers = Get-MerakiHeaders  # Get the headers
    
    return (Invoke-RestMethod -Uri $EndpointURL -Method GET -Headers $headers -Body ($params | ConvertTo-Json) -TimeoutSec 300)  # Invoke the API and return the response
}

# Define a function to get the BSSID for a given serial number
function Get-MerakiAPBSSID {    
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$SN
    )
    $EndpointURL = $BaseAPIURL + "/devices/$SN/wireless/status"  # Construct the API endpoint URL
    $headers = Get-MerakiHeaders  # Get the headers
    
    return (Invoke-RestMethod -Uri $EndpointURL -Method GET -Headers $headers -TimeoutSec 300).basicServiceSets  # Invoke the API and return the response
}


    

# Set API Key - This is generated via the Meraki Portal
# $APIKey = ''


if (-not $APIKey) {
    $APIKey = Read-Host "Please enter your Meraki API key: "
}
else {
    Write-Host "Your Meraki API Key is set via the source code: $($APIKey)"
}

# Create API Key object
Add-Type @"
 using System;
 public class MerakiBSSIDMapper {
   public static string APIKey() {
     return `"$APIKey`";
   }
 }
"@

# Get Meraki organizations
$orgResults = Get-MerakiOrganizations

# Prompt user to select organization
Write-Host "Please Choose a Selection:"
$i = 0
$orgResults | % { Write-Host $([string]$i + ". " + $_.name); $i++ }
[int]$orgSelection = Read-Host $("Enter number of selection [0-" + ($i - 1) + "]")
$OrgID = $orgResults[$orgSelection].id
$Continue = $false

function LIST_APs($Filter) {
    Write-Host -ForegroundColor Yellow "Querying Meraki API."
    $FilterParam = if ($Filter) { @{"Filter" = $Filter } } else { @{} }
    $APList = Get-MerakiOrganizationDevices -OrgID $OrgID @FilterParam | Where-Object { $_.model -like "MR*" }

    $i = 1
    $table = @()
    $APList | ForEach-Object {
        $table += [pscustomobject]@{
            Item     = $i
            Name     = $_.Name
            Model    = $_.model
            MAC      = $_.mac
            Serial   = $_.serial
            Firmware = $_.firmware
        }
        $i++
    }
    $table | Format-Table -AutoSize
}







# Prompt user to enter filter
$Filter = Read-Host "Enter a filter for the Access Point name (e.g. 'AP-US' for APs starting with 'AP-US'), or leave it blank for all APs. Note: Meraki only returns up to 1000 devices and devices with no name are filtered."

# Call the LIST_APs function with the user's filter
LIST_APs -Filter $Filter

# Display message to indicate processing
Write-Host -ForegroundColor Yellow "Identifying BSSID data"

# Get list of Meraki devices for the selected organization and filter by MR* models
$APList = Get-MerakiOrganizationDevices -OrgID $OrgID -Filter $Filter | where { $_.model -like "MR*" } 

$Continue = $true

if ($Continue) {

    $i = 0
    # Loop through the list of access points and process each one
    $APTable = foreach ($AP in $APList) {

        # Display progress bar
        Write-Progress -Activity $("Processing AP: " + $AP.name) -PercentComplete $([Math]::Ceiling((($i / $APList.Count) * 100))); $i++

        # Get list of BSSIDs for the access point
        $BSSIDs = (Get-MerakiAPBSSID -SN $($AP.serial))

        # Loop through the BSSIDs and create a PSObject for each one
        $BSSIDs | ForEach-Object {
            $bssid = $_.bssid
            if ($bssid -like '*:*') {
                $bssid = $bssid -replace ':', '-'
            }
            New-Object -TypeName PSObject -Property ([ordered]@{
                    NAME             = $AP.name
                    MODEL            = $AP.model
                    BSSID            = $bssid
                    SSID             = $_.ssidName
                    BAND             = $_.band
                    CHANNEL          = $_.channel
                    TEAMSLOCATION_ID = 'ENTER LOCATION ID HERE'
                })
        }
    }


    # Get the path to the user's documents folder
    $documentsFolderPath = [Environment]::GetFolderPath("MyDocuments")

    # Set the filename
    $filename = "MerakiBSSIDs_.xlsx"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $timestampedFilename = "{0}_{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($filename), $timestamp, [System.IO.Path]::GetExtension($filename)
    $SavePath = Join-Path $documentsFolderPath $timestampedFilename

    $APTable | Export-Excel -Path $SavePath -Autosize

    # Export data to XLS file
    $APTable | Export-Excel -Path $SavePath -AutoSize -PassThru -WorksheetName 'APs' -TableName 'APTable' -ClearSheet `
        -ErrorAction SilentlyContinue -Verbose:$false | Out-Null

    # Load the Excel file into memory
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($SavePath)
    $worksheet = $workbook.Worksheets.Item(1)

    # Add Description column header
    $worksheet.Cells[1, 9].Value = "POWERSHELL"

    # Add CONCAT formula to each row in Description column
    $startRow = 2
    $totalRows = $worksheet.UsedRange.Rows.Count
    for ($row = $startRow; $row -le $totalRows + 1; $row++) {
        $worksheet.Cells[$row, 9].Formula = "=CONCATENATE(""set-CsOnlineLisWirelessAccessPoint -BSSID '"",C$row,""' -Description '"",A$row, "" "", B$row, "" "", D$row, "" "", E$row, ""' -LocationID '"",G$row,""'"")"
    }

    # Save and close the workbook
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()


}

Write-Host -ForegroundColor Yellow "Done! File Saved to $SavePath"
