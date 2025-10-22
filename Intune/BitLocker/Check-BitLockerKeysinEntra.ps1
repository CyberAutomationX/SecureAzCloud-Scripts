<#
.SYNOPSIS
    Check if BitLocker keys for Windows devices are stored in Entra ID.

.DESCRIPTION
    This script connects to Microsoft Graph API, retrieves all Windows devices from Intune,
    and checks if each device has a BitLocker key stored in Entra ID. The results are 
    displayed in a table and exported to a CSV file.

.EXAMPLE
    .\BitLocker-EntraID_CheckKeys.ps1

.NOTES
    Author: Ankit Gupta
    GitHub: https://github.com/AnkitG365/SecureAzCloud-Scripts/
    LinkedIn: https://www.linkedin.com/in/ankytgupta/

    Version: 1.0
    Updated: 12/08/2024
    Changes: Updated to use Microsoft Graph v1.0 API.

    Required Permissions:
    - DeviceManagementManagedDevices.Read.All
    - BitlockerKey.Read.All

    Disclaimer: This script is provided AS IS without warranty of any kind.
#>
  
# Function to get BitLocker key for a device
function Get-BitLockerKey {
    param (
        [string]$azureADDeviceId
    )

    $keyIdUri = "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys?`$filter=deviceId eq '$azureADDeviceId'"
    $keyIdResponse = Invoke-MgGraphRequest -Uri $keyIdUri -Method GET

    if ($keyIdResponse.value.Count -gt 0) {
        return "Yes"
    }
    return "No"
}

# Get all Windows devices from Intune (with pagination)
$devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'"
$devices = @()

do {
    $response = Invoke-MgGraphRequest -Uri $devicesUri -Method GET
    $devices += $response.value
    $devicesUri = $response.'@odata.nextLink'
} while ($devicesUri)

$results = @()

foreach ($device in $devices) {
    $hasBitlockerKey = Get-BitLockerKey -azureADDeviceId $device.azureADDeviceId

    $results += [PSCustomObject]@{
        DeviceName = $device.deviceName
        SerialNumber = $device.serialNumber
        "BitLocker Key in EntraID" = $hasBitlockerKey
        "Last Sync With Intune" = $device.lastSyncDateTime.ToString("yyyy-MM-dd")
    }
}

# Display results
$results | Format-Table -AutoSize

# Calculate summary statistics
$totalDevices = $results.Count
$devicesWithKey = ($results | Where-Object { $_.'BitLocker Key in EntraID' -eq 'Yes' }).Count
$devicesWithoutKey = $totalDevices - $devicesWithKey

# Display summary
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "Total Windows devices in Intune: $totalDevices" -ForegroundColor Yellow
Write-Host "Devices with BitLocker key stored in Entra ID: $devicesWithKey" -ForegroundColor Green
Write-Host "Devices without BitLocker key stored in Entra ID: $devicesWithoutKey" -ForegroundColor Red

$results | Export-Csv -Path "BitLockerKeyStatus.csv" -NoTypeInformation
