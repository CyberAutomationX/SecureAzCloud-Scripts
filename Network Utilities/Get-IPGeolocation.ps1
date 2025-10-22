<#
.SYNOPSIS
    This script retrieves detailed geolocation information for a given IP address using an external API.

.DESCRIPTION
    The Get-IPGeolocation function sends a request to an IP geolocation API to gather detailed 
    information such as the status, city, country, region, zip code, latitude, longitude, ISP, organization, 
    and more associated with a specified IP address. The function returns this information in a structured PowerShell object.

.PARAMETER IPAddress
    The IP address for which geolocation information will be retrieved.

.EXAMPLE
    .\Get-IPGeolocation -IPAddress "8.8.8.8"
    This example retrieves geolocation data for the IP address 8.8.8.8.

.NOTES
    Author: Ankit Gupta
    Version: 1.1 - 25-Aug-2024
    GitHub Link: https://github.com/AnkitG365/SecureAzCloud-Scripts

    This script should be tested in a non-production environment before being used in production.

#>

function Get-IPGeolocation {
    Param (
        [string]$IPAddress
    )

    # Invoke the REST API to retrieve geolocation data for the specified IP address
    $GeoData = Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$IPAddress"

    # Return the geolocation information as a PowerShell custom object
    [PSCustomObject]@{
        Status      = $GeoData.Status
        Country     = $GeoData.Country
        CountryCode = $GeoData.CountryCode
        Region      = $GeoData.Region
        RegionName  = $GeoData.RegionName
        City        = $GeoData.City
        Zip         = $GeoData.Zip
        Latitude    = $GeoData.Lat
        Longitude   = $GeoData.Lon
        TimeZone    = $GeoData.Timezone
        ISP         = $GeoData.Isp
        Organization= $GeoData.Org
        AS          = $GeoData.As
        IPAddress   = $GeoData.Query
    } 
}

# Example usage
$IPDetails = Get-IPGeolocation -IPAddress "20.81.111.85"
$IPDetails | Format-List
