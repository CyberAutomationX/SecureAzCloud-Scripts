<#
.SYNOPSIS
    This script exports all Conditional Access Policies from Microsoft Entra ID (formerly Azure AD) to JSON and Excel formats.

.DESCRIPTION
    The `Export-ConditionalAccessPolicies.ps1` script connects to Microsoft Graph and retrieves all Conditional Access Policies 
    configured in your organization's Entra ID. It exports these policies to both a JSON file and an Excel spreadsheet, 
    providing a structured format for review and analysis.

    The script performs the following tasks:
    - Connects to Microsoft Graph with the required scopes.
    - Retrieves all Conditional Access Policies.
    - Resolves GUIDs to human-readable names for users, groups, applications, roles, and locations.
    - Exports the policies to a JSON file.
    - Formats the policies and exports them to an Excel spreadsheet with a colored header.

    This script is useful for:
    - Documenting Conditional Access Policies.
    - Auditing and reviewing policy configurations for compliance and security.
    - Providing insights into policy assignments, conditions, and controls.

.PARAMETER None

.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1

    This example runs the script and exports all Conditional Access Policies to JSON and Excel files in the specified paths.

.NOTES
    Author: Ankit Gupta
    GitHub: https://github.com/AnkitG365/SecureAzCloud-Scripts
    LinkedIn: https://www.linkedin.com/in/ankytgupta/

    Version: 1.0
    Created: 12/08/2024 

.REQUIREMENTS
    - PowerShell 5.1 or later.
    - Microsoft.Graph PowerShell module.
    - ImportExcel PowerShell module (for exporting to Excel).
    - Microsoft Graph API permissions: 
        - `DeviceManagementManagedDevices.Read.All`
        - `ConditionalAccess.Read.All`

.LINK
    https://github.com/AnkitG365/SecureAzCloud-Scripts

.DISCLAIMER
    This script is provided "as-is" without warranty of any kind. Use it at your own risk.
#>

# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel
}

# Connect to Microsoft Graph directly
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "ConditionalAccess.Read.All" -NoWelcome

# Dummy Paths for export
$JsonFilePath = "C:\Scripts\CAP\ConditionalAccessPolicies.json"
$ExcelFilePath = "C:\Scripts\CAP\ConditionalAccessPolicies.xlsx"

# Check if the JSON file exists and delete it if it does
if (Test-Path -Path $JsonFilePath) {
    Remove-Item -Path $JsonFilePath -Force
    Write-Output "Existing JSON file has been deleted."
}

# Fetch all Conditional Access Policies
$ConditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy

# Convert the policies to JSON format
$JsonOutput = $ConditionalAccessPolicies | ConvertTo-Json -Depth 10

# Export the JSON output to a file
Set-Content -Path $JsonFilePath -Value $JsonOutput

Write-Output "Conditional Access Policies have been exported to $JsonFilePath"

# Load the JSON file
$ConditionalAccessPolicies = Get-Content -Path $JsonFilePath | ConvertFrom-Json

# Helper function to resolve user GUIDs to their display names
function Resolve-UserGuids {
    param (
        [string[]]$UserGuids
    )
    $UserNames = @()
    foreach ($Guid in $UserGuids) {
        if ($Guid -match "^[a-f0-9-]{36}$") {
            try {
                $User = Get-MgUser -UserId $Guid -ErrorAction Stop
                $UserNames += $User.DisplayName
            } catch {
                $UserNames += "Unknown User ($Guid)"
            }
        } else {
            $UserNames += $Guid
        }
    }
    return $UserNames -join ", "
}

# Helper function to resolve group GUIDs to their display names
function Resolve-GroupGuids {
    param (
        [string[]]$GroupGuids
    )
    $GroupNames = @()
    foreach ($Guid in $GroupGuids) {
        if ($Guid -match "^[a-f0-0]{36}$") {
            try {
                $Group = Get-MgGroup -GroupId $Guid -ErrorAction Stop
                $GroupNames += $Group.DisplayName
            } catch {
                $GroupNames += "Unknown Group ($Guid)"
            }
        } else {
            $GroupNames += $Guid
        }
    }
    return $GroupNames -join ", "
}

# Helper function to resolve application GUIDs to their display names
function Resolve-AppGuids {
    param (
        [string[]]$AppGuids
    )
    $AppNames = @()
    foreach ($Guid in $AppGuids) {
        if ($Guid -match "^[a-f0-0]{36}$") {
            try {
                $App = Get-MgServicePrincipal -Filter "appId eq '$Guid'" -ErrorAction Stop
                $AppNames += $App.DisplayName
            } catch {
                $AppNames += "Unknown App ($Guid)"
            }
        } else {
            $AppNames += $Guid
        }
    }
    return $AppNames -join ", "
}

# Helper function to resolve role GUIDs to their display names
function Resolve-RoleGuids {
    param (
        [string[]]$RoleGuids
    )
    $RoleNames = @()
    foreach ($Guid in $RoleGuids) {
        if ($Guid -match "^[a-f0-0]{36}$") {
            try {
                $Role = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $Guid -ErrorAction Stop
                $RoleNames += $Role.DisplayName
            } catch {
                $RoleNames += "Unknown Role ($Guid)"
            }
        } else {
            $RoleNames += $Guid
        }
    }
    return $RoleNames -join ", "
}

# Helper function to resolve location GUIDs to their display names
function Resolve-LocationGuids {
    param (
        [string[]]$LocationGuids
    )
    $LocationNames = @()
    foreach ($Guid in $LocationGuids) {
        if ($Guid -match "^[a-f0-0]{36}$") {
            try {
                $Location = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $Guid -ErrorAction Stop
                $LocationNames += $Location.DisplayName
            } catch {
                $LocationNames += "Unknown Location ($Guid)"
            }
        } else {
            $LocationNames += $Guid
        }
    }
    return $LocationNames -join ", "
}

# Prepare a list to hold the formatted policy details
$PolicyDetailsList = @()

# Format each policy into a custom object with the desired columns
foreach ($Policy in $ConditionalAccessPolicies) {
    $IncludeUsers = if ($Policy.Conditions.Users.IncludeUsers) { Resolve-UserGuids -UserGuids $Policy.Conditions.Users.IncludeUsers } else { "" }
    $IncludeGroups = if ($Policy.Conditions.Users.IncludeGroups) { Resolve-GroupGuids -GroupGuids $Policy.Conditions.Users.IncludeGroups } else { "" }
    $ExcludeUsers = if ($Policy.Conditions.Users.ExcludeUsers) { Resolve-UserGuids -UserGuids $Policy.Conditions.Users.ExcludeUsers } else { "" }
    $ExcludeGroups = if ($Policy.Conditions.Users.ExcludeGroups) { Resolve-GroupGuids -GroupGuids $Policy.Conditions.Users.ExcludeGroups } else { "" }
    $IncludeApps = if ($Policy.Conditions.Applications.IncludeApplications) { Resolve-AppGuids -AppGuids $Policy.Conditions.Applications.IncludeApplications } else { "" }
    $ExcludeApps = if ($Policy.Conditions.Applications.ExcludeApplications) { Resolve-AppGuids -AppGuids $Policy.Conditions.Applications.ExcludeApplications } else { "" }
    $IncludeRoles = if ($Policy.Conditions.Users.IncludeRoles) { Resolve-RoleGuids -RoleGuids $Policy.Conditions.Users.IncludeRoles } else { "" }
    $ExcludeRoles = if ($Policy.Conditions.Users.ExcludeRoles) { Resolve-RoleGuids -RoleGuids $Policy.Conditions.Users.ExcludeRoles } else { "" }
    
    # Resolve Locations
    $IncludeLocations = if ($Policy.Conditions.Locations.IncludeLocations) { Resolve-LocationGuids -LocationGuids $Policy.Conditions.Locations.IncludeLocations } else { "" }
    $ExcludeLocations = if ($Policy.Conditions.Locations.ExcludeLocations) { Resolve-LocationGuids -LocationGuids $Policy.Conditions.Locations.ExcludeLocations } else { "" }

    # Handle ClientApps
    $ClientApps = if ($Policy.Conditions.ClientAppTypes -ne $null -and $Policy.Conditions.ClientAppTypes.Count -gt 0) {
        ($Policy.Conditions.ClientAppTypes -join ", ")
    } else {
        "all"
    }

    # Handle Devices
    $Devices = if (($Policy.Conditions.Devices.DeviceFilter.Mode -ne $null -and $Policy.Conditions.Devices.DeviceFilter.Mode -ne "") -or ($Policy.Conditions.Devices.DeviceFilter.Rule -ne $null -and $Policy.Conditions.Devices.DeviceFilter.Rule -ne "")) {
        "Mode: $($Policy.Conditions.Devices.DeviceFilter.Mode), Rule: $($Policy.Conditions.Devices.DeviceFilter.Rule)" -replace ", $", ""
    } else {
        ""
    }

    # Handle SessionControls
    $SessionControlsArray = @()
    if ($Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
        $SessionControlsArray += "Application Enforced Restrictions"
    }
    if ($Policy.SessionControls.CloudAppSecurity.IsEnabled) {
        $SessionControlsArray += "Cloud App Security"
    }
    if ($Policy.SessionControls.DisableResilienceDefaults) {
        $SessionControlsArray += "Disable Resilience Defaults"
    }
    if ($Policy.SessionControls.PersistentBrowser.IsEnabled) {
        $SessionControlsArray += "Persistent Browser"
    }
    if ($Policy.SessionControls.SignInFrequency.IsEnabled) {
        $SessionControlsArray += "Sign-In Frequency: $($Policy.SessionControls.SignInFrequency.Value) $($Policy.SessionControls.SignInFrequency.Period)"
    }
    $SessionControls = if ($SessionControlsArray.Count -gt 0) {
        $SessionControlsArray -join ", "
    } else {
        ""
    }

    $PolicyDetails = [PSCustomObject]@{
        PolicyID = $Policy.Id
        PolicyName = $Policy.DisplayName
        State = $Policy.State
        IncludeUsers = $IncludeUsers
        IncludeGroups = $IncludeGroups
        ExcludedUsers = $ExcludeUsers
        ExcludeGroups = $ExcludeGroups
        IncludeRoles = $IncludeRoles
        ExcludeRoles = $ExcludeRoles
        IncludeApplications = $IncludeApps
        ExcludeApplications = $ExcludeApps
        IncludeLocations = $IncludeLocations
        ExcludeLocations = $ExcludeLocations
        IncludePlatforms = ($Policy.Conditions.Platforms.IncludePlatforms -join ", ")
        ExcludePlatforms = ($Policy.Conditions.Platforms.ExcludePlatforms -join ", ")
        Devices = $Devices
        ClientApps = $ClientApps
        UsingLegacyClients = ($Policy.Conditions.ClientAppTypes -contains "exchangeActiveSync") -or ($Policy.Conditions.ClientAppTypes -contains "other")
        GrantControls = ($Policy.GrantControls.BuiltInControls -join ", ")
        RequireMFA = ($Policy.GrantControls.CustomAuthenticationFactors -join ", ")
        ForMultipleControls = if ($Policy.GrantControls.Operator -eq "AND") { "Require all the selected controls" } elseif ($Policy.GrantControls.Operator -eq "OR") { "Require one of the selected controls" } else { "" }
        SessionControls = $SessionControls
        Duplicate = $false
    }
    $PolicyDetailsList += $PolicyDetails
}

# Export the details to an Excel file with a colored header
$PolicyDetailsList | Export-Excel -Path $ExcelFilePath -AutoSize -BoldTopRow -WorkSheetname "Policies" 

Write-Output "Conditional Access Policies have been exported to $ExcelFilePath"
