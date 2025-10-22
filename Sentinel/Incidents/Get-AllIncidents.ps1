<#
    .SYNOPSIS
        This script is designed to retrieve all incidents from Azure Sentinel using the Azure Sentinel Incident API version 2020-01-01.

    .DESCRIPTION
        Security operations teams often require detailed reports of incidents from Azure Sentinel. This script gathers all relevant incidents from 
        a specified workspace and extracts key details, such as:
            - Unique Incident ID: A unique identifier for each incident, crucial for targeted retrieval.
            - Incident Title: The descriptive title of the incident.
            - Incident Number: A sequential identifier assigned upon creation of the incident.
            - Severity Level: Indicates the severity (Low, Medium, High, Critical) of the incident.
            - Incident Status: The current status (New, In Progress, Closed) of the incident.
            - Classification: Defines the classification (False Positive, True Positive) of the incident.
            - Classification Comments: Additional comments related to the classification.
            - Labels/Tags: Custom tags or labels associated with the incident.
            - Closure Reason: The reason behind closing the incident (False Positive, True Positive).
            - Assignee's Email: The email address of the individual responsible for the incident, retrieved from Azure AD.
            - Assigned To: Name of the person assigned to handle the incident.
            - Time Generated: The timestamp indicating when the incident was first generated.
            - Comment Count: The number of comments associated with the incident.

    .NOTES
        This script is built using the Azure PowerShell (Az) module.

        File Name     : Fetch-AzureSentinelIncidents.ps1
        Version       : 1.0.0.5
        Author        : Ankit Gupta
        Prerequisites : Azure PowerShell (Az) module
        Reference     : https://azsec.azurewebsites.net/2020/03/18/quick-look-at-new-azure-sentinel-incident-api/

    .EXAMPLE
        .\Fetch-AzureSentinelIncidents.ps1 -WorkspaceRg "security-rg" `
                                           -WorkspaceName "corp-workspace" `
                                           -FileName 'SentinelIncidents' `
                                           -Path 'C:\SecurityReports'
 
        This example fetches all incidents from the specified Azure Sentinel workspace and saves the report as a CSV file in the designated path.
#>

Param(
    [Parameter(Mandatory = $true,
        HelpMessage = "The name of the resource group containing the Log Analytics workspace connected to Azure Sentinel",
        Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]
    $WorkspaceRg,

    [Parameter(Mandatory = $true,
        HelpMessage = "The name of the Log Analytics workspace connected to Azure Sentinel",
        Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]
    $WorkspaceName,

    [Parameter(Mandatory = $true,
              HelpMessage = "The desired file name for the exported incident report",
              Position = 2)]
    [ValidateNotNullOrEmpty()]
    [string]
    $FileName,

    [Parameter(Mandatory = $true,
              HelpMessage = "The directory where the incident report will be saved",
              Position = 3)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Path
)

# Define the current timestamp for the report
$date = Get-Date -UFormat "%Y_%m_%d_%H%M%S"

# Retrieve the workspace ID using the provided resource group and workspace name
$workspaceId = (Get-AzOperationalInsightsWorkspace -Name $WorkspaceName `
                                                   -ResourceGroupName $WorkspaceRg).ResourceId
if (!$workspaceId) {
    throw "[!] The specified workspace could not be found. Please verify your input and try again."
}
else {
    Write-Host -ForegroundColor Green "[-] Successfully connected to Azure Sentinel workspace: $WorkspaceName"
}

# Class to store incident details for CSV export
class SentinelIncidentReport {
    [Object]$IncidentID
    [Object]$Title
    [Object]$IncidentNumber
    [Object]$Description
    [Object]$Severity
    [Object]$Status
    [Object]$Labels
    [Object]$Classification
    [Object]$ClassificationComments
    [Object]$Assignee
    [Object]$AssigneeEmail
    [Object]$CreatedTime
    [Object]$FirstActivityTime
    [Object]$LastActivityTime
    [Object]$LastModifiedTime
    [Object]$FirstAlertGenerated
    [Object]$LastAlertGenerated
    [Object]$AlertSource
    [Object]$CommentCount
    [Object]$AlertCount
}

# Obtain an Azure access token for the Resource Manager API
$accessToken = Get-AzAccessToken -ResourceTypeName "ResourceManager"
$authHeader = @{
    'Content-Type'  = 'application/json'
    'Authorization' = 'Bearer ' + $accessToken.Token
}

# Define the API endpoint for fetching incidents
$uri = "https://management.azure.com" + $workspaceId + "/providers/Microsoft.SecurityInsights/incidents/?api-version=2020-01-01"

$response = Invoke-RestMethod -Uri $uri `
                              -Method GET `
                              -Headers $authHeader
$incidents = $response.value

# Handle pagination if there are more results to fetch
while($response.nextLink)
{
    $nextLink = $response.nextLink
    $response = (Invoke-RestMethod -Uri $nextLink -Method GET -Headers $authHeader)
    $incidents += $response.value
}

# Initialize an array to store the incident details
$incidentReport = @()

# Loop through each incident and store the details in the class object
foreach ($incident in $incidents) {
    $incidentObj = [SentinelIncidentReport]::new()
    # Extract relevant incident details
    $incidentObj.IncidentID = $incident.id.Split('/')[12]
    $incidentObj.Title = $incident.properties.title
    $incidentObj.IncidentNumber = $incident.properties.incidentNumber
    $incidentObj.Description =  $incident.properties.description
    $incidentObj.Severity = $incident.properties.severity
    $incidentObj.Status = $incident.properties.status
    $incidentObj.Labels = $incident.properties.labels.labelName -join ", "
    $incidentObj.Classification = $incident.properties.classification
    $incidentObj.ClassificationComments = $incident.properties.classificationComment
    $incidentObj.Assignee = $incident.properties.owner.assignedTo
    $incidentObj.AssigneeEmail = $incident.properties.owner.email
    $incidentObj.CreatedTime = $incident.properties.createdTimeUtc
    $incidentObj.FirstActivityTime = $incident.properties.firstActivityTimeUTC
    $incidentObj.LastActivityTime = $incident.properties.lastActivityTimeUTC
    $incidentObj.LastModifiedTime = $incident.properties.lastModifiedTimeUTC
    $incidentObj.FirstAlertGenerated = $incident.properties.firstActivityTimeGenerated
    $incidentObj.LastAlertGenerated = $incident.properties.lastActivityTimeGenerated
    $incidentObj.AlertSource = $incident.properties.additionalData.alertProductNames -join ", "
    $incidentObj.CommentCount = $incident.properties.additionalData.commentsCount
    $incidentObj.AlertCount = $incident.properties.additionalData.alertsCount
    $incidentReport += $incidentObj
}

# Export the incident report to a CSV file
$incidentReport | Export-Csv -Path "$Path\$($FileName)_$($date).csv" -NoTypeInformation -Encoding UTF8
Write-Host -ForegroundColor Green "[-] Your Azure Sentinel incident report has been saved to: $Path\$($FileName)_$($date).csv"
