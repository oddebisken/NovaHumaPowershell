using namespace System.Net

param($Request, $TriggerMetadata)

$Tenant = $env:Tenant

Write-Host "PowerShell HTTP trigger function processed a request."

# Parse input parameters from query or body
$userPrincipalName = $Request.Query.UserPrincipalName
$attributeNumber = $Request.Query.CustomAttributeNumber
$attributeValue = $Request.Query.Value

if (-not $userPrincipalName) { $userPrincipalName = $Request.Body.UserPrincipalName }
if (-not $attributeNumber) { $attributeNumber = $Request.Body.CustomAttributeNumber }
if (-not $attributeValue) { $attributeValue = $Request.Body.Value }

# Validate input
if (-not $userPrincipalName -or -not $attributeNumber -or -not $attributeValue) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = "Missing required parameters: UserPrincipalName, CustomAttributeNumber, or Value."
    })
    return
}

# Construct the parameter name, e.g., CustomAttribute10
$customAttributeName = "CustomAttribute$attributeNumber"
import-module ExchangeOnlineManagement
try {
    # Connect to Exchange Online (make sure Managed Identity or credentials are configured)
    Connect-ExchangeOnline -ManagedIdentity -Organization $Tenant

    # Build the parameters dynamically
    $params = @{
        Identity = $userPrincipalName
    }
    $params[$customAttributeName] = $attributeValue

    # Execute Set-Mailbox with dynamic parameters
    Set-Mailbox @params

    $message = "Successfully updated $customAttributeName for $userPrincipalName."
    Write-Host $message

    # Output response
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $message
    })
}
catch {
    Write-Error "Error updating mailbox: $_"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = "Failed to update mailbox: $_"
    })
}

finally {
    Disconnect-ExchangeOnline -Confirm:$false
}