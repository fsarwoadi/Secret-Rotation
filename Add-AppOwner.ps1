#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications

$ManagedIdentityObjectId = 'e06a3246-42d6-44ce-b071-f5a35b438563'
$TargetAppId             = 'ca0ce3c5-6d1a-4ec8-9f41-02f83a6d8cae'

# Requires a user who is Application Administrator, Cloud Application Administrator, or an existing owner of the app.
Connect-MgGraph -Scopes 'Application.ReadWrite.All' -NoWelcome

$app = Get-MgApplication -Filter "appId eq '$TargetAppId'"
if (-not $app) { throw "Application with appId '$TargetAppId' not found." }

$alreadyOwner = Get-MgApplicationOwner -ApplicationId $app.Id -All |
    Where-Object { $_.Id -eq $ManagedIdentityObjectId }

if ($alreadyOwner) {
    Write-Host "$ManagedIdentityObjectId is already an owner of '$($app.DisplayName)'."
}
else {
    New-MgApplicationOwnerByRef -ApplicationId $app.Id -BodyParameter @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ManagedIdentityObjectId"
    }
    Write-Host "Added $ManagedIdentityObjectId as an owner of '$($app.DisplayName)'."
}

Disconnect-MgGraph | Out-Null
