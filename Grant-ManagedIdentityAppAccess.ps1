#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications

# Grants a managed identity Application.ReadWrite.OwnedBy on Microsoft Graph
# and makes it owner of the target app registration.

$ManagedIdentityObjectId = 'e06a3246-42d6-44ce-b071-f5a35b438563'
$TargetAppId             = 'eab5c596-7b16-48e0-9ccf-de34fc130ed9'
$PermissionName          = 'Application.ReadWrite.OwnedBy'
$GraphAppId              = '00000003-0000-0000-c000-000000000000'

# Needs rights to grant admin consent (Privileged Role Administrator / Global Administrator)
# and to modify app owners (Application Administrator / Cloud Application Administrator / Global Administrator).
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.ReadWrite.All' -NoWelcome

try {
    # 1. Grant the Graph application permission
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
    $appRole = $graphSp.AppRoles | Where-Object {
        $_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains 'Application'
    }
    if (-not $appRole) { throw "App role '$PermissionName' not found on Microsoft Graph." }

    $existingRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId -All |
        Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $graphSp.Id }

    if ($existingRole) {
        Write-Host "'$PermissionName' is already granted to $ManagedIdentityObjectId."
    }
    else {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId -BodyParameter @{
            principalId = $ManagedIdentityObjectId
            resourceId  = $graphSp.Id
            appRoleId   = $appRole.Id
        } | Out-Null
        Write-Host "Granted '$PermissionName' to $ManagedIdentityObjectId."
    }

    # 2. Make the managed identity an owner of the target application
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
}
finally {
    Disconnect-MgGraph | Out-Null
}
