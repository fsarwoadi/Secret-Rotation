#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications

# Object (principal) ID of the Logic App's managed identity.
$ManagedIdentityObjectId = 'bbd945ce-f17b-49f2-bb61-28e825ec2a7f'
$PermissionName          = 'Application.ReadWrite.OwnedBy'
$GraphAppId              = '00000003-0000-0000-c000-000000000000'

# Requires a signed-in user who can grant admin consent (e.g. Privileged Role Administrator / Global Administrator).
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome

$graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
$appRole = $graphSp.AppRoles | Where-Object {
    $_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains 'Application'
}
if (-not $appRole) { throw "App role '$PermissionName' not found on Microsoft Graph." }

$existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId |
    Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $graphSp.Id }

if ($existing) {
    Write-Host "'$PermissionName' is already granted to $ManagedIdentityObjectId."
}
else {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId -BodyParameter @{
        principalId = $ManagedIdentityObjectId
        resourceId  = $graphSp.Id
        appRoleId   = $appRole.Id
    } | Format-List PrincipalDisplayName, AppRoleId, ResourceDisplayName, CreatedDateTime
    Write-Host "Granted '$PermissionName' to $ManagedIdentityObjectId."
}

Disconnect-MgGraph | Out-Null
