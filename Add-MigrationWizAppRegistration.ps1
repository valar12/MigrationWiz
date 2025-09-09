# Step 1: Install and Import Microsoft Graph Modules
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    $installModuleParams = @{
        Name         = "Microsoft.Graph"
        Scope        = "CurrentUser"
        Force        = $false
        AllowClobber = $true
    }
    Install-Module @installModuleParams
}

Import-Module Microsoft.Graph.Applications
Import-Module Microsoft.Graph.Identity.SignIns

# Step 2: Connect to Microsoft Graph
$connectMgGraphParams = @{
    Scopes = @(
        "Application.ReadWrite.All",
        "Directory.ReadWrite.All",
        "AppRoleAssignment.ReadWrite.All",
        "DelegatedPermissionGrant.ReadWrite.All"
    )
    NoWelcome = $true
}

try {
    Write-Host "Connecting to Microsoft Graph." -ForegroundColor DarkGray
    Connect-MgGraph @connectMgGraphParams
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# Step 3: Check if Application Display Name Already Exists
$appDisplayName = "MigrationWiz"
$existingApp = Get-MgApplication -Filter "displayName eq '$appDisplayName'"

if ($existingApp) {
    Write-Error "An application with the display name '$appDisplayName' already exists. Exiting script."
    exit 1
}

# Step 4: Get the Service Principal for Exchange Online and Permission ID for EWS.AccessAsUser.All
$exchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'

# Retrieve the Exchange Online Service Principal
$exchangeServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$exchangeOnlineAppId'"

# Find Permission ID for EWS.AccessAsUser.All
$ewsPermissionId = ($exchangeServicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Value -eq "EWS.AccessAsUser.All" }).Id

# Retrieve the Service Principal for Office 365 Exchange Online and get the 'full_access_as_app' role
$exchangeOnlineSpDisplayName = 'Office 365 Exchange Online'
$fullAccessRole = (Get-MgServicePrincipal -Filter "DisplayName eq '$exchangeOnlineSpDisplayName'").AppRoles | Where-Object { $_.Value -eq 'full_access_as_app' }
$fullAccessId = $fullAccessRole.Id

# Step 5: Prepare RequiredResourceAccess (Consolidated Permissions)
$requiredResourceAccess = @(
    @{
        ResourceAppId   = $exchangeOnlineAppId
        ResourceAccess  = @(
            # EWS.AccessAsUser.All - Delegated Permission
            @{
                Id   = $ewsPermissionId
                Type = 'Scope'
            },
            # full_access_as_app - Application Permission
            @{
                Id   = $fullAccessId
                Type = 'Role'
            }
        )
    }
)

# Step 6: Create the Application
$newAppParams = @{
    DisplayName            = $appDisplayName
    RequiredResourceAccess = $requiredResourceAccess
}
$application = New-MgApplication @newAppParams

$appId = $application.AppId
$objectId = $application.Id
$directoryId = (Get-MgOrganization).Id

Write-Output "====================="
Write-Output "Application ID: $appId"
Write-Output "Object ID: $objectId"
Write-Output "Directory ID (Tenant ID): $directoryId"
Write-Output "====================="

# Step 7: Create Service Principal for the New Application if it doesn't exist
$sp = Get-MgServicePrincipal -Filter "AppId eq '$appId'"
if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $appId
}

# Step 8: Grant Admin Consent for All Permissions (Roles and Scopes)
$graphSpId = (Get-MgServicePrincipal -Filter "AppId eq '$exchangeOnlineAppId'").Id

# Assigning Scopes (Delegated Permissions)
New-MgOauth2PermissionGrant -ClientId $sp.Id -ConsentType "AllPrincipals" -ResourceId $graphSpId -Scope "EWS.AccessAsUser.All" | Out-Null

# Assigning Roles (Application Permission)
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -AppRoleId $fullAccessId -ResourceId $graphSpId | Out-Null

# Step 9: Add a New Password (Client Secret) to the Application
$isoTimestamp  = (Get-Date).ToString("yyyy-MM-ddTHH-mm-ss")
$csDisplayName = "$appDisplayName-$isoTimestamp"
$pwdParams = @{
    ApplicationId      = $application.Id
    PasswordCredential = @{
        StartDateTime = (Get-Date)
        EndDateTime   = (Get-Date).AddYears(1)   # Password expires in 1 year
        DisplayName   = $csDisplayName           # Customize as needed
    }
}

$clientSecret = Add-MgApplicationPassword @pwdParams
Write-Output "====================="
Write-Output "Generated Client Secret: $($clientSecret.SecretText)"
Write-Output "====================="

# Copy the secret value to the Windows clipboard
$clientSecret.SecretText | Set-Clipboard
# Inform the user
Write-Host "`nThe secret value is copied to the clipboard and is valid for one year." -ForegroundColor Green

# Step 10: Disconnect from Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
