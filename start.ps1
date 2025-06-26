$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..."
        Install-Module $module -Scope CurrentUser -Force
    } else {
        Write-Host "$module is already installed." -ForegroundColor Green
    }

    try {
        Import-Module $module -Force -ErrorAction Stop
        Write-Host "$module imported successfully.`n" -ForegroundColor Green
    } catch {
        Write-Host "Failed to import ${module}: ${_}`n" -ForegroundColor Red
    }
}
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"

# Function to check if user exists
function ValidateAADUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )

    try {
        Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        Write-Host "User '$UserPrincipalName' exists in Azure AD." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "User '$UserPrincipalName' does NOT exist in Azure AD." -ForegroundColor Red
        return $false
    }
}

# Use regex to check the format of the UPN in valid
# Validate the user via the ValidateAADUser function
$pattern = '^[a-zA-Z0-9._%+,-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

$flag = $null
$user
do {
    $UserPrincipalName = Read-Host "Enter the user's UPN (Universal Principal Name)"
    if ($UserPrincipalName -notmatch $pattern) {
        Write-Host "Invalid format. Try again.`n"
        continue
    }
    try {
        $flag = ValidateAADUser $UserPrincipalName
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
    } catch {$flag = $null}
} until ($flag)

# Group names to assign
$groupNames = @(
    'sec-azure-zpa-all-users',
    'sec-azure-miro-users',
    'AutoPilot Users (Apps)',
    'sec-azure-SSPR-Enable',
    'Sec_G_CrossTenantSyncProd'
)


foreach ($groupName in $groupNames) {

    # Retrieve the group by name
    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    if ($group) {
        try {
            # Create the reference body to add the user
            $body = @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"}

            # Add the user to the group
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $body

            Write-Host "$($user.DisplayName) successfully added to '$($group.DisplayName)'" -ForegroundColor Green

        } catch {
            Write-Host "Error adding to '$groupName': $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Group '$groupName' not found." -ForegroundColor Red
    }
}

try{
    # Set usage location to NZ
    Update-MgUser -UserId $user.Id -UsageLocation "NZ"
    Write-Host "Usage location has been set" -ForegroundColor Green
}catch{Write-Host $_}





