$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups", "ExchangeOnlineManagement", "Microsoft.Graph.Users.Actions")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..."
        Install-Module $module -Scope CurrentUser -Force
    } else {Write-Host "$module is already installed." -ForegroundColor Green}

    try {
        Import-Module $module -Force -ErrorAction Stop
        Write-Host "$module imported successfully.`n" -ForegroundColor Green
    } catch {Write-Host "Failed to import ${module}: ${_}`n" -ForegroundColor Red}
}

Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All", "Directory.Read.All"

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

# ValidateRBAC

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

# Prompt user to dertermine the correct DoneSafe group
# Loop until the user selects a valid number
while($true){
    try {
         # Validates the users input is an integer
         $doneSafe = [int](Read-Host "Please choose from one of the following:`n1: The user reports to the CEO.`n2: The user has direct reports.`n3: None of the above.")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($doneSafe -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($doneSafe -eq 1) {$groupNames += "DoneSafe Z Executives"}
             elseif($doneSafe -eq 2){$groupNames += "DoneSafe People Leaders"}
             elseif($doneSafe -eq 3){$groupNames += 'DoneSafe Leaders of Self'}
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

foreach ($groupName in $groupNames) {

    # Retrieve the group by name
    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    [string]$groupId = $group.id
    if ($group) {
        try {
            # Create the reference body to add the user
            $body = @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"}

            # Add the user to the group
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $body

            Write-Host "$($user.DisplayName) successfully added to '$($group.DisplayName)'" -ForegroundColor Green
        } catch {Write-Host "Error adding to '$groupName': $_" -ForegroundColor Red}
    } else {Write-Host "Group '$groupName' not found." -ForegroundColor Red}
}

$params = @{
    addLicenses = @(SkuId = 'b622badb-1b45-48d5-920f-4b27a2c0996c')
    removeLicenses = @()
}

#  Define  the  user and  license  SKU
$userPrincipalName = $user.UserPrincipalName

# Get  all  available  license SKUs
$skus = Get-MgSubscribedSku

#  Find  the SKU  that  includes  Viva Insights  (e.g.,  Viva  Suite or  M365  E5)
# Replace  this  with  the correct  SKU  part  number if  known
$vivaSku = $skus  |  Where-Object { $_.SkuPartNumber -like "*VIVA*" }

if ($vivaSku) {
    $skuId = $vivaSku.SkuId

    #  Prepare  license  assignment
    $addLicenses = @(
        @{
            SkuId = $skuId
        }
    )

    $removeLicenses = @()

    #  Assign the  license
    Set-MgUserLicense  -UserId $userPrincipalName  -AddLicenses  $addLicenses  -RemoveLicenses $removeLicenses

    Write-Host  "Viva Insights  license  assigned  to $userPrincipalName"
}
else {
    Write-Host  "Viva  Insights  SKU not  found.  Please  verify the  SKU  part  number."
}

try {
    Set-MgUserLicense -UserId $user.Id -AddLicenses $params
    Write-Host "Assigned Viva Insights license" -ForegroundColor Green
}catch{Write-Host $_}


try{
    # Set usage location to NZ
    Update-MgUser -UserId $user.Id -UsageLocation "NZ"
    Write-Host "Usage location has been set" -ForegroundColor Green
}catch{Write-Host $_}





