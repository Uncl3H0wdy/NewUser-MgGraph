<#
    This script does the following:
    1. Promt for the target user's UPN and validates the input against a regex.
    2. Verify the UPN object exists in Entra before proceeding.
    3. Prompts for the location and managerial position and depending. 
    4. Adds user to Entra Security Groups.
    5. Adds user to user defined Distribution Groups.
    6. Sets the usage location to NZ.
    7. Assigns Viva Insights license.
    8. Configures users TrustedSendersAndDomains property in Exchange Online.
#>

Install-Module -Name Microsoft.Graph.Users -RequiredVersion 2.32.0 -Force
Install-Module -Name Microsoft.Graph.Groups -RequiredVersion 2.32.0 -Force
Install-Module -Name Microsoft.Graph.Users.Actions -RequiredVersion 2.32.0 -Force
Install-Module -Name ExchangeOnlineManagement -Force
Import-Module -Name Microsoft.Graph.Users
Import-Module -Name Microsoft.Graph.Groups
Import-Module -Name Microsoft.Graph.Users.Actions
Import-Module -Name ExchangeOnlineManagement

Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.ReadWrite.All", "Directory.Read.All" -NoWelcome

# Check if the user exists
function ValidateAADUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )

    # Check that the user exists. This function is called immediatley after the user inputs the UPN
    try {
        Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        return $true
    } catch {
        Write-Host "User '$UserPrincipalName' does NOT exist in Azure AD." -ForegroundColor Red
        return $false
    }
}

# Use regex to check the format of the UPN in valid
$pattern = '^[a-zA-Z0-9._%+,-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
$flag = $null
$user
$userLocation
$teamsNumber

# Validates the user's input against the regex defined in $pattern
do {
    $UserPrincipalName = Read-Host "Enter the user's UPN (Universal Principal Name)"
    if ($UserPrincipalName -notmatch $pattern) {
        Write-Host "Invalid format. Try again.`n"
        continue
    }
    try {
        # Function call which validates the user object exists in Entra
        $flag = ValidateAADUser $UserPrincipalName
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        Clear-Host
    } catch {$flag = $null}
} until ($flag)

# Create an array containing the groups to add the user too
$groupNames = @(
    'sec-azure-zpa-all-users',
    'sec-azure-miro-users',
    'AutoPilot Users (Apps)',
    'sec-azure-SSPR-Enable',
    'Sec_G_CrossTenantSyncProd'
)

$dlNames = @('DL All Users')
# Prompt user to dertermine the correct DoneSafe group
# Loop until the user selects a valid number
while($true){
    try {
         $userInput = [int](Read-Host "`nPlease choose from one of the following:`n[1]: The user reports to the CEO.`n[2]: The user has direct reports.`n[3]: None of the above.")
         # Checks if the input matches exactly '1', '2' or '3'
         if($userInput -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($userInput -eq 1) {$groupNames += "DoneSafe Z Executives"}
             elseif($userInput -eq 2){$groupNames += "DoneSafe People Leaders"}
             elseif($userInput -eq 3){$groupNames += 'DoneSafe Leaders of Self'}
             Clear-Host
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

 while($true){
    try {
         # Validates the users input is an integer
         $userInput = [int](Read-Host "`nSelect the location of the user:`n[1]: Wellington.`n[2]: Auckland.`n[3]: Christchurch.")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($userInput -match '\b[1-3]\b'){
             # Check the value of $userInput and add the corresponding value to the $dlNames array
             if ($userInput -eq 1) {$dlNames += "DL WEL Users"}
             elseif($userInput -eq 2){$dlNames += "DL Te Whare Rama"}
             elseif($userInput -eq 3){$dlNames += 'DL CHC Users'}
             Clear-Host
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

# Loops through each value in $groupName array and adds the user to it. This is done using the RESTful API call passing $user.Id as the paramater
foreach ($groupName in $groupNames) {
    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    [string]$groupId = $group.id
    if ($group) {
        try { 
            $body = @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"}
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $body -ErrorAction Stop
            Write-Host "$($user.DisplayName) successfully added to '$($group.DisplayName)' in Entra ID" -ForegroundColor Green           
        } catch {
            $errorMessage = $_.Exception.Message
            # Write-Host  "Raw error  message:  $errorMessage" #Debugging line
            if($errorMessage -match "One or more added object references already exist"){
                Write-Host "$($user.DisplayName) is already a member of '$groupName'" -ForegroundColor Yellow
            }else{
                Write-Host "An unexpected error occured" -ForegroundColor Red
            }
        }
    } else {Write-Host "Group '$groupName' not found." -ForegroundColor Red}
}

try{
    # Set usage location to NZ
    Update-MgUser -UserId $user.Id -UsageLocation "NZ"
    Write-Host "Usage location has been set to New Zealand in Entra ID" -ForegroundColor Green
}catch{Write-Host $_}

# Start-Sleep -Seconds 10
# Assign Viva Insights license
$vivaLicenseSKU = '3d957427-ecdc-4df2-aacd-01cc9d519da8'
$E5LicenseSKU = '06ebc4ee-1bb5-47dd-8120-11324bc54e06'
$licenseDetails =  Get-MgUserLicenseDetail  -UserId $user.UserPrincipalName

# Check if the user has an existing license
if($licenseDetails.SkuId -contains $E5LicenseSKU){
    Write-Host "E5 License assigned via AutoPilot Users (Apps) security group" -ForegroundColor Green
}else{
    Write-Host "$($User.DisplayName) assignment failed via Autopilot group. Please revert to manual allocation" -ForegroundColor Red
}

if ($licenseDetails.SkuId -contains $vivaLicenseSKU) {
    Write-Host "$($User.DisplayName) already has a Viva Insights License" -ForegroundColor Yellow
}else{
    Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = '3d957427-ecdc-4df2-aacd-01cc9d519da8'} -RemoveLicenses @() | Out-Null
    Write-Host "Assigend Microsoft Viva Insights License via M365 Admin Center" -ForegroundColor Green
}

Disconnect-MgGraph *> $null

Connect-ExchangeOnline -ShowBanner:$false

# Loop through the $dlNames array and add the user to each DL
foreach($dl in $dlNames){
    try{
        Add-DistributionGroupMember -Identity $dl -Member $user.UserPrincipalName -ErrorAction Stop
        Write-Host "$($user.DisplayName) successfully added to '$dl' in Exchange Online" -ForegroundColor Green
    }catch{
        if($errorMessage -match "One or more added object references already exist"){
            Write-Host "$($user.DisplayName) is already a member of '$dl'" -ForegroundColor Yellow
        }else{
            Write-Host "An unexpected error occured - Please continue with manual steps" -ForegroundColor Red
        }
    }
 }


