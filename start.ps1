$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups", "ExchangeOnlineManagement")

# Loop through each module defined in the $requiredModules array and check if its installed and imported.
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

Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.ReadWrite.All", "Directory.Read.All" #"RoleManagement.Read.Exchange"

# Check if the user exists
function ValidateAADUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )

    # Check that the user exists. This function is called immediatley after the user inputs the UPN
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
         $userInput = [int](Read-Host "Please choose from one of the following:`n[1]: The user reports to the CEO.`n[2]: The user has direct reports.`n[3]: None of the above.")
         # Checks if the input matches exactly '1', '2' or '3'
         if($userInput -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($userInput -eq 1) {$groupNames += "DoneSafe Z Executives"}
             elseif($userInput -eq 2){$groupNames += "DoneSafe People Leaders"}
             elseif($userInput -eq 3){$groupNames += 'DoneSafe Leaders of Self'}
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

 while($true){
    try {
         # Validates the users input is an integer
         $userInput = [int](Read-Host "Select the location of the user:`n[1]: Wellington.`n[2]: Auckland.`n[3]: Christchurch.")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($userInput -match '\b[1-3]\b'){
             # Check the value of $userInput and add the corresponding value to the $dlNames array
             if ($userInput -eq 1) {$dlNames += "DL WEL Users"}
             elseif($userInput -eq 2){$dlNames += "DL Te Whare Rama"}
             elseif($userInput -eq 3){$dlNames += 'DL CHC Users'}
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
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $body
            Write-Host "$($user.DisplayName) successfully added to '$($group.DisplayName)' in Entra ID" -ForegroundColor Green
        } catch {Write-Host "Error adding to '$groupName': $_" -ForegroundColor Red}
    } else {Write-Host "Group '$groupName' not found." -ForegroundColor Red}
}

try{
    # Set usage location to NZ
    Update-MgUser -UserId $user.Id -UsageLocation "NZ"
    Write-Host "Usage location has been set to New Zealand in Entra ID" -ForegroundColor Green
}catch{Write-Host $_}

# Pause for 5 seconds to avoid DDOS suspicion
Start-Sleep -Seconds 2
Connect-ExchangeOnline -DisableWAM -ShowBanner:$false

# Loop through the $dlNames array and add the user to each DL
foreach($dl in $dlNames){
    Add-DistributionGroupMember -Identity $dl -Member $user.UserPrincipalName
    Write-Host "$($user.DisplayName) successfully added to '$dl' in Exchange Online" -ForegroundColor Green
 }

 try{
    # Pause for 5 seconds to avoid DDOS suspicion
    Start-Sleep -Seconds 5
    # Assign Viva Insights license
    Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = '3d957427-ecdc-4df2-aacd-01cc9d519da8'} -RemoveLicenses @()
    Write-Host "Assigend Microsoft Viva Insights License via M365 Admin Center" -ForegroundColor Green
    Write-Host "Assigned MS E5 license via AutoPilot Users (Apps) security group" -ForegroundColor Green
 }catch{
    Write-Host $_
 }

# Create an array containing the emails to add to the users TrustedSendersAndDomains properties
$safeSenders = @(
    "matt.halliday@ampol.com.au",
    "sdm@ampol.com.au",
    "communications@ampol.com.au",
    "brent.merrick@ampol.com.au",
    "support@txn.mail.rewardgateway.net",
    "reply@txn.mail.rewardgateway.net"
)

$target = Get-Mailbox $user.UserPrincipalName

# Loop through the $safeSenders array and add each value to the users TrustedSendersAndDomains property in EXO
foreach($thissender in $safeSenders){
    try{
        Set-MailboxJunkEmailConfiguration $target.Name -TrustedSendersAndDomains @{Add = $thissender}
        Write-Host "Successfully added $thissender to the Trusted Senders and Domains list" -ForegroundColor Green
    }catch{
        Write-Host $_
    }
}



