#Script to disabe accounts & remove licenses and move to correct OU all based on the SOP step by step guide
#For Leavers.

#Required to connect to exchange to have access to user/sharedmailboxes
#required to connect to MsolService to remove add licenses



#Connect to Exchange & Connect to MsolService 

#Connect to MsolService 
Connect-MsolService 

# Connect-ExchangeOnline
Connect-ExchangeOnline

#Import users from csv file
$users = Import-Csv -Path "C:\script\eFiles\UsersToDisable.csv"

#Loop through the array of objects and convert each user into SharedMailbox
foreach ($user in $users) {
    Set-Mailbox -Identity $user.User -Type Shared 
} 

#Remove Microsoft Office 365  Business Standard Licenses
foreach ($user in $users) {
    Set-MsolUserLicense -UserPrincipalName $user.User -RemoveLicenses "myclub:O365_BUSINESS_PREMIUM"
}

#Block Sign in Access from microsoft office 365 
foreach ($user in $users) {
    Set-MsolUser -UserPrincipalName $user.User -BlockCredential $true
}


# Loop through each user and remove all groups except for the primary group
foreach ($user in $users) {
    # Get the user object, DistingushedName, Status, GivenName, Memberof, PrimaryGroup, SamAccountName etc...
    $userObj = Get-ADUser -Filter "UserPrincipalName -eq '$($user.User)'" -Properties MemberOf, PrimaryGroup

    # Get the DN of the primary group (Domain Users)
    $primaryGroupDN = (Get-ADGroup -Filter "Name -eq 'Domain Users'").DistinguishedName

    # Remove all groups except for the primary group
    foreach ($group in $userObj.MemberOf) {
        if ($group -ne $primaryGroupDN) {
            Remove-ADGroupMember -Identity $group -Members $userObj.SamAccountName -Confirm:$false
        }
    }
}

#Sets the users msExchHideFromAddressList to true so its not visible when searched.
foreach ($user in $users) {

    #Get the user objects 
    $userADObj = Get-ADUser -Filter "UserPrincipalName -eq '$($user.User)'"
    #sets msexchhideaddresslist to true
    Set-ADUser -Identity $userADObj -Replace @{ msExchHideFromAddressLists = $true}

}

#changes description to to add date account was disabled
$disabledDate = Get-Date -Format "yyyy-mm-dd"
foreach ($user in $users) {
    #Gets the users whole Object
    $userADObj = Get-ADUser -Filter "UserPrincipalName -eq '$($user.User)'"
    #Sets the users Description to Disabled on current date
    Set-ADUser -Identity $userADObj -Description "Disabled $disabledDate"

}

#disable Users and move them to a different OU group 
foreach ($user in $users) {
    #Gets the Users whole Object
    $userADObj = Get-ADUser -Filter "UserPrincipalName -eq '$($user.User)'"
    #Disables Users Active Directory Account
    Disable-ADAccount -Identity $userADObj

    #sets the to be deleted ou into a variable 
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq 'Temp Office365 Sync'" 

    #as theres 2 with the same name, grabs the second one in the array of objects which is users to be deleted other is to temp delete
    $ou = $ou[1]

    #moves the users object into the users to be deleted folder
    Move-AdObject -Identity $userADObj -TargetPath $ou.DistinguishedName
}

