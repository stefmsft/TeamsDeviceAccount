#Function used by the script
#
#This EASPoliciy function come from the legacy script written by Eric Scherlinger
function EASPolicy()
#Will check customer env for a suitable EAS Policy. If not will create. function returns the name of the EASPolicy Name as string.
{
    write-host "Configuring ActiveSync Policy for your environment..."
    $easpolicy =$null
    $easpolicy = Get-MobileDeviceMailboxPolicy | Where-Object {$_.PasswordEnabled -eq $false -and $_.name -notlike "default"}

    if($easpolicy)
        {

            $i = 1
            $easPolicy | % {
            Write-Host -NoNewLine $i
            Write-Host -NoNewLine ": EASPolicyName: "
            Write-Host  $_.Name
            $i++
                    }
            $ieaspolicy = 0;
            if($easPolicy.Length)
                {
                do{$ieaspolicy = Read-Host 'Choose the number for the EAS Policy you want to pick or type "n" for a new EAS Policy'} 
                while ($ieaspolicy -ne "n" -and( $ieaspolicy -lt 1 -or $ieaspolicy -gt $easPolicy.Length)) 
                }
            else 
                {
                do{$ieaspolicy = Read-Host 'Choose the number for the EAS Policy you want to pick or type "n" for a new EAS Policy'} 
                while ($ieaspolicy -ne "n" -and $ieaspolicy -ne 1 )  
                }

            if($ieaspolicy -eq "n"){
                $strPolicy = Read-Host 'Please enter the name for a new device ActiveSync policy that will be created and applied to this account.
                We will configure that policy to be compatible with Surface Hub/MTR devices.'  
                $easpolicy = New-MobileDeviceMailboxPolicy -Name $strPolicy -PasswordEnabled $false -AllowNonProvisionableDevices $true -ErrorAction SilentlyContinue
            }
            else{
                $easpolicy = $easpolicy[$ieaspolicy - 1]
                $strPolicy = $easpolicy.name
                }
            write-host "We Will use $strpolicy for the device."
        }  
    else 
        {
        $strPolicy = Read-Host 'Please enter the name for a new Surface Hub ActiveSync policy that will be created and applied to this account.
        We will configure that policy to be compatible with Surface Hub / MTR devices.'

        $easpolicy = New-MobileDeviceMailboxPolicy -Name $strPolicy -PasswordEnabled $false -AllowNonProvisionableDevices $true -ErrorAction SilentlyContinue
        $strpolicy = $easpolicy.name
        if($easpolicy) {write-host "A new EAS policy $strPolicy has been created." }
        else {Write-Error "We were unable to create the EAS Policy because $error"}
        }

    $strpolicy
}

#
#Main part of the script Script
#
#Gather info from an input stream if any
#Sample usage :
#       import-csv .\file.csv |% { .\New-CreateHUBAccount.ps1 }
#

$upn = $_.upn
$passw = $_.passw
$ccountry = $_.country
$EasPolicyName = $_.EasPolicyName

if ($null -eq $upn) {
    do
        {
            $upn = Read-Host 'What is the device UPN '
            $domain = $upn.Split("@")[1]
        } while (($upn -eq "") -or ($null -eq $upn) -or ($null -eq $upn.Split("@")[1]))
    $streamedImput = $True
    }
    else {
        $streamedImput = $True
    }

if ($null -eq $passw) {
    do
        {
            $passw = Read-Host "Please enter the password for $upn "
        } while ($passw -eq "")
    }

if ($null -eq $ccountry) {
    do
        {
            $ccountry = Read-Host 'What is the device country code '
        } while (($ccountry -eq "") -or ($ccountry.length -gt 2))
    }

if ($streamedImput) {
    write-host "Data gathered from CSV"
    write-host "UPN : $upn"
    write-host "Password $passw"
    write-host "Country Code : $ccountry"
    write-host "EAS Policy : $EasPolicyName"   
}



$domain = $upn.Split("@")[1]
$alias = $upn.Split("@")[0]

$admin = "admin@" + $domain

#Load required module
Import-Module ExchangeOnlineManagement
Import-module AzureAD

#Check upn
Connect-AzureAD -AccountID $admin
$NewUSerOK=$false
try {
    Get-AzureADUser -ObjectId $upn
}
catch{
    $NewUSerOK=$true
}

if(!$NewUSerOK){
    write-host "User $strUpn already exist, choose another alias ..."
    exit
}

#Make the EOL Connexion
Connect-ExchangeOnline -UserPrincipalName $admin -ShowProgress $true

#Select the EAS Policy
if ($null -eq $EasPolicyName) {
    $EasPolicyName = EASPolicy
}

write-host "One moment, We are creating the account ..."
#Create the mailbox
$user = New-Mailbox -MicrosoftOnlineServicesID $upn -Alias $alias -Name $alias -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String $passw -AsPlainText -Force)

#Check user availability
$NewUSerOK=$False
do
{
    try {
        $user = Get-AzureADUser -ObjectId $upn
        $NewUSerOK=$true
    }
    catch{
        start-sleep -Seconds 3
    }
} while ($NewUSerOK -eq $False)

#Fix some calendar behaviors 
Set-CalendarProcessing -Identity $upn -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowConflicts $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false  -AllowRecurringMeetings $true -ProcessExternalMeetingMessages $true
Set-CalendarProcessing -Identity $upn -AddAdditionalResponse $true -AdditionalResponse "This is a Microsoft Surface Hub. Please make sure this meeting is a Microsoft Teams meeting!"

#Add pasword, etc ...
$PasswordProfile = New-Object Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $passw
$PasswordProfile.EnforceChangePasswordPolicy = $false
$PasswordProfile.ForceChangePasswordNextLogin = $false

Set-AzureADUser -ObjectId $upn -AccountEnabled $True -DisplayName $alias -PasswordProfile $PasswordProfile -PasswordPolicies "DisablePasswordExpiration" -MailNickName $alias -UsageLocation $ccountry

#Assign Policy to the device
if ($null -eq $EasPolicyName) {
    Write-Verbose ("Warning : No EAS Policy applied of the device")
} else {
    try {
        Set-Mailbox $upn -Type Regular
        Set-CASMailbox -Identity $upn -ActiveSyncMailboxPolicy $EasPolicyName
        Set-Mailbox $upn -Type Room        
    } catch {
        write-error "ActiveSyncPolicy assignation failed"
    }
}

#Assign the Teams License
try {
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense 
    $License.SkuId = "c7df2760-2c81-4ef7-b578-5b5392b571df" 
    $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses 
    $Licenses.AddLicenses = $License 
    Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $Licenses    
} catch {
    write-error "Teams license assignation failed"
}

Write-Host "The account is created for $upn (See the details below)"
Get-AzureADUser -ObjectId $upn |Format-List ObjectId,AccountEnabled,DisplayName,Mail,UserPrincipalName,UsageLocation

return