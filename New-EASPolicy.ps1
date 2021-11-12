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

do
    {
        $domain = Read-Host 'What is the Tenant Domain '
    } while (($domain -eq "") -or ($null -eq $domain))
    
$admin = "admin@" + $domain

#Load required module
Import-Module ExchangeOnlineManagement
Import-module AzureAD

#Check upn
Connect-AzureAD -AccountID $admin
#Make the EOL Connexion
Connect-ExchangeOnline -UserPrincipalName $admin -ShowProgress $true

EASPolicy