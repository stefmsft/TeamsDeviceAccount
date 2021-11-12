do
{
    $upn = Read-Host 'What is the device UPN '
    $domain = $upn.Split("@")[1]
} while (($upn -eq "") -or ($null -eq $upn) -or ($null -eq $upn.Split("@")[1]))

$domain = $upn.Split("@")[1]

$admin = "admin@" + $domain

#Load required module
Import-Module ExchangeOnlineManagement
Import-module AzureAD

#Check upn
Connect-AzureAD -AccountID $admin
Connect-ExchangeOnline -UserPrincipalName $admin -ShowProgress $true

Get-AzureADUser -ObjectId $upn
Get-Mailbox -Identity $upn