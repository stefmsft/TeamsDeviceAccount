# Repo Purpose
This project purpose is to drop the latest scripts I made to create a functional M365 Account for Surface HUB or MTR

# Scripts

**New-TeamsDeviceAccount.ps1** script allow you to create an AAD Account that works as an working identity for HUB and MTR devices.

### Usage

If you run the script directly, it will ask for the information needed like :
- the UPN
- the password
- the country
- the EAS Policy name

... Then the account will be created

If you wish to automate the creation in a batch way, you can redirect the stream of an CSV file to the script. It will then make the creation of an account for each line of information provided

The calling syntax would be :

    import-csv .\file.csv |% { .\New-TeamsDeviceAccount.ps1 }

The header format is :

    upn,passw,country,EASPolicyName

A sample csv file is part of the repo.

NOTA: When used in pipe mode, the EASPolicy specified in the csv file must already exist.

**PreReq-TeamsDeviceAccount.ps1** script is to run only once as it justs install the needed module to run properly the New-CreateTeamsDeviceAccount script. 