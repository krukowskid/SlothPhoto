# SlothPhoto
Old PowerShell scripts for easy photo management in ActiveDirectory, Azure, Teams, Sharepoint and all related services.
It was working well as 'backend' in web (php) photo management application, but should work as standalone.

## How it works?
Main file gets Active Directory user email, samAccountName and 300x300px image encoded in base64 format as imput and uploads it directly to Sharepoint Online, Exchange Online/OnPremises, Delve, Active Directory so you do not have to wait for replication process which could take from 24 hours to ∞.

## Installation
1. Clone this repository
2. Install `sharepointclientcomponents_16-6906-1200_x64-en-us.msi`
3. Set variables in `configuration.ini`
4. Generate credentials using functions from `Functions-PSStoredCredentials.ps1` file and place them in `_CREDENTIALS` folder.

## Usage
Run main file with params
```PowerShell
cd SlothPhoto
. .\SlothPhoto.ps1 -UserEmail john.doe@contoso.com -userSAM jdoe@contoso.corp -base64Photo $base64string
```

If you need help just open an issue. It is bit old project but I am pretty sure that many people will find it useful as a refernce or ready solution.
