﻿<#
.SYNOPSIS
Functions-PSStoredCredentials - PowerShell functions to manage stored credentials for re-use

.DESCRIPTION 
This script adds two functions that can be used to manage stored credentials
on your admin workstation.

.EXAMPLE
. .\Functions-PSStoredCredentials.ps1

.LINK
https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp
#>


Function New-StoredCredential {

    <#
    .SYNOPSIS
    New-StoredCredential - Create a new stored credential

    .DESCRIPTION 
    This function will save a new stored credential to a .cred file.

    .EXAMPLE
    New-StoredCredential

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

            Set-Variable -Name KeyPath -Scope Global -Value .\_CREDENTIALS
   

    $Credential = Get-Credential -Message "Enter a user name and password"

    $Credential.Password | ConvertFrom-SecureString | Out-File "$($KeyPath)\$($Credential.Username).cred" -Force

    # Return a PSCredential object (with no password) so the caller knows what credential username was entered for future recalls
    New-Object -TypeName System.Management.Automation.PSCredential($Credential.Username,(new-object System.Security.SecureString))

}



Function Get-StoredCredential {

    <#
    .SYNOPSIS
    Get-StoredCredential - Retrieve or list stored credentials

    .DESCRIPTION 
    This function can be used to list available credentials on
    the computer, or to retrieve a credential for use in a script
    or command.

    .PARAMETER UserName
    Get the stored credential for the username

    .PARAMETER List
    List the stored credentials on the computer

    .EXAMPLE
    Get-StoredCredential -List

    .EXAMPLE
    $credential = Get-StoredCredential -UserName admin@tenant.onmicrosoft.com

    .EXAMPLE
    Get-StoredCredential -List

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    param(
        [Parameter(Mandatory=$false, ParameterSetName="Get")]
        [string]$UserName,
        [Parameter(Mandatory=$false, ParameterSetName="List")]
        [switch]$List
        )
      Set-Variable -Name KeyPath -Scope Global -Value .\_CREDENTIALS

      if($List){
        try {
        $CredentialList = @(Get-ChildItem -Path $keypath -Filter *.cred -ErrorAction STOP)

        foreach ($Cred in $CredentialList) {
            Write-Host "Username: $($Cred.BaseName)"
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }

    }

    if ($UserName) {
        if (Test-Path "$($KeyPath)\$($Username).cred") {
        
            $PwdSecureString = Get-Content "$($KeyPath)\$($Username).cred" | ConvertTo-SecureString
            
            $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $PwdSecureString
        }
        else {
            throw "Unable to locate a credential for $($Username)"
        }

        return $Credential
    }

    }