
Param(
  [Parameter(Mandatory=$True)]
  [String]$userEmail,

  [Parameter(Mandatory=$True)]
  [String]$userSAM,

  [Parameter(Mandatory=$True)]
  [String]$base64Photo
)

$base64Photo = Get-content -path $base64Photo
$functions = Get-ChildItem ("_FUNCTIONS") `
                | Where-Object { $_.name -like "*.ps1" }
foreach($function in $functions)
{
    . ($function.DirectoryName+'\'+$function.name)
}

$GlobalParamTable = Get-Content "configuration.ini" `
                        | ConvertFrom-StringData

$GlobalParamTable | foreach-object {
    $Name = [string]($_.Keys) 
    $Value = [string]($_.Values)
    New-Variable    -Name $Name `
                    -Value $Value
}

$LogFileFullName = "_LOGS\$PrefixLogFileName$SufixLogFileName"
Write-ToLogFile -FileName $LogFileFullName `
                -Message "================================================="
Write-ToLogFile -FileName $LogFileFullName `
                -Message "Starting SlothPhoto V1.0 as $env:UserName user..."
Write-ToLogFile -FileName $LogFileFullName `
                -Message "================================================="

if($sharepointSiteURL -eq "https://yourtenantname-my.sharepoint.com")
{
    Write-ToLogFile -FileName $LogFileFullName `
                    -Message "Please provide your tentant information in configuration file."
    return
}

elseif($SPOAdminPortalUrl -eq "https://yourtenantname-admin.sharepoint.com")
{
    Write-ToLogFile -FileName $LogFileFullName `
                    -Message "Please provide your tentant information in configuration file"
    return
}

Upload-UserPhoto

Write-ToLogFile -FileName $LogFileFullName `
                -Message "My job it is done. I am going to sleep. Goodnight."
