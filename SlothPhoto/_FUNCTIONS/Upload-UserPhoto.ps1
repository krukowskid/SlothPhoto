
Function Upload-UserPhoto()
{
    #Load SPO modules
    Try
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Loading SPO Modules..."
        Add-Type    -Path ([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client").location) `
                    -ErrorAction Stop
        Add-Type    -Path ([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.runtime").location) `
                    -ErrorAction Stop
        Add-Type    -Path ([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles").location) `
                    -ErrorAction Stop
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "SPO Modules loaded."
    }
    Catch
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Unable to Load SPO Modules. Error message: $($_.Exception.Message)"
        return
    }
    Try
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Loading Drawing Module"
        Add-Type    -AssemblyName System.Drawing `
                    -ErrorAction Stop
    }
    Catch
    {
        Write-ToLogFile -FileName $LogFileFullName -Message "Unable to Load Drawing module. Error message: $($_.Exception.Message)"
        return
    }
    #Connect to Exchange online 
    
    Try
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Connecting to Exchange Online"
        $exchangeCredentials = Get-StoredCredential -UserName $exchangeAdministrator `
                                                    -ErrorAction Stop
        $Session = New-PSSession    -ConfigurationName Microsoft.Exchange `
                                    -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS `
                                    -Credential $exchangeCredentials `
                                    -Authentication Basic `
                                    -AllowRedirection `
                                    -ErrorAction Stop

        Import-PSSession $Session -ErrorAction Stop
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Connected to Exchange Online using $exchangeAdministrator account."
    }
    Catch
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Unable to connect to Exchange Online. Error message: $($_.Exception.Message)"
        return
    }

#Connect to Sharepoint online
    Try
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message $sharepointAdministrator
        $sharepointCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials(
                                        (Get-StoredCredential -UserName $sharepointAdministrator).UserName, `
                                        (Get-StoredCredential -UserName $sharepointAdministrator).Password
                                    ) -ErrorAction Stop
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($sharepointSiteURL) -ErrorAction Stop
        $Context.Credentials = $sharepointCredentials
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Connected to Sharepoint Online using $sharepointAdministrator account."
    }
    Catch
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Unable to connect to Sharepoint Online. Error message: $($_.Exception.Message)"
        return 
    }


    #Get folder name
    $rootFolderName = "User Photos"
    $rootFolderUrl = "/"+$rootFolderName
    $web = $Context.Web.Folders.GetByUrl($rootFolderUrl)
    $context.Load($web)
    $context.Load($web.Folders)
    $context.ExecuteQuery()
    $profilePicturesFolder = ($web.Folders `
                                | Where-Object {$_.Name -ne "Forms"} `
                                    | Select-Object Name).Name

    if($profilePicturesFolder.count -ne 1)
    {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "There is more than one subfolder in $SiteURL$rootFolderUrl location. Please specify Profile Images folder name in configuration file."
        return
    }

    $uploadFolderUrl = $rootFolderUrl + "/" + $profilePicturesFolder
    $photoResolution = @{"_SThumb" = "48"; "_MThumb" = "72"; "_LThumb" = "300"}


    [byte[]]$userPhoto = [System.Convert]::FromBase64String("$base64Photo")
    $photoStreamFromGUI = New-Object    -TypeName 'System.IO.MemoryStream' `
                                        -ArgumentList (,$userPhoto)

    try
    {
        $userOnlineUPN = (Get-User -Filter "WindowsEmailAddress -eq '$userEmail'").UserPrincipalName
    }
    Catch
        {
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Unable to get user online UPN for $userEmail. Error message: $($_.Exception.Message)"
        return
    }
    
    Foreach($photo in $photoResolution.GetEnumerator())
    {
    #Covert image into different size of image
        $image = [System.Drawing.Image]::FromStream($photoStreamFromGUI)
        [int32]$width = $photo.Value
        [int32]$height = $photo.Value
        $newImage = New-Object System.Drawing.Bitmap($width, $height)
        $drawing = [System.Drawing.Graphics]::FromImage($newImage)
        $drawing.DrawImage($image, 0, 0, $width, $height)

    #Covert image into memory stream
        $memoryStream = New-Object -TypeName System.IO.MemoryStream
        $imageFormat = [System.Drawing.Imaging.ImageFormat]::Jpeg
        $newImage.Save($memoryStream, $imageFormat)
        $streamseek = $memoryStream.Seek(0, [System.IO.SeekOrigin]::Begin)

    #Upload image into sharepoint online
        $fullFileName=(($userOnlineUPN).Replace("@", "_").Replace(".", "_"))+$photo.Name+".jpg"
        $destinationURL=($uploadFolderUrl+"/"+$FullFilename).Replace(" ","%20")
        [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($context,$destinationURL, $memoryStream, $true)
        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "SharePoint online image $fullFileName uploaded successful for $user.mail"

        if($photo.value -eq 300)
        {
            try
            {
                Set-UserPhoto   -Identity $userEmail `
                                -PictureData ($memoryStream.ToArray()) `
                                -Confirm:$false `
                                -ErrorAction Stop

                Write-ToLogFile -FileName $LogFileFullName `
                                -Message "Exchange online image uploaded successful for $userEmail"

                $Session | Remove-PSSession
            }
            Catch
            {
                Write-ToLogFile -FileName $LogFileFullName `
                                -Message "Unable to set Exchange online photo. It is OK if user has onpremises mailbox. Error message: $($_.Exception.Message)"
                $Session | Remove-PSSession
                try
                {   
                    $SessionOnpremises = New-PSSession  -ConfigurationName Microsoft.Exchange `
                                                        -ConnectionUri $onpremisesExchangePowershellUri `
                                                        -Authentication Kerberos `
                                                        -ErrorAction Stop -
                    Import-PSSession $SessionOnpremises
                    Write-ToLogFile -FileName $LogFileFullName `
                                    -Message "Connected to Exchange On-premises"
                    Set-UserPhoto   -Identity $userEmail `
                                    -PictureData ($memoryStream.ToArray()) `
                                    -Confirm:$false `
                                    -ErrorAction Stop
                    Write-ToLogFile -FileName $LogFileFullName `
                                    -Message "Exchange 2016 image uploaded successful for $userEmail"
                }
                Catch
                {
                    Write-ToLogFile -FileName $LogFileFullName `
                                    -Message "Unable to set Onpremise Exchange online photo. It is OK if user has mailbox in Exchange 2010. Gosh... so many problems in this project..."
                }
                Finally
                {
                    $SessionOnpremises | Remove-PSSession
                }        
            }
            try
            {
                Set-ADUser  -identity $userSAM `
                            -Replace @{thumbnailPhoto=($memoryStream.ToArray())} `
                            -ErrorAction Stop
                Write-ToLogFile -FileName $LogFileFullName `
                                -Message "AD thumbnail photo uploaded successfully for $userSAM"
            }
            Catch
            {
                Write-ToLogFile -FileName $LogFileFullName `
                                -Message "Unable to setup AD thumbnail photo. Error message: $($_.Exception.Message)"
                return
            }
        }
    }

    $profileImageURL=$sharepointSiteURL+$uploadFolderUrl+"/"+(($userOnlineUPN).Replace("@", "_").Replace(".", "_"))+"_MThumb.jpg"
    try
    {
        Set-SharepointValue -targetAccount $userOnlineUPN  `
                            -PropertyName PictureURL `
                            -Value $profileImageURL `
                            -SharepointAdminPortalUrl $SPOAdminPortalUrl `
                            -Credentials $sharepointCredentials

        Set-SharepointValue -targetAccount $userOnlineUPN `
                            -PropertyName SPS-PicturePlaceholderState `
                            -Value 0 `
                            -SharepointAdminPortalUrl $SPOAdminPortalUrl `
                            -Credentials $sharepointCredentials

        Set-SharepointValue -targetAccount $userOnlineUPN  `
                            -PropertyName SPS-PictureExchangeSyncState `
                            -Value 0 `
                            -SharepointAdminPortalUrl $SPOAdminPortalUrl `
                            -Credentials $sharepointCredentials

        Set-SharepointValue -targetAccount $userOnlineUPN `
                            -PropertyName SPS-PictureTimestamp `
                            -Value 63605901091 `
                            -SharepointAdminPortalUrl $SPOAdminPortalUrl `
                            -Credentials $sharepointCredentials

        Write-ToLogFile -FileName $LogFileFullName `
                        -Message "Sharepoint and Deleve profile photo set for user $userEmail"
    }
    Catch
    {
        Write-ToLogFile -FileName $LogFileFullName -Message "Unable to setup SPO thumbnail photo. Error message: $($_.Exception.Message)"
        return 
    }
    return "Success"
}
