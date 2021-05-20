
Function Set-SharepointValue()
{
  Param(
    [Parameter(Mandatory=$True)]
    [String]$targetAccount,

    [Parameter(Mandatory=$True)]
    [String]$PropertyName,

    [Parameter(Mandatory=$False)]
    [String]$Value, 

    [Parameter(Mandatory=$True)]
    [String]$SharepointAdminPortalUrl,

    [Parameter(Mandatory=$True)]
    [System.Net.ICredentials]$Credentials

  )
  $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharepointAdminPortalUrl)
  $context.Credentials = $Credentials
  $peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($context)
  $targetAccount = ("i:0#.f|membership|" + $targetAccount)
  $peopleManager.SetSingleValueProfileProperty($targetAccount, $PropertyName, $Value)
  $context.ExecuteQuery()
}
