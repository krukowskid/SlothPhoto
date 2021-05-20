
Function Write-ToLogFile()
{
  Param(
    [Parameter(Mandatory=$True)]
    [String]$Message,

    [Parameter(Mandatory=$True)]
    [String]$FileName
  )
  [String]$TimeStamp = (Get-Date -Format yyyy-MM-dd_HH:mm:ss)
  Add-content $FileName -value "[$TimeStamp] $Message"
}
