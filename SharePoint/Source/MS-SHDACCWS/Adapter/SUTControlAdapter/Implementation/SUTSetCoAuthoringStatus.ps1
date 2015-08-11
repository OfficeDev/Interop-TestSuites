$script:ErrorActionPreference = "Stop"
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName

$fileGUID = .\Get-ConfigurationPropertyValue.ps1 FileIdOfCoAuthoring
$siteCollectionUrl =.\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
  param(
      [GUID]$fileGUID,
      [string]$siteCollectionUrl
  )
  $script:ErrorActionPreference = "Stop"
  $retValue = $false
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite "$siteCollectionUrl"
      $web = $spSites.OpenWeb()
      $file = $web.GetFile($fileGUID)
      $file.CreateSharedAccessRequest()
      $file.Update() | out-null
      if($file.IsSharedAccessRequested)
      {
          $retValue = $true
      }
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      if($spSites -ne $null)
      {
          $spSites.Dispose()
      }
  }

  return $retValue

}-argumentlist $fileGUID, $siteCollectionUrl

return $ret