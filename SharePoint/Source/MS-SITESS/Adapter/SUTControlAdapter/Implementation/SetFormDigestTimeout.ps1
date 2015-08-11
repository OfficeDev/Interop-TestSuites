$script:ErrorActionPreference = "Stop"
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$siteCollectionUrl = .\Get-ConfigurationPropertyValue.ps1 SiteCollectionUrl
$mainUrl = "http://" + $computerName

$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{

  param(
       [int]$timeout,
       [string]$mainUrl,
       [string]$siteCollectionUrl
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  [int]$newTimeout=0
  try
  {
      if([string]::IsNullOrEmpty($siteCollectionUrl))
      {
          $spSite = new-object Microsoft.SharePoint.SPSite($mainUrl)
      }
      else
      {
          $spSite = new-object Microsoft.SharePoint.SPSite($siteCollectionUrl)
      }
      $webApp = $spSite.WebApplication
      $webApp.FormDigestSettings.Timeout = new-object system.timespan(0, 0, $timeout)
      $webApp.Update()
      $totalSecondes= $webApp.FormDigestSettings.Timeout.TotalSeconds
      $newTimeout=[int]$totalSecondes
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      if($spSite -ne $null)
      {
          $spSite.Dispose()
      }
  }
  return $newTimeout
  }-argumentlist $timeout, $mainUrl, $siteCollectionUrl

  return $ret