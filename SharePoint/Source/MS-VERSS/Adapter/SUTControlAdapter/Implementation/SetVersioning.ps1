$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 RequestUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$listName,
       [bool]$enableVersioning,
       [string]$requestUrl,       
       [bool]$enableMinorVersions
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  $ret=$false
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite $requestUrl
      $spWeb =  $spSites.openweb()

      $List = $spWeb.Lists[$listName]
      if($enableVersioning -eq $true)
      {
          $List.EnableVersioning  = $enableVersioning;
          $List.EnableMinorVersions = $enableMinorVersions;
      }
      else
      {
          $List.EnableVersioning  = $enableVersioning;
          $List.EnableMinorVersions = $false;
      }
      $List.Update()
      $ret=$true
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
  return $ret
} -argumentlist $listName, $enableVersioning, $requestUrl, $enableMinorVersions

return $ret