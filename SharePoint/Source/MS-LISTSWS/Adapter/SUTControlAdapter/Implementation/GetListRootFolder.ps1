$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$listName,
       [string]$requestUrl
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  try
  {
      $directionValue = $null
	  $spSites = new-object Microsoft.SharePoint.SPSite $requestUrl
	  $spWeb =  $spSites.openweb()

	  $List = $spWeb.Lists[$listName]
	  $directionValue = $List.RootFolder.ServerRelativeUrl
	  return $directionValue
  }
  catch
  {
	return $directionValue
  }
  finally
  {
	    if ($spSites -ne $null)
        {
            $spSites.Dispose()
        }
  }
} -argumentlist $listName, $requestUrl

return $ret