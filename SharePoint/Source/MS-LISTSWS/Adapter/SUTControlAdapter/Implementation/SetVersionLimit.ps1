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
       [string]$listId,
       [int]$majorVersionLimit,
	   [int]$minorVersionLimit,
       [string]$requestUrl
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  try
  {
	  $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
	  $spWeb =  $spSites.openweb()

	  $List = $spWeb.Lists | where {$_.ID -eq $listId}
	  if($List.EnableVersioning -eq $true)
	  {
		$List.EnableMinorVersions = $true;
		$List.MajorVersionLimit = $majorVersionLimit;
		$List.MajorWithMinorVersionsLimit = $minorVersionLimit;
	  }
	  else
	  {
		$List.EnableVersioning = $true;
		$List.EnableMinorVersions = $true;
		$List.MajorVersionLimit = $majorVersionLimit;
		$List.MajorWithMinorVersionsLimit = $minorVersionLimit;
	  }
	  $List.Update() 
	  return $true
  }
  catch
  {
	return $false
  }
  finally
  {
	 if ($spSites -ne $null)
     {
        $spSites.Dispose()
     }
  }
} -argumentlist $listId, $majorVersionLimitValue, $majorWithMinorVersionsLimitValue, $requestUrl

return $ret