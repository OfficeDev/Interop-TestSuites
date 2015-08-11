
$script:ErrorActionPreference = "Stop"
$userName= .\Get-ConfigurationPropertyValue.ps1 UserName1
$password= .\Get-ConfigurationPropertyValue.ps1 Password1
$domain= .\Get-ConfigurationPropertyValue.ps1 Domain
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$requestUrl= .\Get-ConfigurationPropertyValue.ps1 TargetSiteCollectionUrl

#get $listName as String
$listName= .\Get-ConfigurationPropertyValue.ps1 MSFSSHTTPFSSHTTPBLibraryName

$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

# sent scripts to server.
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
  param(
      [string]$siteUrl,
      [string]$listName,
      [bool]$isCheckoutRequired
      )

  # load assemblies.
  [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null

  try
  {
	  $spSite = new-object Microsoft.SharePoint.SPSite($siteUrl)
	  $spweb = $spSite.OpenWeb()
	  $spList = $spweb.Lists[$listName]
	  $spList.ForceCheckout = $isCheckoutRequired
	  $spList.Update()
	  $spSite.Close();
  }
  finally
  {
	$spSite.Dispose()
  }
} -argumentlist $requestUrl, $listName, $isCheckoutRequired
  
if(!$?)
{
	return $false
}

return $true