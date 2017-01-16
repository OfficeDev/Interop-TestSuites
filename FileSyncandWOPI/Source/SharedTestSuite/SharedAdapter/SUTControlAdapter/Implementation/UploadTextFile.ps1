$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName1
$password = .\Get-ConfigurationPropertyValue.ps1 Password1
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$documentLibraryUrl,
	   [string]$fileName,
       [string]$transportType
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");

    $siteCollectionUrl = $documentLibraryUrl.Substring(0, $documentLibraryUrl.LastIndexOf("/"))
	$documentLibraryName = $documentLibraryUrl.Substring($documentLibraryUrl.LastIndexOf("/") + 1)
	$spSites = new-object Microsoft.SharePoint.SPSite "$siteCollectionUrl"
	$spWeb =  $spSites.RootWeb
	$targetDocList = $spWeb.Lists[$documentLibraryName]
	$listFolder = $targetDocList.RootFolder

	if($listFolder -ne $null)
	{
		$folderName = $listFolder.Name
		$Files = $listFolder.Files
		$TimeFormat = [System.DateTime]::Now.ToString("yyyyHHmmss_fff")
		$fileData=[System.Text.Encoding]::Unicode.GetBytes("Test file content, generated on [$TimeFormat]")
		$addedFile = $Files.Add($FileName, $fileData)
		$pathOfAddedFile = $transportType + "://" + $spSites.HostName + $addedFile.ServerRelativeUrl
	}

	$spSites.Dispose()
} -argumentlist $documentLibraryUrl, $fileName, $transportType

if(!$?)
{
	# return $false
	throw $Error[0]
}

return $true