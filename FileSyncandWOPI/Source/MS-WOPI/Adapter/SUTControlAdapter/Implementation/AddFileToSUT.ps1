$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType

$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetSiteCollectionUrl

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$documentLibraryName,
       [string]$requestUrl,
	   [string]$fileName,
       [string]$transportType
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");

	$spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
	$spWeb =  $spSites.RootWeb
	$targetDocList = $spWeb.Lists[$documentLibraryName]
	$listFolder = $targetDocList.RootFolder

	if($listFolder -ne $null)
	{
		$folderName = $listFolder.Name
		$Files = $listFolder.Files
		$TimeFormat = [System.DateTime]::Now.ToString("yyyyHHmmss_fff")
		$fileData=[System.Text.Encoding]::Unicode.GetBytes("MS-WOPI Test purpose, generated on [$TimeFormat]")
		$addedFile = $Files.Add($FileName, $fileData)
		$pathOfAddedFile = $transportType + "://" + $spSites.HostName + $addedFile.ServerRelativeUrl
		$ret = $pathOfAddedFile
	}

	$spSites.Dispose()
    return $ret
} -argumentlist $documentLibraryName, $requestUrl, $fileName, $transportType

return $ret