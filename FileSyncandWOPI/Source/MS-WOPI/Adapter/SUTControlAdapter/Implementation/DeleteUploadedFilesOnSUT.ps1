$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType
$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetSiteCollectionUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$currentDoclibraryListName,
       [string]$requestUrl,
       [string]$uploadedfilesUrls
  )
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  try
  {    
      $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
	  $spWeb =  $spSites.RootWeb
      $folderitems = $spWeb.Folders
      $listFolder = $folderitems[$currentDoclibraryListName]
      $ErrorCounter = 0
        if($listFolder -ne $null)
        {  
           $Files = $listFolder.Files
		   $urlsArray = $uploadedfilesUrls.Split(',')
           foreach($urlItem in $urlsArray)
           {  
              try
              {
                $Files.Delete($urlItem);
              }
              catch
              {  
                $ErrorCounter ++
              }
           }
         }

	   $spSites.Dispose()
       $ret = $ErrorCounter -eq 0
       return $ret
  }
  catch
  {
    $ret = $false
  }
} -argumentlist $currentDoclibraryListName, $requestUrl, $uploadedfilesUrls

return $ret