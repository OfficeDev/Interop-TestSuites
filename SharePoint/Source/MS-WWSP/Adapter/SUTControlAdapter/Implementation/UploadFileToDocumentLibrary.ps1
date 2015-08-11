$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType

$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$documentLibraryTitle,
       [string]$requestUrl,
       [string]$transportType
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
   try
  {
    $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
    $spWeb =  $spSites.RootWeb
    $targetDocList = $spWeb.Lists[$documentLibraryTitle]
    $listFolder = $targetDocList.RootFolder

    if($listFolder -ne $null)
    {
       $folderName = $listFolder.Name
       $Files = $listFolder.Files
       $TimeFormat = [System.DateTime]::Now.ToString("yyyyHHmmss_fff")
       $FileName = "MSWWSP_" + $TimeFormat + ".txt"
       $fileData=[System.Text.Encoding]::Unicode.GetBytes("MSWWSP Test on [$TimeFormat]")
       $addedFile = $Files.Add($FileName, $fileData)
       $pathOfAddedFile = $transportType + "://" + $spSites.HostName + $addedFile.ServerRelativeUrl
       $ret = $pathOfAddedFile
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
  return $ret
} -argumentlist $documentLibraryTitle, $requestUrl, $transportType

return $ret