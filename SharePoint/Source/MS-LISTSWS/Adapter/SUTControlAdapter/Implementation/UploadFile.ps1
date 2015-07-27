#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

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
    $folderitems = $spWeb.Folders
    $listFolder = $folderitems[$documentLibraryTitle]
    $ret = $null

    if($listFolder -ne $null)
    {
       $folderName = $listFolder.Name
       $Files = $listFolder.Files
       $TimeFormat = [System.DateTime]::Now.ToString("HHmmss_fff")
       $FileName = "MSLISTSW_" + $TimeFormat + ".txt"
       $fileData=[System.Text.Encoding]::Unicode.GetBytes("MSLISTSWSTEST Test on [$TimeFormat]")
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
       if ($spSites -ne $null)
       {
           $spSites.Dispose()
       }
  }
  return $ret
} -argumentlist $documentLibraryTitle, $requestUrl, $transportType

return $ret