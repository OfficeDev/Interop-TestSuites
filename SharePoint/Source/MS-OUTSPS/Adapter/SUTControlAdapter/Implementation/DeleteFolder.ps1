#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

#---------------------------------------------------------------------------------
# Purpose: Delete specified folder by folder name.
#---------------------------------------------------------------------------------
$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName

$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$listTitle,
       [string]$subfolderName,
       [string]$requestUrl
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
  function CloseSPSite 
  {    
      param(
      [Microsoft.SharePoint.SPSite]$SPSite
      )
      if($spSite -ne $null)
      {
          $spSite.Dispose()
      }
  }
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
      $spWeb =  $spSites.RootWeb
      $targetDocList = $spWeb.Lists[$listTitle]
      $listFolder = $targetDocList.RootFolder
    
      if($listFolder -ne $null)
      {  
          $subFolders = $listFolder.SubFolders
          if($listFolder.SubFolders.count -eq 0)
          {  
              $ret = $FALSE
              return 
          }

          $targetSubFolder = $subFolders[$subfolderName]
          if($targetSubFolder -eq $null)
          {
              $ret = $FALSE
              return 
          }
          $targetSubFolder.Delete()
          $ret = $true
       }
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      CloseSPSite $spSites
  }
  return $ret

} -argumentlist $listTitle, $subfolderName, $requestUrl

return $ret