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

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 RequestUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$listName,
       [string]$fileName,
       [string]$fileVersion,
       [string]$requestUrl
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite $requestUrl
      $spWeb =  $spSites.openweb()
      $File = $spWeb.GetFolder($listName).Files[$fileName]
      $fileInformations = $File.Versions
      [string]$result = ""
      if ($File.UIVersionLabel -eq $fileVersion)
      {
          $result = $result + $File.ModifiedBy.Name + "^"
          $result = $result + $File.Length
      }
      else
      {
          for($i = 0; $i -le $fileInformations.Count - 1; $i++)
          {
              if ($fileVersion -eq $fileInformations[$i].VersionLabel)
              {
                  $fileInformation = $fileInformations[$i]
                  $result = $result + $fileInformation.CreatedBy.Name + "^"
                  $result = $result + $fileInformation.Size
                  break
              }
          }
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
  return $result
} -argumentlist $listName, $fileName, $fileVersion, $requestUrl

return $ret