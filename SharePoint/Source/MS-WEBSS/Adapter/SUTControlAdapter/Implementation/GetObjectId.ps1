#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"

$password= .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$domain=.\Get-ConfigurationPropertyValue.ps1 Domain
$userName=.\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$mainUrl = .\Get-ConfigurationPropertyValue.ps1 SiteCollectionUrl

$ret = Invoke-Command -ComputerName $computerName -Credential $credential -ScriptBlock{
  param(
       [string]$webSiteName,
       [string]$objectName,
       [string]$mainUrl
  )
  $script:ErrorActionPreference = "Stop"
  [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
  try
  {
    $spSite = new-object Microsoft.SharePoint.SPSite($mainUrl)
    $spWeb =  $spSite.OpenWeb($webSiteName)
    $spPage = $spWeb.Lists["Site Pages"]

    switch($objectName)
    {
        "list"
        {
            return $spPage.ID
        }

        "listItem"
        {
            $spPageFiles = $spPage.Items
            foreach($spPageFile in $spPageFiles)
            {
                if(!$spPageFile.DisplayName.compareTo("Home")) 
                {
                    return $spPageFile.ID
                }
            }
        }

        "site_features"
        {
            $spFeatures = $spWeb.Features
            foreach($spFeature in $spFeatures)
            {
                $featureGUID  += " "+$spFeature.DefinitionId
            }
            return $featureGUID
        }

        "site_collection_features"
        {
            $spFeatures = $spSite.Features
            foreach($spFeature in $spFeatures)
            {
                $featureGUID  += " "+$spFeature.DefinitionId
            }
            return $featureGUID
        }
    }
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      if($spSite -ne $null)
      {
          $spSite.Dispose()
      }
  }

} -argumentlist $webSiteName,$objectName,$mainUrl

if(!($objectName.compareTo("list")))
{
    return $ret.GUID
}
return $ret.ToString()