#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"

#PTF Properties
$ServerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$Domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$UserName = .\Get-ConfigurationPropertyValue.ps1 UserName
$Password = .\Get-ConfigurationPropertyValue.ps1 Password

# Escape and secure the password
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
# Create a PSCredential for remote invoke
$credential = New-Object Management.Automation.PSCredential(($Domain+"\"+$UserName),$securePassword)

# Execute script on remote server
$ret = Invoke-Command -ComputerName $ServerName -Credential $credential -ArgumentList $siteUrl -ScriptBlock {
  param(
       [string]$siteUrl
  )
  $script:ErrorActionPreference = "Stop"
  # Load SharePoint library
  [reflection.assembly]::Loadwithpartialname("Microsoft.SharePoint") | out-null
  # Get the site collection
  try
  {
      $siteCollection = new-object Microsoft.SharePoint.SPSite($siteUrl)
      if(!$?){throw $error[0]}
      # Open the specified site
      $site = $siteCollection.OpenWeb()
      if(!$?){throw $error[0]}
      # Delete all subsites
      foreach($subSite in $site.Webs)
      {
          $subSite.Delete()
          if(!$?){throw $error[0]}
      }
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      if($siteCollection -ne $null)
      {
          $siteCollection.Dispose()
      }
  }
  }
  # Check if the remote execution failed
  if(!$?){throw $error[0]}
  return $true