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
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

#Convert password from plain test to secure strings
$securePassword = convertTo-SecureString $password -AsPlainText -Force

#Initializes a new instance of the System.Net.NetworkCredential class with the specified user name and password read from ptfconfig file
$credential=new-object Management.Automation.PSCredential(($domain+"\"+$userName), $securePassword)

#Execute the following script remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
	param( [string] $url , [string]$listName)

	$script:ErrorActionPreference = "Stop"

    #Load Assemblies Microsoft.SharePoint.dll PSPowerShell.dll
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
        
    #Get SharePoint List
    $site = new-object Microsoft.SharePoint.SPSite($url)
	
	if ($site -eq $null) {
      return $false
	}

	$web = $site.OpenWeb() 
	$list = $web.Lists[$listName]

	if ($list -eq $null) {
      return $false
	}

	$items = $list.Items

	for($i=$items.Count-1; $i -ge 0;  $i--)
	{
		# Document Set is not file we submitted and is used to test, so ignore it.
		if($items[$i].ContentType.Name -ne "Document Set")
		{
			$items[$i].Delete()
		}
	}

	# Delete folder in list if exist
	if($list.Folders.Count > 0)
	{
		for($i=$list.Folders.Count-1; $i -ge 0;  $i--)
		{
			$list.Folders[$i].Delete()
		}
	}

	return $true

} -argumentlist $siteUrl, $listName 

return $ret