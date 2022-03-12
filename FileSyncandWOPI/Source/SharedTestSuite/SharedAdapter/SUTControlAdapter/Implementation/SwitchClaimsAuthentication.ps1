
$script:ErrorActionPreference = "Stop"
$userName= .\Get-ConfigurationPropertyValue.ps1 UserName1
$password= .\Get-ConfigurationPropertyValue.ps1 Password1
$domain= .\Get-ConfigurationPropertyValue.ps1 Domain
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$requestUrl= .\Get-ConfigurationPropertyValue.ps1 TargetSiteCollectionUrl

$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

# sent scripts to server.
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
	 param(
      [string]$siteUrl,
      [bool]$isClaimsAuthentication
      )

	# SharePoint name:the first element is short name of SharePoint,second is display name in registry.
	$WindowsSharePointServices3     = "WindowsSharePointServices3","Microsoft Windows SharePoint Services 3.0"
	$SharePointServer2007           = "SharePointServer2007","Microsoft Office SharePoint Server 2007 "
	$SharePointFoundation2010       = "SharePointFoundation2010","Microsoft SharePoint Foundation 2010"
	$SharePointServer2010           = "SharePointServer2010","Microsoft SharePoint Server 2010"
	$SharePointFoundation2013       = "SharePointFoundation2013","Microsoft SharePoint Foundation 2013"
	$SharePointServer2013           = "SharePointServer2013","Microsoft SharePoint Server 2013 "        
	$SharePointServer2016           = "SharePointServer2016","Microsoft SharePoint Server 2016"        
	$SharePointServer2019           = "SharePointServer2019","Microsoft SharePoint Server 2019" 
	$SharePointServerSubscriptionEdition           = "SharePointServerSubscriptionEdition","Microsoft SharePoint Server Subscription Edition" 

	$SharePointVersion              = "Unknown Version"
	$keys = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
	$items = $keys | foreach-object {Get-ItemProperty $_.PsPath}    
	foreach ($item in $items)
	{
		if($item.DisplayName -eq $WindowsSharePointServices3[1])
		{
			$SharePointVersion = $WindowsSharePointServices3[0]
			foreach ($item in $items)
			{
				if($item.DisplayName -eq ($SharePointServer2007[1]))
				{
					$SharePointVersion = $SharePointServer2007[0]
					break
				}
			}
			break
		 }
		elseif($item.DisplayName -eq $SharePointFoundation2010[1])
		{
			$SharePointVersion = $SharePointFoundation2010[0]
			break
		}        
		elseif($item.DisplayName -eq $SharePointServer2010[1])
		{
			$SharePointVersion = $SharePointServer2010[0]
			break
		}        
		elseif($item.DisplayName -eq $SharePointFoundation2013[1])
		{
			$SharePointVersion = $SharePointFoundation2013[0]
			break
		}        
		elseif($item.DisplayName -eq $SharePointServer2013[1])
		{
			$SharePointVersion = $SharePointServer2013[0]
			break
		}
		elseif($item.DisplayName -eq $SharePointServer2016[1])
		{
			$SharePointVersion = $SharePointServer2016[0]
			break
		}
		elseif($item.DisplayName -eq $SharePointServer2019[1])
		{
			$SharePointVersion = $SharePointServer2019[0]
			break
		}
		elseif($item.DisplayName -eq $SharePointServerSubscriptionEdition[1])
		{
			$SharePointVersion = $SharePointServerSubscriptionEdition[0]
			break
		}
	}

	# If the SharePoint version is "Unknown", then return false
	if($SharePointVersion -eq $null -or $SharePointVersion -eq "Unknown Version")
	{
		return $false
	}
	
	# If the SharePoint version is SharePointFoundation2013 or SharePointServer2013, then need to update the authentication mode.
	if($SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0]  -or $SharePointVersion -eq $SharePointServer2016[0] -or $SharePointVersion -eq $SharePointServer2019[0] -or $SharePointVersion -eq $SharePointServerSubscriptionEdition[0])
	{
        # load assemblies.
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
  
		$Uri = new-object System.Uri($siteUrl)
		$webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($Uri)

		$webApp.UseClaimsAuthentication = $isClaimsAuthentication
		$webApp.Update()
		return $true
	}

	# All other versions, do not need to switch claims authentication mode.
	return $true
} -argumentlist $requestUrl, $isClaimsAuthentication

if(!$?)
{
	return $false
}

return $ret