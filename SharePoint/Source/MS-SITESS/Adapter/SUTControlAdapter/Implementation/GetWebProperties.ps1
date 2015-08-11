$script:ErrorActionPreference = "Stop"
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$siteCollectionUrl = .\Get-ConfigurationPropertyValue.ps1 SiteCollectionUrl
$mainUrl = "http://" + $computerName

$language=.\Get-ConfigurationPropertyValue.ps1 SubSitePropertyLanguage
$locale = .\Get-ConfigurationPropertyValue.ps1 SubSitePropertyLocale
$currentUser = .\Get-ConfigurationPropertyValue.ps1 SubSitePropertyCurrentUser
$permissions = .\Get-ConfigurationPropertyValue.ps1 SubSitePropertyUserNameInPermissions
$defaultLanguage = .\Get-ConfigurationPropertyValue.ps1 SubSitePropertyDefaultLanguage
$anonymous = .\Get-ConfigurationPropertyValue.ps1 SubSitePropertyAnonymous

$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{

param(
    [string]$siteName,
    [string]$webName,
    [string]$mainUrl,
    [string]$siteCollectionUrl,
    [string]$language,
    [string]$locale,
    [string]$currentUser,
    [string]$permissions,
    [string]$defaultLanguage,
    [string]$anonymous
)
    $script:ErrorActionPreference = "Stop"
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

    $ret = $null

    try
    {
        if([string]::IsNullOrEmpty($siteCollectionUrl))
        {
            $spSite = new-object Microsoft.SharePoint.SPSite($mainUrl)
        }
        else
        {
            $spSite = new-object Microsoft.SharePoint.SPSite($siteCollectionUrl)
        }

        if(![string]::IsNullOrEmpty($webName) -and ![string]::IsNullOrEmpty($siteName))
		{
			$spWeb = $spSite.OpenWeb($siteName +"/"+ $webName)
		}
        elseif(![string]::IsNullOrEmpty($siteName))
		{
			$spWeb=$spSite.OpenWeb($siteName)
		}
        else
		{
			$spWeb =  $spSite.RootWeb
		}

        $ret = $language + ":" + $spWeb.Language.tostring() + ";"
        $ret += $locale + ":" + $spWeb.Locale.LCID.tostring() + ";"
        $ret += $currentUser + ":" + $spWeb.CurrentUser.Name + ";"
		$currentSPUser = $spWeb.CurrentUser
        $permissionStr = ""

		if($spWeb.Permissions.count -eq 1)
		{
			$permissionSPUser = $spWeb.Permissions[0].member
			if($currentSPUser.UserToken.CompareUser($permissionSPUser.UserToken))
			{
				$permissionStr = $permissions + ":" + $currentSPUser.Name + ";"
			}
			else
			{
				$permissionStr = $permissions + ":" + $permissionSPUser.Name + ";"
			}
		}
		else
		{
			$permissionStr = $permissions + ":"
			foreach($permission in $spWeb.Permissions)
			{
				$permissionStr += $permission.member.Name + ","
			}
			
            $permissionStr = $permissionStr.trim(',');
            $permissionStr += ";"
		}

		$ret += $permissionStr
		$ret += $defaultLanguage + ":" + [Microsoft.SharePoint.SPRegionalSettings]::GlobalInstalledLanguages[0].LCID + ";"
		if($spWeb.AnonymousState.ToString().equals("Disabled"))
		{
			$ret += $anonymous + ":false;" 
		}
		else
		{
			$ret += $anonymous + ":true;" 
		}
		$ret = $ret.trim(';')
    }
    finally
    {
        if ($spWeb -ne $null)
        {
            $spWeb.Close()
            $spWeb.Dispose()
        }
        if ($spSite -ne $null)
        {
            $spSite.Close()
            $spSite.Dispose()
        }
    }

    return $ret
}-argumentlist $siteName, $webName, $mainUrl, $siteCollectionUrl, $language, $locale, $currentUser, $permissions, $defaultLanguage, $anonymous

return $ret