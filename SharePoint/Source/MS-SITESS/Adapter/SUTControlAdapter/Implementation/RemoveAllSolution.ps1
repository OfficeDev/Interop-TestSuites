$script:ErrorActionPreference = "Stop"
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$siteCollectionUrl = .\Get-ConfigurationPropertyValue.ps1 SiteCollectionUrl
$mainUrl = "http://" + $computerName

$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
param(
    [string]$siteName,
    [string]$webName,
    [string]$solutionGalleryName,
    [string]$mainUrl,
    [string]$siteCollectionUrl
)
    $script:ErrorActionPreference = "Stop"
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

    $result = $false
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

        $spUserSolutionCollection = [Microsoft.SharePoint.SPUserSolutionCollection]$spSite.Solutions;
        [int]$count = $spUserSolutionCollection.Count
        $spUserSolutions = new-object Microsoft.SharePoint.SPUserSolution[] $count
        $spUserSolutionCollection.CopyTo($spUserSolutions,0)

        for( $intIndex = 0; $intIndex -lt $count; $intIndex++)
        {
            $spUserSolutionCollection.Remove($spUserSolutions[$intIndex])
        }

        $list = $spWeb.Lists[$solutionGalleryName]
        $listDocItems = $list.Items
        [int]$count = $listDocItems.Count

        for( $intIndex = 0; $intIndex -lt $count; $intIndex++)
        {
            $listDocItems.Delete(0)
        }

        $result = $true
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

    return $result
}-argumentlist $siteName, $webName, $solutionGalleryName, $mainUrl, $siteCollectionUrl

return $ret