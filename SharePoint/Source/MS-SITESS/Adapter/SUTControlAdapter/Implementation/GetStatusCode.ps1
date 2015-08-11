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
    [string]$listName,
    [string]$mainUrl,
    [string]$fileName,
    [string]$siteCollectionUrl
)
    $script:ErrorActionPreference = "Stop"
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

    $statusCode = ""

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

        $list = $spWeb.Lists[$listName] 
        $listDocItems = $list.Items

        Foreach ($listDocItem in $listDocItems)
        {
            if($listDocItem.Name -eq $fileName)
            {
                $sr = new-object System.IO.StreamReader($listDocItem.File.OpenBinaryStream());
                $statusCode = $sr.ReadLine();
                if($statusCode -eq $null)
                {
                    $statusCode = [string]::Empty;
                }
                $sr.Dispose()
                break
            }
        }
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

    return $statusCode
}-argumentlist $siteName, $webName, $listName, $mainUrl, $fileName, $siteCollectionUrl

return $ret