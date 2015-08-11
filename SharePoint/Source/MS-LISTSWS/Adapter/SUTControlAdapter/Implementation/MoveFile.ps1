$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
    param(
        [string]$sourceUrl,
        [string]$destinationUrl,
        [string]$requestUrl
    )
	$script:ErrorActionPreference = "Stop"
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
    try
    {

        $spSite = new-object Microsoft.SharePoint.SPSite "$requestUrl"
        $spWeb =  $spSite.RootWeb
        $sourceFile = $spWeb.GetFile($sourceUrl)
        if ($sourceFile.Exists)
        {
            $sourceFile.MoveTo($destinationUrl, $true)
        }
        else
        {
            throw "Failed to move file because source file '$sourceFile' is not found."
        }
    }
	catch
	{
	    throw $error[0]
	}
    finally
    {
       	if ($spSite -ne $null)
        {
            $spSite.Dispose()
        }
    }
} -argumentlist $sourceUrl, $destinationUrl, $requestUrl

exit 0