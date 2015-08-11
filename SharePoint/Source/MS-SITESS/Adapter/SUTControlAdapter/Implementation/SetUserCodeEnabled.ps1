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
    [bool]$enable,
    [string]$mainUrl,
    [string]$siteCollectionUrl,
	[string]$computerName
)
    $script:ErrorActionPreference = "Stop"
    # Load SharePoint Snapin
    $snapin = Get-PSSnapin|Where-Object -FilterScript {$_.Name -eq "Microsoft.SharePoint.PowerShell"}
    if($snapin -eq $null)
    {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }

    # check the User Code Service's status, if the service disabled, enable it.
	$spServiceInstance = Get-SPServiceInstance | where {$_.TypeName -eq "Microsoft SharePoint Foundation Sandboxed Code Service" -and $_.Server.Address -eq $computerName}
	$isUpdated = $false
	if($spServiceInstance -eq "" -or $spServiceInstance -eq $null)
	{
		Throw "Can not find the Microsoft SharePoint Foundation Sandboxed Code Service."
	}
	else
	{
		if($spServiceInstance.Status -ne "Online")
			{
				$spServiceInstance.Provision()
				$spServiceInstance.Update()
				$isUpdated = $true
			}
		if($isUpdated)
		{
			Start-Sleep -s 15
			$spServiceInstance = Get-SPServiceInstance | where {$_.TypeName -eq "Microsoft SharePoint Foundation Sandboxed Code Service" -and $_.Server.Address -eq $computerName}
			if($spServiceInstance.Status -ne "Online")
			{
				Throw "Failed to start the User Code Service in 15 seconds."
			}
		}
	}

    try
    {
        if([string]::IsNullOrEmpty($siteCollectionUrl))
        {
            $spSite = Get-SPSite -Identity $mainUrl
        }
        else
        {
            $spSite = Get-SPSite -Identity $siteCollectionUrl
        }

        if($enable)
        {
            $spSite.Quota.UserCodeMaximumLevel = 300
            $spSite.Quota.UserCodeWarningLevel = 100
        }
        else 
        {
            $spSite.Quota.UserCodeWarningLevel = 0
            $spSite.Quota.UserCodeMaximumLevel = 0
        }

        # Get the reference of the spSite again because the data was refreshed.
        if ($spSite -ne $null)
        {
            $spSite.Close()
            $spSite.Dispose()
        }

        if([string]::IsNullOrEmpty($siteCollectionUrl))
        {
            $spSite = Get-SPSite -Identity $mainUrl
        }
        else
        {
            $spSite = Get-SPSite -Identity $siteCollectionUrl
        }

        $spSiteUserCodeEnabled=$spSite.UserCodeEnabled
    }
    finally
    {
        if ($spSite -ne $null)
        {
            $spSite.Close()
            $spSite.Dispose()
        }
    }
    return $spSiteUserCodeEnabled
}-argumentlist $enable, $mainUrl, $siteCollectionUrl, $computerName

return $ret