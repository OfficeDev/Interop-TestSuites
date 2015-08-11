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
	param([string] $url , [bool]$isEnable)

	$script:ErrorActionPreference = "Stop"

    #Load Assemblies Microsoft.SharePoint.dll PSPowerShell.dll
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
        
    #Get SharePoint List
    $site = new-object Microsoft.SharePoint.SPSite($url)
	
	if ($site -eq $null) {
      return $false
	}

	$web = $site.OpenWeb()
	$web.ParserEnabled = $isEnable
	$web.Update()
	return $true

} -argumentlist $siteUrl, $isEnable 

return $ret