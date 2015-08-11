$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName
$transportType = .\Get-ConfigurationPropertyValue.ps1 TransportType
$testClientName = .\Get-ConfigurationPropertyValue.ps1 TestClientName

$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetSiteCollectionUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
    param(
        [string]$testClientName
    )
    $script:ErrorActionPreference = "Stop"
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
    $ret = $false
	Set-spwopizone internal-http
	New-SPWOPIBinding -ServerName $testClientName -allowhttp
	$ret = $true
	return $ret
} -argumentlist $testClientName

return $ret