#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

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
    $ret=$false
    Remove-SPWOPIBinding -Confirm:$false -Server $testClientName
    $ret=$true
    return $ret
} -argumentlist $testClientName

return $ret
