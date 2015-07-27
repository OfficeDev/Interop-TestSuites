#---------------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#---------------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"

#---------------------------------------------------
# Purpose: Enable/Disable Microsoft-Server-ActiveSync SSL
#---------------------------------------------------
$securePassword= convertto-securestring $userPassword -asplaintext -force
$credential = new-object Management.Automation.PSCredential(($userDomain+"\"+$userName),$securePassword)
$EASWebSettingsObj = get-wmiobject -namespace "root/MicrosoftIISv2" -query "select * from IIsWebVirtualDirSetting where Name='W3SVC/1/ROOT/Microsoft-Server-ActiveSync'" -computer $serverComputerName -Credential $credential -EnableAllPrivileges -Authentication PacketPrivacy
$EASWebSettingsObj.AccessSSL = $enableSSL
$EASWebSettingsObj.Put()

#---------------------------------------------------
# Check whether the SSL is configured
#---------------------------------------------------
$retryCount = $ptfpropRetryCount

do
{ 
  Start-Sleep -Milliseconds $ptfpropWaitTime
  $EASWebSettingsObj = get-wmiobject -namespace "root/MicrosoftIISv2" -query "select * from IIsWebVirtualDirSetting where Name='W3SVC/1/ROOT/Microsoft-Server-ActiveSync'" -computer $serverComputerName -Credential $credential -EnableAllPrivileges -Authentication PacketPrivacy
  $checkSSLStatus = $EASWebSettingsObj.AccessSSL
  $retryCount = $retryCount -1 
}
while($checkSSLStatus -ne $enableSSL -and $retryCount -gt 0)

if($checkSSLStatus -eq $enableSSL)
{
	return $true
}
else
{
	return $false
}