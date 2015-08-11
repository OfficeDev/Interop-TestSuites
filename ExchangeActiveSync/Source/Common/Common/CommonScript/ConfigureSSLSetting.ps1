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