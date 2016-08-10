$script:ErrorActionPreference = "Stop"

$FolderPath = (Get-Location).Path

$serverName = $PtfPropSutComputerName
$serverVersion = $PtfPropSutVersion

$credentialUserName = "$PtfPropDomain\$PtfPropAdminUserName"
$credentialSecurePassword = ConvertTo-SecureString $PtfPropUserPassword -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($credentialUserName,$credentialSecurePassword)

#--------------------------------------------------------------------------------------------------------------------------------------
# Copy SUT control adapter scripts to server
#--------------------------------------------------------------------------------------------------------------------------------------
$ServerScriptPath = invoke-command -computername $serverName -Credential $credential -ScriptBLock{
    
	$RegScriptFolderKeyPath = "HKLM:\SOFTWARE\Microsoft\ExchangeTestSuite"
    $RegScriptFolderKeyName = "SUTAdapterScriptFolder"
    return (Get-ItemProperty -Path $RegScriptFolderKeyPath).$RegScriptFolderKeyName
}

$ServerScriptUncPath = "\\" + $serverName + "\" + $ServerScriptPath.Replace(":", "$")

$index = $ServerScriptUncPath.IndexOf('$') + 1
$share = $ServerScriptUncPath.substring(0,$index)

$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}
$disableMailboxFileName = "InServer-DisableMailbox" + "-" + $PtfPropUserName + "-" + $timestamp + ".ps1"
$addExchangeSnapInFileName = "AddExchangeSnapIn" + "-" + $PtfPropUserName + "-" + $timestamp + ".ps1"

net use $share $PtfPropUserPassword /USER:$credentialUserName

Copy-Item  "$FolderPath\InServer-DisableMailbox.ps1" -Destination "$ServerScriptUncPath\$disableMailboxFileName" -Force
Copy-Item  "$FolderPath\AddExchangeSnapIn.ps1" -Destination "$ServerScriptUncPath\$addExchangeSnapInFileName" -Force

net use $share /d /y

#--------------------------------------------------------------------------------------------------------------------------------------
# This script is used to call a script in server to disable a user's mailbox
#--------------------------------------------------------------------------------------------------------------------------------------
invoke-command -computername $serverName -Credential $credential  -ScriptBLock{
param(
    [String]$userName,          # Indicates the user name
    [String]$serverName,        # Indicates the exchange server name
    [String]$serverVersion,     # Indicates the exchange server version
    [String]$ServerScriptPath,	# Indicates the path of script on server
    [String]$credentialUserName,            # Indicate the user has the permission to create a session, formatted as "domain\user"
    [String]$PtfPropUserPassword,           # Indicate the password of credentialUserName
	[String]$disableMailboxFileName,        # Indicate the ps file to disable the mailbox
	[String]$addExchangeSnapInFileName      # Indicate the ps file to add Exchange snap-in
)

#----------------------------------------------------------------------------
# Verify required parameters
#----------------------------------------------------------------------------
if ($userName -eq $null -or $userName -eq "")
{
    Throw "Parameter userName is required."
}

$result = ""
$Error.Clear()

#----------------------------------------------------------------------------
# Remotely run a script in server to disable mailbox
#----------------------------------------------------------------------------
$psFile = "$ServerScriptPath\$disableMailboxFileName"
Powershell -File $psFile $userName $serverVersion $ServerScriptPath $serverName $credentialUserName $PtfPropUserPassword $addExchangeSnapInFileName

try
{
	Remove-Item -Path "$ServerScriptPath\$disableMailboxFileName"
	Remove-Item -Path "$ServerScriptPath\$addExchangeSnapInFileName"
}
catch [System.Exception]
{

}

if($Error.Count -ne 0)
{
	$ErrorNum=$Error.Count-1
	while($ErrorNum -ge 0)
	{
		$result += $Error[$ErrorNum].Exception.Message
		$ErrorNum--
	}	
}
else
{
	$result = "success"
}
return $result
}-ArgumentList $userName,$serverName,$serverVersion,$ServerScriptPath,$credentialUserName,$PtfPropUserPassword,$disableMailboxFileName,$addExchangeSnapInFileName