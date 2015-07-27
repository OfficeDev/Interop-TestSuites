#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

param(
[String]$userName,                # Indicate the user name
[String]$serverVersion,           # Indicate the exchange server version
[String]$scriptPath,              # Indicate the scripts file path
[String]$serverName,              # Indicate the server name
[String]$credentialUserName,      # Indicate the user has the permission to create a session, formatted as "domain\user"
[String]$password,                  # Indicate the password of credentialUserName
[String]$addExchangeSnapInFileName  # Indicate the ps file to add Exchange snap-in
)

$script:ErrorActionPreference = "Stop"

#----------------------------------------------------------------------------
# Verify required parameters
#----------------------------------------------------------------------------
if ($userName -eq $null -or $userName -eq "")
{
    Throw "Parameter userName is required."
}

if ($serverVersion -eq $null -or $serverVersion -eq "")
{
    Throw "Parameter serverVersion is required."
}

if ($scriptPath -eq $null -or $scriptPath -eq "")
{
    Throw "Parameter scriptPath is required."
}

if ($serverName -eq $null -or $serverName -eq "")
{
    Throw "Parameter serverName is required."
}

if ($credentialUserName -eq $null -or $credentialUserName -eq "")
{
    Throw "Parameter credentialUserName is required."
}

if ($password -eq $null -or $password -eq "")
{
    Throw "Parameter password is required."
}

#----------------------------------------------------------------------------
# Add Exchange Snapin so that Exchange shell command could be executed
#----------------------------------------------------------------------------
if($serverVersion -eq "ExchangeServer2007")
{
	& $scriptPath\$addExchangeSnapInFileName $serverVersion
}
else
{
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($credentialUserName,$securePassword)
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$serverName/PowerShell/" -Credential $credential -Authentication Kerberos
	Import-PSSession $Session
}

$databaseName = ""

$databaseObject = Get-MailboxDatabase
$serverNameWithoutDomain = $serverName.Split(".")[0]
# There are more than 1 mailbox databases
if($databaseObject.Count -gt 1)
{
	$databaseNum= 0
	while($databaseNum -lt $databaseObject.Count)
	{
		if($databaseObject[$databaseNum].Server.ToString().ToLower() -eq $serverNameWithoutDomain.ToLower())
		{
			$databaseName = $databaseObject[$databaseNum].Name
			break
		}

		$databaseNum++
	}	
}
else
{
	# There is just 1 mailbox database
	$databaseName = $databaseObject.Name
}

if($databaseName -eq $null -or $databaseName -eq "")
{
	Throw "The mailbox database on " + $serverName + " can't be found!"
}

#----------------------------------------------------------------------------
# Enable mailbox
#----------------------------------------------------------------------------
# The result of Enable-Mailbox contains an non-readable char 0x15, 
# $result can avoid writing the non-readable char to log file to cause the log writer close.
$result = Enable-Mailbox -Identity $userName -Database $databaseName -Confirm:$False

if($serverVersion -ne "ExchangeServer2007")
{
	Remove-PSSession $Session
}