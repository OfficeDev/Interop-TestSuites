param(
[String]$userName,                 # Indicate the user name
[String]$serverVersion,            # Indicate the exchange server version
[String]$scriptPath,               # Indicate the scripts file path
[String]$serverName,               # Indicate the server name
[String]$credentialUserName,       # Indicate the user has the permission to create a session, formatted as "domain\user"
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

#----------------------------------------------------------------------------
# Disable mailbox
#----------------------------------------------------------------------------
Disable-Mailbox -Identity $userName -Confirm:$False

if($serverVersion -ne "ExchangeServer2007")
{
	Remove-PSSession $Session
}