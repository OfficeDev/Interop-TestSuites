$script:ErrorActionPreference = "Stop"
$SutComputerName = $PtfPropSutComputerName
$UserName = "$PtfPropDomain\$PtfPropUserName"
$Password = $PtfPropPassword
$credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($UserName,$credentialSecurePassword)


invoke-command -computername $SutComputerName -Credential $credential  -ScriptBLock{
param(
	[string]$SutComputerName,
	[string]$UserName,
	[string]$Password
      )

    $credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($UserName,$credentialSecurePassword)
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$SutComputerName/PowerShell/ -Credential $credential -Authentication Kerberos
    Import-PSSession $session

$result = Get-MailboxJunkEmailConfiguration MSOXWSCORE_User01 | select BlockedSendersAndDomains
[string]$BlockSender =$result.BlockedSendersAndDomains

    Remove-PSSession $session
return $BlockSender
} -ArgumentList $SutComputerName, $UserName, $Password
