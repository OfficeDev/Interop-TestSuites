$script:ErrorActionPreference = "Stop"
$SutComputerName = $PtfPropSutComputerName
$Admin = "$PtfPropDomain\$PtfPropUserName"
$Password = $PtfPropPassword
$credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($Admin,$credentialSecurePassword)

invoke-command -computername $SutComputerName -Credential $credential  -ScriptBLock{
param(
	[string]$SutComputerName,
	[string]$Admin,
	[string]$Password,
	[string]$UserName
      )

    $credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($Admin,$credentialSecurePassword)
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$SutComputerName/PowerShell/ -Credential $credential -Authentication Kerberos
    Import-PSSession $session

$result = Get-MailboxJunkEmailConfiguration $UserName | select BlockedSendersAndDomains
[string]$BlockSender =$result.BlockedSendersAndDomains

    Remove-PSSession $session
return $BlockSender
} -ArgumentList $SutComputerName, $Admin, $Password, $UserName
