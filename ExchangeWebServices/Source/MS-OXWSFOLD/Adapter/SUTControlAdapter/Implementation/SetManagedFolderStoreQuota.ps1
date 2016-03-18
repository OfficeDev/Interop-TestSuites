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
	[string]$ManagedFolderName
      )

    $credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($Admin,$credentialSecurePassword)
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$SutComputerName/PowerShell/ -Credential $credential -Authentication Kerberos
    Import-PSSession $session

[string] $result = Set-ManagedFolder -Identity $ManagedFolderName -StorageQuota 100KB

    Remove-PSSession $session
return $result
} -ArgumentList $SutComputerName, $Admin, $Password, $ManagedFolderName
