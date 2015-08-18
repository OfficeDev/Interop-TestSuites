$script:ErrorActionPreference = "Stop"
$credentialUserName = "$PtfPropDomain\$PtfPropAdminUserName"
$credentialSecurePassword = ConvertTo-SecureString $PtfPropAdminUserPassword -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($credentialUserName,$credentialSecurePassword)

invoke-command -computername $PtfPropSutComputerName -Credential $credential  -ScriptBLock{
param(
)
    $result=Get-WmiObject Win32_OperatingSystem | select Version,ServicePackMajorVersion,ServicePackMinorVersion
    $OsVersion=$result.Version+"."+$result.ServicePackMajorVersion+"."+$result.ServicePackMinorVersion

    return $OsVersion
}-ArgumentList $PtfPropSutComputerName