$script:ErrorActionPreference = "Stop"
$credentialUserName = "$PtfPropDomain\$PtfPropAdminUserName"
$credentialSecurePassword = ConvertTo-SecureString $PtfPropAdminUserPassword -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($credentialUserName,$credentialSecurePassword)

invoke-command -computername $PtfPropSutComputerName -Credential $credential  -ScriptBLock{
param(
 [String]$computerName
)
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName) 
    $regKey = $reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersSystem", $true) 
    $regkey.SetValue("Async Rpc Notify Enabled",0)

    $Service = Get-Service -Computer $computerName -Name MSExchangeIS
    Restart-Service -InputObject $Service

    exit
} -ArgumentList $PtfPropSutComputerName