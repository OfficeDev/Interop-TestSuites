#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

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