#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"
$UserName = "$PtfPropDomain\$PtfPropUser1Name"
$credentialSecurePassword = ConvertTo-SecureString $PtfPropUser1Password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($UserName,$credentialSecurePassword)

$SutComputerName = $PtfPropSutComputerName
$Password = $PtfPropUser1Password

invoke-command -computername $SutComputerName -Credential $credential  -ScriptBLock{
param(
      [String]$SutComputerName,  # Indicates the server name
      [String]$UserName,      # Indicates the user has permission to create a session, formatted as "domain\user"
      [String]$Password          # Indicates the password of credentialUserName
      )

    $credentialSecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($UserName,$credentialSecurePassword)
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$SutComputerName/PowerShell/ -Credential $credential -Authentication Kerberos
    Import-PSSession $session

    # The Discovery Search Mailboxes are excluded since they are not returned in the result of the methods of the MS-OXNSPI protocol.
    $list = Get-Recipient -RecipientPreviewFilter {(Alias -ne $null) -and (HiddenFromAddressListsEnabled -ne $true)}
    $ret=[Convert]::ToUint32($list.Length)
    Remove-PSSession $session
    return $ret
}-ArgumentList $SutComputerName, $UserName, $Password

