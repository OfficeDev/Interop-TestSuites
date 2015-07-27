#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"

$adminDomain = $userDomain
$adminName = $userName
$adminPassword = $userPassword
$sutVersion = $ptfpropSutVersion

$adminAccount = $adminName+"@"+$adminDomain

If($sutVersion -ge "ExchangeServer2010")
{
  $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
  $credential = new-object Management.Automation.PSCredential($adminAccount,$securePassword)

  #Invoke function remotely
  $result = invoke-command -ComputerName $serverComputerName -Credential $credential -scriptblock {
			param(
			[string]$adminAccount,
			[string]$adminPassword,
			[string]$serverComputerName
			)
			$connectUri="http://" + $serverComputerName + "/PowerShell"
			$securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
			$credential =new-object Management.Automation.PSCredential($adminAccount,$securePassword)
			$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectUri -Credential $credential -Authentication Kerberos
			Import-PSSession $session -AllowClobber -DisableNameChecking

			#Set mailbox folder access permission of User1
			$identity = $adminAccount+":\Calendar"
			$currentAccessRights = Get-MailboxFolderPermission $identity –User Default
			return $currentAccessRights.AccessRights
			
		}-ArgumentList $adminAccount,$adminPassword,$serverComputerName

	return $result
}