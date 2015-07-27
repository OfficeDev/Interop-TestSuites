#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

#----------------------------------------------------------------------------
# <param name="unattendedXmlName">The unattended SUT configuration XML.</param>
#----------------------------------------------------------------------------
param(
[string]$unattendedXmlName
)

#----------------------------------------------------------------------------
# Starting script
#----------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
[String]$containerPath = Get-Location
$logPath               = $containerPath + "\SetupLogs"
$logFile               = $logPath+"\ExchangeSUTConfiguration.ps1.log"
$debugLogFile          = $logPath+"\ExchangeSUTConfiguration.ps1.debug.log"
if(!(Test-Path $logPath))
{
    New-Item $logPath -ItemType directory
}
elseif(Test-Path $logFile)
{
    Remove-Item $logFile -Force
}
Start-Transcript $debugLogFile -Force

#-----------------------------------------------------
# Import the common function library file
#-----------------------------------------------------
$scriptDirectory = Split-Path $MyInvocation.Mycommand.Path 
$commonScriptDirectory = $scriptDirectory.SubString(0,$scriptDirectory.LastIndexOf("\")+1) +"Common"
.(Join-Path $commonScriptDirectory CommonConfiguration.ps1)
.(Join-Path $commonScriptDirectory ExchangeCommonConfiguration.ps1)

AddTimesStampsToLogFile "Start" "$logFile"

#----------------------------------------------------------------------------
# Values for configuration 
#----------------------------------------------------------------------------
$domain                             = $env:USERDNSDOMAIN
$sutComputerName                    = $env:ComputerName
$userName                           = $env:UserName 

$environmentResourceFile            = "$commonScriptDirectory\ExchangeTestSuite.config"

$defaultPublicFolderDatabase        = ReadConfigFileNode "$environmentResourceFile" "defaultPublicFolderDatabaseOnPrimarySUT"
$defaultPublicFolderMailboxPrefix   = ReadConfigFileNode "$environmentResourceFile" "defaultPublicFolderMailboxPrefixOnPrimarySUT"
$defaultPublicFolderMailbox         = $defaultPublicFolderMailboxPrefix + $sutComputerName

$MSOXCFOLDCommonUser                = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDCommonUser"
$MSOXCFOLDCommonUserPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDCommonUserPassword"
$MSOXCFOLDAdminUser                 = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDAdminUser"
$MSOXCFOLDAdminUserPassword         = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDAdminUserPassword"
$MSOXCFOLDPublicFolderMailEnabled   = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDPublicFolderMailEnabled"

$MSOXCFXICSAdminUser                = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSAdminUser"
$MSOXCFXICSAdminUserPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSAdminUserPassword"
$MSOXCFXICSGhostedPublicFolder      = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSGhostedPublicFolder"
$MSOXCFXICSPublicFolder             = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSPublicFolder"

$MSOXCMAPIHTTPAdminUser             = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPAdminUser"
$MSOXCMAPIHTTPAdminUserPassword     = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPAdminUserPassword"
$MSOXCMAPIHTTPGeneralUser           = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPGeneralUser"
$MSOXCMAPIHTTPGeneralUserPassword   = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPGeneralUserPassword"
$MSOXCMAPIHTTPDistributionGroup     = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPDistributionGroup"

$MSOXCMSGCommonUser                 = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGCommonUser"
$MSOXCMSGCommonUserPassword         = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGCommonUserPassword"
$MSOXCMSGAdminUser                  = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGAdminUser"
$MSOXCMSGAdminUserPassword          = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGAdminUserPassword"

$MSOXCNOTIFUser                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCNOTIFUser"
$MSOXCNOTIFUserPassword             = ReadConfigFileNode "$environmentResourceFile" "MSOXCNOTIFUserPassword"

$MSOXCPERMUser1                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser1"
$MSOXCPERMUser1Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser1Password"
$MSOXCPERMUser2                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser2"
$MSOXCPERMUser2Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser2Password"

$MSOXCPRPTUser                      = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTUser"
$MSOXCPRPTUserPassword              = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTUserPassword"
$MSOXCPRPTPublicFolder              = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTPublicFolder"

$MSOXCROPSUser                      = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSUser"
$MSOXCROPSUserPassword              = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSUserPassword"
$MSOXCROPSEmailAlias                = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSEmailAlias"
$MSOXCROPSEmailAliasPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSEmailAliasPassword"

$MSOXCRPCNormalUser                 = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCNormalUser"
$MSOXCRPCNormalUserPassword         = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCNormalUserPassword"
$MSOXCRPCAdminUser                  = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCAdminUser"
$MSOXCRPCAdminUserPassword          = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCAdminUserPassword"

$MSOXCSTORUser                      = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORUser"
$MSOXCSTORUserPassword              = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORUserPassword"
$MSOXCSTORMailboxOnServer1          = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer1"
$MSOXCSTORMailboxOnServer1Password  = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer1Password"
$MSOXCSTORDisableMailbox            = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORDisableMailbox"
$MSOXCSTORDisableMailboxPassword    = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORDisableMailboxPassword"

$MSOXCTABLSender1                   = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender1"
$MSOXCTABLSender1Password           = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender1Password"
$MSOXCTABLSender2                   = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender2"
$MSOXCTABLSender2Password           = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender2Password"

$MSOXNSPIUser1                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser1"
$MSOXNSPIUser1Password              = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser1Password"
$MSOXNSPIUser2                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser2"
$MSOXNSPIUser2Password              = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser2Password"
$MSOXNSPIUser3                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser3"
$MSOXNSPIUser3Password              = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser3Password"
$MSOXNSPIPublicFolderMailEnabled    = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIPublicFolderMailEnabled"
$MSOXNSPIDynamicDistributionGroup   = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIDynamicDistributionGroup"
$MSOXNSPIDistributionGroup          = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIDistributionGroup"
$MSOXNSPIMailContact                = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIMailContact"

$MSOXORULEUser1                     = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser1"
$MSOXORULEUser1Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser1Password"
$MSOXORULEUser2                     = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser2"
$MSOXORULEUser2Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser2Password"

$regScriptFolderKeyPath             = "HKLM:\SOFTWARE\Microsoft\ExchangeTestSuite"
$regScriptFolderKeyName             = "SUTAdapterScriptFolder"
$regScriptFolderKeyValue            = (Get-Location).Path + "\SUTAdapterScripts"

#-----------------------------------------------------
# Check whether the unattended SUT configuration XML is available if run in unattended mode.
#-----------------------------------------------------
if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
{    
    Output "The SUT setup script will run in attended mode." "White"
}
else
{    
    While($unattendedXmlName -ne "" -and $unattendedXmlName -ne $null)
    {   
        if(Test-Path $unattendedXmlName -PathType Leaf)
        {
            Output "The SUT setup script will run in unattended mode with information provided by the `"$unattendedXmlName`" file." "White"
            $unattendedXmlName = Resolve-Path $unattendedXmlName
            break
        }
        else
        {
            Output "The SUT configuration XML path `"$unattendedXmlName`" is not correct." "Yellow"
            Output "Retry with the correct file path or press `"Enter`" if you want the SUT setup script to run in attended mode." "Cyan"
            $unattendedXmlName = Read-Host
        }
    }
}

#-----------------------------------------------------
# Begin to configure server
#-----------------------------------------------------
Output "Begin to configure server ..." "White"
Output "Steps for manual configuration:" "Yellow" 
Output "Enable remoting in Powershell." "Yellow"
Invoke-Command {
    $ErrorActionPreference = "Continue"
    Enable-PSRemoting -Force
}

[int]$recommendedMaxMemory = 1024
Output "Ensure that the maximum amount of memory allocated per shell for remote shell management is at least $recommendedMaxMemory MB." "Yellow"
[int]$originalMaxMemory = (Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).Value
if($originalMaxMemory -lt $recommendedMaxMemory)
{
    Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB $recommendedMaxMemory
    $actualMaxMemory = (Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).Value
    Output "The maximum amount of memory allocated per shell for remote shell management is increased from $originalMaxMemory MB to $actualMaxMemory MB." "White"
}
else
{
    Output "The maximum amount of memory allocated per shell for remote shell management is $originalMaxMemory MB." "White"
}

InstallWindowsFeature RSAT-AD-PowerShell
StartService "msexchange*" "auto"

#-----------------------------------------------------
# Get Exchange server basic information
#-----------------------------------------------------
$global:ExchangeVersion = GetExchangeServerVersion
Output "The basic information of the main Exchange server:" "White"
Output "Domain name: $domain" "White"
Output "The name of the main Exchange server: $sutComputerName" "White"
Output "The version of the Exchange server: $global:ExchangeVersion" "White"
Output "The logon name of the current user: $userName " "White"

#-----------------------------------------------------
# Add Exchange PowerShell SnapIn 
#-----------------------------------------------------
AddExchangePSSnapIn

if(($global:ExchangeVersion -eq $global:Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    Set-AdminAuditLogConfig -AdminAuditLogEnabled $False -WarningAction SilentlyContinue
}

#-----------------------------------------------------
# Check whether Exchange server is installed on a domain controller.
#-----------------------------------------------------
CheckExchangeInstalledOnDCOrNot

#----------------------------------------------------------------------------
# Start configuration of RPC over HTTP for Exchange server
#----------------------------------------------------------------------------
ConfigureRPCOverHTTP
Output "The SUT is now configured for RPC over HTTP transport." "Green"

#----------------------------------------------------------------------------
# Start configuration of Mailbox users for Exchange server
#----------------------------------------------------------------------------
# Create Mailbox users
Output "Mailbox users are currently being created on the Exchange server; please wait..." "White"
$mailboxDatabases = Get-MailboxDatabase -Server $sutComputerName
if(@($mailboxDatabases).count -gt 1)
{
    $mailboxDatabaseName = $mailboxDatabases[0].Name
}
else
{
    $mailboxDatabaseName = $mailboxDatabases.Name
}

CreateMailboxUser  $MSOXCFOLDCommonUser       $MSOXCFOLDCommonUserPassword       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCFOLDAdminUser        $MSOXCFOLDAdminUserPassword        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCFXICSAdminUser       $MSOXCFXICSAdminUserPassword       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCMSGCommonUser        $MSOXCMSGCommonUserPassword        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCMSGAdminUser         $MSOXCMSGAdminUserPassword         $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCNOTIFUser            $MSOXCNOTIFUserPassword            $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCPERMUser1            $MSOXCPERMUser1Password            $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCPERMUser2            $MSOXCPERMUser2Password            $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCPRPTUser             $MSOXCPRPTUserPassword             $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCROPSEmailAlias       $MSOXCROPSEmailAliasPassword       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCROPSUser             $MSOXCROPSUserPassword             $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCRPCNormalUser        $MSOXCRPCNormalUserPassword        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCRPCAdminUser         $MSOXCRPCAdminUserPassword         $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCSTORMailboxOnServer1 $MSOXCSTORMailboxOnServer1Password $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCSTORDisableMailbox   $MSOXCSTORDisableMailboxPassword   $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCSTORUser             $MSOXCSTORUserPassword             $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCTABLSender1          $MSOXCTABLSender1Password          $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXCTABLSender2          $MSOXCTABLSender2Password          $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXORULEUser1            $MSOXORULEUser1Password            $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXORULEUser2            $MSOXORULEUser2Password            $mailboxDatabaseName $domain

if($global:ExchangeVersion -eq $global:Exchange2007)
{
    $orgAdminGroup = "OrgAdmin"
    $pfAdminGroup = "PublicFolderAdmin"  
}
elseif($global:ExchangeVersion -ge $global:Exchange2010)
{
    $orgAdminGroup = "Organization Management"
    $pfAdminGroup = "Public Folder Management"
}

#Add user to Exchange organization management group
AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCFOLDAdminUser  $orgAdminGroup
AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCMSGAdminUser   $orgAdminGroup
AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCRPCAdminUser   $orgAdminGroup
AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCSTORUser       $orgAdminGroup

if($global:ExchangeVersion -ge $global:Exchange2010)
{
    CreateMailboxUser  $MSOXNSPIUser1     $MSOXNSPIUser1Password     $mailboxDatabaseName $domain
    CreateMailboxUser  $MSOXNSPIUser2     $MSOXNSPIUser2Password     $mailboxDatabaseName $domain
    CreateMailboxUser  $MSOXNSPIUser3     $MSOXNSPIUser3Password     $mailboxDatabaseName $domain

    # Set user setting
    Set-User  -Identity  $MSOXNSPIUser1 -AssistantName "assistant" -PhoneticDisplayName "phoneticdisplayname"
    Set-User  -Identity  $MSOXNSPIUser2 -AssistantName "assistant" -PhoneticDisplayName "phoneticdisplayname" -Office "Test"  -Department "Test" -OtherHomePhone {"BusinessOne"}

    Output "Add a certificate to $MSOXNSPIUser1." "White"
    $certFile = "$env:SystemDrive" + "\CN=" + $sutComputerName + ".Cer" 
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate $certFile
    Import-Module ActiveDirectory
    Set-ADUser "$MSOXNSPIUser1" -Certificates @{Add=$cert}

    #Add user to Exchange organization management group
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXNSPIUser1  $orgAdminGroup

    #Create mail contacts
    $mailContact = Get-MailContact -Filter {Name -eq $MSOXNSPIMailContact}
    if($mailContact -eq $null -or $mailContact -eq "")
    {
        Output "Creating a mail contact $MSOXNSPIMailContact." "White"
        New-MailContact -Name $MSOXNSPIMailContact -ExternalEmailAddress "$MSOXNSPIMailContact@$domain" | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    }
    else
    {
        Output "$MSOXNSPIMailContact already exists." "Yellow"
    }

    # Create distribution groups
    $dynamicDistributionArray = Get-DynamicDistributionGroup -Filter {Name -eq $MSOXNSPIDynamicDistributionGroup}
    if($dynamicDistributionArray -eq $null -or $dynamicDistributionArray -eq "")
    {
        Output "Creating a dynamic distribution group $MSOXNSPIDynamicDistributionGroup." "White"
        New-DynamicDistributionGroup -Name $MSOXNSPIDynamicDistributionGroup -IncludedRecipients "AllRecipients" | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    }
    else
    {
        Output "$MSOXNSPIDynamicDistributionGroup already exists." "Yellow"
    }

    CreateDistrbutionGroup $MSOXNSPIDistributionGroup "Distribution" $MSOXNSPIDistributionGroup
    Output "Set $MSOXNSPIDistributionGroup to be managed by $MSOXNSPIUser1" "White"
    Set-DistributionGroup -Identity $MSOXNSPIDistributionGroup -ManagedBy $MSOXNSPIUser1
}

if($global:ExchangeVersion -ge $global:Exchange2013)
{
    CreateMailboxUser  $MSOXCMAPIHTTPAdminUser   $MSOXCMAPIHTTPAdminUserPassword   $mailboxDatabaseName $domain
    CreateMailboxUser  $MSOXCMAPIHTTPGeneralUser $MSOXCMAPIHTTPGeneralUserPassword $mailboxDatabaseName $domain

    # Add user to Exchange organization management group
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCMAPIHTTPAdminUser  $orgAdminGroup

    CreateDistrbutionGroup $MSOXCMAPIHTTPDistributionGroup "Distribution" $MSOXCMAPIHTTPDistributionGroup
    Output "Set $MSOXCMAPIHTTPDistributionGroup to be managed by $MSOXCMAPIHTTPAdminUser" "White"
    Set-DistributionGroup -Identity $MSOXCMAPIHTTPDistributionGroup -ManagedBy $MSOXCMAPIHTTPAdminUser
}

Output "All required mailbox users are created and configured!" "Green"

#----------------------------------------------------------------------------
# Start configuration of public folder databases and public folders for Exchange server
#----------------------------------------------------------------------------
$rootPublicFolderName = "\"
$OwnerAccessRights = "Owner"
if($global:ExchangeVersion -le $global:Exchange2010)
{
    # Create public folder database
    CreatePublicFolderDatabase $defaultPublicFolderDatabase $sutComputerName

    # Create public folders
    CreatePublicFolder $MSOXCFOLDPublicFolderMailEnabled $sutComputerName
    CreatePublicFolder $MSOXCFXICSGhostedPublicFolder    $sutComputerName
    CreatePublicFolder $MSOXCFXICSPublicFolder           $sutComputerName
    CreatePublicFolder $MSOXCPRPTPublicFolder            $sutComputerName
    if($global:ExchangeVersion -ge $global:Exchange2010)
    {
        CreatePublicFolder $MSOXNSPIPublicFolderMailEnabled  $sutComputerName
    }

    # Add user to Exchange public folder management group
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCFOLDAdminUser  $pfAdminGroup
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCFXICSAdminUser $pfAdminGroup
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCMSGAdminUser   $pfAdminGroup
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCPRPTUser       $pfAdminGroup
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXCROPSUser       $pfAdminGroup
    AddUserToExchangeAdminGroup $global:ExchangeVersion $MSOXORULEUser1      $pfAdminGroup
}
elseif($global:ExchangeVersion -ge $global:Exchange2013)
{
    Output "Public folder replication is not supported in $global:ExchangeVersion" "Yellow"
    $publicFolderMailboxName = CreatePublicFolderMailbox $defaultPublicFolderMailbox $sutComputerName $mailboxDatabaseName

    # Add user permission to public folder    
    AddUserPublicFolderClientPermission $userName            $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion 
    AddUserPublicFolderClientPermission $MSOXCFOLDAdminUser  $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXCFXICSAdminUser $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXCMSGAdminUser   $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXCPRPTUser       $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXCROPSUser       $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXCSTORUser       $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion
    AddUserPublicFolderClientPermission $MSOXORULEUser1      $rootPublicFolderName $OwnerAccessRights $global:ExchangeVersion

    # Create public folders
    CreatePublicFolder $MSOXCFOLDPublicFolderMailEnabled $sutComputerName $publicFolderMailboxName
    CreatePublicFolder $MSOXCFXICSPublicFolder           $sutComputerName $publicFolderMailboxName 
    CreatePublicFolder $MSOXCFXICSGhostedPublicFolder    $sutComputerName $publicFolderMailboxName
    CreatePublicFolder $MSOXCPRPTPublicFolder            $sutComputerName $publicFolderMailboxName
    CreatePublicFolder $MSOXNSPIPublicFolderMailEnabled  $sutComputerName $publicFolderMailboxName
}
AddUserPublicFolderClientPermission $MSOXCFXICSAdminUser  "$rootPublicFolderName$MSOXCFXICSPublicFolder"        $OwnerAccessRights $global:ExchangeVersion
AddUserPublicFolderClientPermission $MSOXCFXICSAdminUser  "$rootPublicFolderName$MSOXCFXICSGhostedPublicFolder" $OwnerAccessRights $global:ExchangeVersion
AddUserPublicFolderClientPermission $MSOXCPRPTUser        "$rootPublicFolderName$MSOXCPRPTPublicFolder"         $OwnerAccessRights $global:ExchangeVersion

EnableMailOnPublicFolder "\$MSOXCFOLDPublicFolderMailEnabled"
if($global:ExchangeVersion -ge $global:Exchange2010)
{
    EnableMailOnPublicFolder "\$MSOXNSPIPublicFolderMailEnabled"
}
Output "All required public folder databases and public folders are created and configured!" "Green"

#----------------------------------------------------------------------------
# Start other configurations for Exchange server
#----------------------------------------------------------------------------
#Disable encryption on the Exchange server
DisableEncryption $sutComputerName $global:ExchangeVersion

#Enable MAPIHTTP on the Exchange server
if($global:ExchangeVersion -ge $global:Exchange2013)
{
    Output "Enable MAPIHTTP on the Exchange server..." "White"
    Set-OrganizationConfig -MapiHttpEnabled $true
    IISReset /restart
    Output "Enabled MAPIHTTP on the Exchange server successfully." "Green"
}
# Configure ClientMonitoringEnableFlags in the registry
Output "Configure ClientMonitoringEnableFlags in the registry ..." "White" 
Output "Steps for manual configuration:" "Yellow"  
Output "In HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\MSExchangeIS\ParametersSystem, create a key named `"ClientMonitoringEnableFlags`" and set the key value to 8" "Yellow" 
cmd /c reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\MSExchangeIS\ParametersSystem" /v "ClientMonitoringEnableFlags" /t "REG_DWORD" /d 8 /f | Out-File -FilePath $logFile -Append -encoding ASCII -width 100

# Set the max extended rule size
Output "Set the max extended rule size..." "Yellow" 
New-ItemProperty -Path "HKLM:System\CurrentControlSet\Services\MSExchangeIS\ParametersSystem" -PropertyType DWORD -Name "Max Extended Rule Size" -Value "96" -Force | Out-File -FilePath $logFile -Append -encoding ASCII -width 100

# Configure EWS
Output "Configure EWS..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step. In the `"SSL Settings`" page of `"Default Web Site/EWS`" in IIS, select `"Require SSL`", and select `"Ignore`" for client certificates" "Yellow" 
$step++
Output "$step. Restart IIS" "Yellow"
cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/EWS" /commit:APPHOST /section:system.webServer/security/access /sslFlags:"Ssl,Ssl128" | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
IISReset /restart
Output "Configured EWS successfully." "Green"

# Set SUT Adapter Script path in the registry.
Output "Setting the SUT adapter script path in the the registry ..." "White" 
Output "Steps for manual configuration:" "Yellow"
Output "1: Create a folder ($regScriptFolderKeyValue) if it doesn't exist." "Yellow"
Output "2: Create a key ($regScriptFolderKeyPath) in the registry if it doesn't exist." "Yellow"
Output "3: Create a new string value under $regScriptFolderKeyPath in the registry." "Yellow"
Output "   ValueName:$regScriptFolderKeyName" "Yellow"
Output "   ValueType:String" "Yellow"
Output "   Value:$regScriptFolderKeyValue" "Yellow"
if(!(Test-Path $regScriptFolderKeyValue))
{
     New-Item $regScriptFolderKeyValue -ItemType directory | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 
}
if(!(Test-Path $regScriptFolderKeyPath))
{
     New-Item $regScriptFolderKeyPath -ItemType directory | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
}
New-ItemProperty -Path $regScriptFolderKeyPath -Name $regScriptFolderKeyName -Value $regScriptFolderKeyValue -Force | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
Output "Successfully set the SUT adapter script path in the registry editor." "Green"

if(($global:ExchangeVersion -eq $global:Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    Set-AdminAuditLogConfig -AdminAuditLogEnabled $true -WarningAction SilentlyContinue
}

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
Output "[ExchangeSecondSUTConfiguration.PS1] has run sucessfully." "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0