#-------------------------------------------------------------------------
# Copyright (c) 2015 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

[String]$containerPath = Get-Location
$logPath        = $containerPath + "\SetupLogs"
$logFile        = $logPath + "\ExchangeClientConfiguration.ps1.log"
$debugLogFile   = $logPath + "\ExchangeClientConfiguration.ps1.debug.log"

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
# Import the Common Function Library File
#-----------------------------------------------------
$scriptDirectory = Split-Path $MyInvocation.Mycommand.Path 
$commonScriptDirectory = $scriptDirectory.SubString(0,$scriptDirectory.LastIndexOf("\")+1) +"Common"       
.(Join-Path $commonScriptDirectory CommonConfiguration.ps1)
.(Join-Path $commonScriptDirectory ExchangeCommonConfiguration.ps1)

AddTimesStampsToLogFile "Start" "$logFile"
$environmentResourceFile            = "$commonScriptDirectory\ExchangeTestSuite.config"

#----------------------------------------------------------------------------
# Default Variables for Configuration 
#----------------------------------------------------------------------------
$MSOXWSATTUser01              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSATTUser01"
$MSOXWSATTUser01Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSATTUser01Password"

$MSOXWSBTRFUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSBTRFUser01"
$MSOXWSBTRFUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSBTRFUser01Password"

$MSOXWSCONTUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCONTUser01"
$MSOXWSCONTUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCONTUser01Password"

$MSOXWSCOREUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser01"
$MSOXWSCOREUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser01Password"
$MSOXWSCOREUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser02"
$MSOXWSCOREUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser02Password"
$MSOXWSCOREPublicFolder       = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREPublicFolder"

$MSOXWSFOLDUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser01"
$MSOXWSFOLDUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser01Password"
$MSOXWSFOLDUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser02"
$MSOXWSFOLDUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser02Password"

$MSOXWSMSGUser01              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser01"
$MSOXWSMSGUser01Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser01Password"
$MSOXWSMSGUser02              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser02"
$MSOXWSMSGUser02Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser02Password"
$MSOXWSMSGUser03              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser03"
$MSOXWSMSGUser03Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGUser03Password"
$MSOXWSMSGRoom01              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGRoom01"
$MSOXWSMSGRoom01Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMSGRoom01Password"

$MSOXWSMTGSUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser01"
$MSOXWSMTGSUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser01Password"
$MSOXWSMTGSUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser02"
$MSOXWSMTGSUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser02Password"
$MSOXWSMTGSRoom01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSRoom01"
$MSOXWSMTGSRoom01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSRoom01Password"

$MSOXWSSYNCUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser01"
$MSOXWSSYNCUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser01Password"
$MSOXWSSYNCUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser02"
$MSOXWSSYNCUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser02Password"

$MSOXWSTASKUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSTASKUser01"
$MSOXWSTASKUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSTASKUser01Password"

$MSOXWSFOLDManagedFolder1    = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDManagedFolder1"
$MSOXWSFOLDManagedFolder2    = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDManagedFolder2"

#-----------------------------------------------------
# Paths for all PTF configuration files.
#-----------------------------------------------------
$commonDeploymentFile      = resolve-path "..\..\Source\Common\ExchangeCommonConfiguration.deployment.ptfconfig"
$MSOXWSATTDeploymentFile   = resolve-path "..\..\Source\MS-OXWSATT\TestSuite\MS-OXWSATT_TestSuite.deployment.ptfconfig"
$MSOXWSBTRFDeploymentFile  = resolve-path "..\..\Source\MS-OXWSBTRF\TestSuite\MS-OXWSBTRF_TestSuite.deployment.ptfconfig"
$MSOXWSCONTDeploymentFile  = resolve-path "..\..\Source\MS-OXWSCONT\TestSuite\MS-OXWSCONT_TestSuite.deployment.ptfconfig"
$MSOXWSCOREDeploymentFile  = resolve-path "..\..\Source\MS-OXWSCORE\TestSuite\MS-OXWSCORE_TestSuite.deployment.ptfconfig"
$MSOXWSFOLDDeploymentFile  = resolve-path "..\..\Source\MS-OXWSFOLD\TestSuite\MS-OXWSFOLD_TestSuite.deployment.ptfconfig"
$MSOXWSMSGDeploymentFile   = resolve-path "..\..\Source\MS-OXWSMSG\TestSuite\MS-OXWSMSG_TestSuite.deployment.ptfconfig"
$MSOXWSMTGSDeploymentFile  = resolve-path "..\..\Source\MS-OXWSMTGS\TestSuite\MS-OXWSMTGS_TestSuite.deployment.ptfconfig"
$MSOXWSSYNCDeploymentFile  = resolve-path "..\..\Source\MS-OXWSSYNC\TestSuite\MS-OXWSSYNC_TestSuite.deployment.ptfconfig"
$MSOXWSTASKDeploymentFile  = resolve-path "..\..\Source\MS-OXWSTASK\TestSuite\MS-OXWSTASK_TestSuite.deployment.ptfconfig"

#-----------------------------------------------------
# Check and make sure that the SUT configuration is finished before running the client setup script.
#-----------------------------------------------------
OutputQuestion "The SUT must be configured before running the client setup script."
OutputQuestion "Did you either run the SUT setup script or configure the SUT as described by the Test Suite Deployment Guide? (Y/N)"
$isSutConfiguredChoices = @("Y","N")
$isSutConfigured = ReadUserChoice $isSutConfiguredChoices
if($isSutConfigured -eq "N")
{
    OutputText "You input `"N`"."
    OutputWarning "Exiting the client setup script now."
    OutputWarning "Configure the SUT and run the client setup script again."
    Stop-Transcript
    exit 0
}

#-----------------------------------------------------
# Check the Operating System (OS) version
#-----------------------------------------------------
OutputText "Check the Operating System (OS) version of the local machine ..."
CheckOSVersion -computer localhost

#-----------------------------------------------------
# Check the Application environment.
#-----------------------------------------------------
OutputText "Check whether the required applications have been installed ..."
$vsInstalledStatus = CheckVSVersion "12.0"
$ptfInstalledStatus = CheckPTFVersion "1.0.2220.0"
if(!$vsInstalledStatus -or !$ptfInstalledStatus)
{
    OutputQuestion "Would you like to continue without installing the application(s) or exit and install the application(s) (highlighted in yellow above)?"
    OutputQuestion "1: CONTINUE (Without installing the recommended application(s) , it may cause some risk on running the test cases)."
    OutputQuestion "2: EXIT."
    $runWithoutRequiredAppInstalledChoices = @('1','2')
    $runWithoutRequiredAppInstalled = ReadUserChoice $runWithoutRequiredAppInstalledChoices
    if($runWithoutRequiredAppInstalled -eq "2")
    {
        Stop-Transcript
        exit 0
    }
}

#-----------------------------------------------------
# Configuration for common ptfconfig file.
#-----------------------------------------------------
OutputText "Configure the ExchangeCommonConfiguration.deployment.ptfconfig file ..."
OutputQuestion "Enter the computer name of the SUT:"
OutputWarning "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported."
$sutcomputerName = ReadComputerName 
OutputText "The computer name of SUT you entered: $sutcomputerName"

OutputQuestion "Enter the domain name of SUT(for example: contoso.com):"
$domainInVM = Read-Host
OutputText "The domain name you entered: $domainInVM"

OutputQuestion "Select the Exchange Server version"
OutputQuestion "1: Microsoft Exchange Server 2007"
OutputQuestion "2: Microsoft Exchange Server 2010"
OutputQuestion "3: Microsoft Exchange Server 2013"

While (($serverVersion -eq $null) -or ($serverVersion -eq ""))
{ 
    [String]$version = Read-Host   
    Switch ($version)
    {
        "1" { $serverVersion = "ExchangeServer2007"; break }
        "2" { $serverVersion = "ExchangeServer2010"; break }
        "3" { $serverVersion = "ExchangeServer2013"; break }
        default {OutputWarning "Your input is invalid, select the Exchange Server version again"}
    }
}
OutputSuccess "The Exchange server version installed on the server is $serverVersion."

OutputQuestion "Select the transport type"
OutputQuestion "1: HTTP"
OutputQuestion "2: HTTPS"

While (($transportType -eq $null) -or ($transportType -eq ""))
{ 
    [string]$type = Read-Host
    Switch ($type)
    {
        "1" { $transportType = "HTTP";  break }
        "2" { $transportType = "HTTPS"; break }
        default {OutputWarning "Your input is invalid, select the transport type again"}
    }
}
OutputText "The transportType you entered: $transportType"
 
OutputText "Configure the ExchangeCommonConfiguration.deployment.ptfconfig file ..."

#Append to the url of Exchange Web Service.
$appendingURL ="EWS/Exchange.asmx"
#The URL which links to the service.
$serviceUrl= "[TransportType]://[SutComputerName].[Domain]/[AppendingURL]"

OutputText "Modify the properties as necessary in the ExchangeCommonConfiguration.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $commonDeploymentFile"
$step++
OutputWarning "$step.Find the property `"SutComputerName`", and set the value as $sutcomputerName"
$step++
OutputWarning "$step.Find the property `"Domain`", and set the value as $domainInVM"
$step++
OutputWarning "$step.Find the property `"SutVersion`", and set the value as $serverVersion"
$step++
OutputWarning "$step.Find the property `"TransportType`", and set the value as $transportType"
$step++
OutputWarning "$step.Find the property `"ServiceUrl`", and set the value as $serviceUrl"
$step++
OutputWarning "$step.Find the property `"AppendingURL`", and set the value as $appendingURL"

ModifyConfigFileNode $commonDeploymentFile "SutComputerName"       $sutcomputerName
ModifyConfigFileNode $commonDeploymentFile "Domain"                $domainInVM
ModifyConfigFileNode $commonDeploymentFile "SutVersion"            $serverVersion
ModifyConfigFileNode $commonDeploymentFile "TransportType"         $transportType
ModifyConfigFileNode $commonDeploymentFile "ServiceUrl"            $serviceUrl
ModifyConfigFileNode $commonDeploymentFile "AppendingURL"          $appendingURL

OutputSuccess "Configuration for the ExchangeCommonConfiguration.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSATT_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSATT_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSATTDeploymentFile"
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSOXWSATTUser01"
$step++
OutputWarning "$step.Find the property `"UserPassword`", and set the value as $MSOXWSATTUser01Password"

ModifyConfigFileNode $MSOXWSATTDeploymentFile "UserName"      "$MSOXWSATTUser01"
ModifyConfigFileNode $MSOXWSATTDeploymentFile "UserPassword"  "$MSOXWSATTUser01Password"

OutputSuccess "Configuration for the MS-OXWSATT_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSBTRF_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSBTRF_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSBTRFDeploymentFile"
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSOXWSBTRFUser01"
$step++
OutputWarning "$step.Find the property `"UserPassword`", and set the value as $MSOXWSBTRFUser01Password"

ModifyConfigFileNode $MSOXWSBTRFDeploymentFile "UserName"      "$MSOXWSBTRFUser01"
ModifyConfigFileNode $MSOXWSBTRFDeploymentFile "UserPassword"  "$MSOXWSBTRFUser01Password"

OutputSuccess "Configuration for the MS-OXWSBTRF_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSCONT_TestSuite.deployment.ptfconfig  file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSCONT_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSCONTDeploymentFile"
$step++
OutputWarning "$step.Find the property `"ContactUserName`", and set the value as $MSOXWSCONTUser01"
$step++
OutputWarning "$step.Find the property `"ContactUserPassword`", and set the value as $MSOXWSCONTUser01Password"

ModifyConfigFileNode $MSOXWSCONTDeploymentFile "ContactUserName"      "$MSOXWSCONTUser01"
ModifyConfigFileNode $MSOXWSCONTDeploymentFile "ContactUserPassword"  "$MSOXWSCONTUser01Password"

OutputSuccess "Configuration for the MS-OXWSCONT_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSCORE_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSCORE_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSCOREDeploymentFile"
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSOXWSCOREUser01"
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $MSOXWSCOREUser01Password"
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSOXWSCOREUser02"
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $MSOXWSCOREUser02Password"
$step++
OutputWarning "$step.Find the property `"PublicFolderName`", and set the value as $MSOXWSCOREPublicFolder"

ModifyConfigFileNode $MSOXWSCOREDeploymentFile "User1Name"        "$MSOXWSCOREUser01"
ModifyConfigFileNode $MSOXWSCOREDeploymentFile "User1Password"    "$MSOXWSCOREUser01Password"
ModifyConfigFileNode $MSOXWSCOREDeploymentFile "User2Name"        "$MSOXWSCOREUser02"
ModifyConfigFileNode $MSOXWSCOREDeploymentFile "User2Password"    "$MSOXWSCOREUser02Password"
ModifyConfigFileNode $MSOXWSCOREDeploymentFile "PublicFolderName" "$MSOXWSCOREPublicFolder"

OutputSuccess "Configuration for the MS-OXWSCORE_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSFOLD_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSFOLD_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSFOLDDeploymentFile"
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSOXWSFOLDUser01"
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $MSOXWSFOLDUser01Password"
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSOXWSFOLDUser02"
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $MSOXWSFOLDUser02Password"
$step++
OutputWarning "$step.Find the property `"ManagedFolderName1`", and set the value as $MSOXWSFOLDManagedFolder1"
$step++
OutputWarning "$step.Find the property `"ManagedFolderName2`", and set the value as $MSOXWSFOLDManagedFolder2"

ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "User1Name"           "$MSOXWSFOLDUser01"
ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "User1Password"       "$MSOXWSFOLDUser01Password"
ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "User2Name"           "$MSOXWSFOLDUser02"
ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "User2Password"       "$MSOXWSFOLDUser02Password"
ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "ManagedFolderName1"  "$MSOXWSFOLDManagedFolder1"
ModifyConfigFileNode $MSOXWSFOLDDeploymentFile "ManagedFolderName2"  "$MSOXWSFOLDManagedFolder2"

OutputSuccess "Configuration for the MS-OXWSFOLD_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSMSG_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSMSG_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSMSGDeploymentFile"
$step++
OutputWarning "$step.Find the property `"Sender`", and set the value as $MSOXWSMSGUser01"
$step++
OutputWarning "$step.Find the property `"SenderPassword`", and set the value as $MSOXWSMSGUser01Password"
$step++
OutputWarning "$step.Find the property `"Recipient1`", and set the value as $MSOXWSMSGUser02"
$step++
OutputWarning "$step.Find the property `"Recipient1Password`", and set the value as $MSOXWSMSGUser02Password"
$step++
OutputWarning "$step.Find the property `"Recipient2`", and set the value as $MSOXWSMSGUser03"
$step++
OutputWarning "$step.Find the property `"Recipient2Password`", and set the value as $MSOXWSMSGUser03Password"
$step++
OutputWarning "$step.Find the property `"RoomName`", and set the value as $MSOXWSMSGRoom01"

ModifyConfigFileNode $MSOXWSMSGDeploymentFile "Sender"              "$MSOXWSMSGUser01"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "SenderPassword"      "$MSOXWSMSGUser01Password"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "Recipient1"          "$MSOXWSMSGUser02"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "Recipient1Password"  "$MSOXWSMSGUser02Password"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "Recipient2"          "$MSOXWSMSGUser03"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "Recipient2Password"  "$MSOXWSMSGUser03Password"
ModifyConfigFileNode $MSOXWSMSGDeploymentFile "RoomName"            "$MSOXWSMSGRoom01"

OutputSuccess "Configuration for the MS-OXWSMSG_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSMTGS_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSMTGS_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSMTGSDeploymentFile"
$step++
OutputWarning "$step.Find the property `"OrganizerName`", and set the value as $MSOXWSMTGSUser01"
$step++
OutputWarning "$step.Find the property `"OrganizerPassword`", and set the value as $MSOXWSMTGSUser01Password"
$step++
OutputWarning "$step.Find the property `"AttendeeName`", and set the value as $MSOXWSMTGSUser02"
$step++
OutputWarning "$step.Find the property `"AttendeePassword`", and set the value as $MSOXWSMTGSUser02Password"
$step++
OutputWarning "$step.Find the property `"RoomName`", and set the value as $MSOXWSMTGSRoom01"

ModifyConfigFileNode $MSOXWSMTGSDeploymentFile "OrganizerName"       $MSOXWSMTGSUser01
ModifyConfigFileNode $MSOXWSMTGSDeploymentFile "OrganizerPassword"   $MSOXWSMTGSUser01Password
ModifyConfigFileNode $MSOXWSMTGSDeploymentFile "AttendeeName"        $MSOXWSMTGSUser02
ModifyConfigFileNode $MSOXWSMTGSDeploymentFile "AttendeePassword"    $MSOXWSMTGSUser02Password
ModifyConfigFileNode $MSOXWSMTGSDeploymentFile "RoomName"            $MSOXWSMTGSRoom01

OutputSuccess "Configuration for the MS-OXWSMTGS_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSSYNC_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSSYNC_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSSYNCDeploymentFile"
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSOXWSSYNCUser01"
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $MSOXWSSYNCUser01Password"
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSOXWSSYNCUser02"
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $MSOXWSSYNCUser02Password"

ModifyConfigFileNode $MSOXWSSYNCDeploymentFile "User1Name"         $MSOXWSSYNCUser01
ModifyConfigFileNode $MSOXWSSYNCDeploymentFile "User1Password"     $MSOXWSSYNCUser01Password
ModifyConfigFileNode $MSOXWSSYNCDeploymentFile "User2Name"         $MSOXWSSYNCUser02
ModifyConfigFileNode $MSOXWSSYNCDeploymentFile "User2Password"     $MSOXWSSYNCUser02Password

OutputSuccess "Configuration for the MS-OXWSSYNC_TestSuite.deployment.ptfconfig file is complete"

#-------------------------------------------------------
# Configuration for the MS-OXWSTASK_TestSuite.deployment.ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-OXWSTASK_TestSuite.deployment.ptfconfig file..."
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSOXWSTASKDeploymentFile"
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSOXWSTASKUser01"
$step++
OutputWarning "$step.Find the property `"UserPassword`", and set the value as $MSOXWSTASKUser01Password"

ModifyConfigFileNode $MSOXWSTASKDeploymentFile "UserName"         $MSOXWSTASKUser01
ModifyConfigFileNode $MSOXWSTASKDeploymentFile "UserPassword"     $MSOXWSTASKUser01Password

OutputSuccess "Configuration for the MS-OXWSTASK_TestSuite.deployment.ptfconfig file is complete"


#----------------------------------------------------------------------------
# End script
#----------------------------------------------------------------------------

OutputSuccess "[ExchangeClientConfiguration.PS1] has run successfully."
Stop-Transcript
exit 0