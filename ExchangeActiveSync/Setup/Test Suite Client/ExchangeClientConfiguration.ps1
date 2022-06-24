#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

##
#---------------------------------------------------------------------------------
# <param name="unattendedXmlName">The unattended client configuration XML.</param>
#---------------------------------------------------------------------------------
param(
[string]$unattendedXmlName
)

#----------------------------------------------------------------------------
# Start script
#----------------------------------------------------------------------------
[String]$containerPath = Get-Location
$logPath        = $containerPath + "\SetupLogs"
$logFile        = $logPath + "\ExchangeClientConfiguration.ps1.log"
$debugLogFile   = $logPath + "\ExchangeClientConfiguration.ps1.debug.log"

if(!(Test-Path $logPath))
{
    New-Item $logPath -ItemType directory
}
Start-Transcript $debugLogFile -force

#-----------------------------------------------------
# Import the Common Function Library File
#-----------------------------------------------------
$scriptDirectory = Split-Path $MyInvocation.Mycommand.Path 
$commonScriptDirectory = $scriptDirectory.SubString(0,$scriptDirectory.LastIndexOf("\")+1) +"Common"       
.(Join-Path $commonScriptDirectory CommonConfiguration.ps1)
.(Join-Path $commonScriptDirectory ExchangeCommonConfiguration.ps1)

If (Test-Path $logFile)
{
    Remove-Item $logFile -Force
}
AddTimesStampsToLogFile "Start" "$logFile"
$environmentResourceFile            = "$commonScriptDirectory\ExchangeTestSuite.config"

#----------------------------------------------------------------------------
# Default Variables for Configuration 
#----------------------------------------------------------------------------
$userPassword                        = ReadConfigFileNode "$environmentResourceFile" "userPassword"

$MSASAIRSUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASAIRSUser01"
$MSASAIRSUser02                      = ReadConfigFileNode "$environmentResourceFile" "MSASAIRSUser02"

$MSASCALUser01                       = ReadConfigFileNode "$environmentResourceFile" "MSASCALUser01"
$MSASCALUser02                       = ReadConfigFileNode "$environmentResourceFile" "MSASCALUser02"

$MSASCMDUser01                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser01"
$MSASCMDUser02                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser02"
$MSASCMDUser03                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser03"
$MSASCMDUser04                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser04"
$MSASCMDUser07                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser07"
$MSASCMDUser08                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser08"
$MSASCMDUser09                       = ReadConfigFileNode "$environmentResourceFile" "MSASCMDUser09"
$MSASCMDTestGroup                    = ReadConfigFileNode "$environmentResourceFile" "MSASCMDTestGroup"
$MSASCMDLargeGroup                   = ReadConfigFileNode "$environmentResourceFile" "MSASCMDLargeGroup"
$MSASCMDSharedFolder                 = ReadConfigFileNode "$environmentResourceFile" "MSASCMDSharedFolder"
$MSASCMDNonEmptyDocument             = ReadConfigFileNode "$environmentResourceFile" "MSASCMDNonEmptyDocument"
$MSASCMDEmptyDocument                = ReadConfigFileNode "$environmentResourceFile" "MSASCMDEmptyDocument"
$MSASCMDEmailSubjectName             = ReadConfigFileNode "$environmentResourceFile" "MSASCMDEmailSubjectName"

$MSASCNTCUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASCNTCUser01"
$MSASCNTCUser02                      = ReadConfigFileNode "$environmentResourceFile" "MSASCNTCUser02"

$MSASCONUser01                       = ReadConfigFileNode "$environmentResourceFile" "MSASCONUser01"
$MSASCONUser02                       = ReadConfigFileNode "$environmentResourceFile" "MSASCONUser02"
$MSASCONUser03                       = ReadConfigFileNode "$environmentResourceFile" "MSASCONUser03"

$MSASDOCUser01                       = ReadConfigFileNode "$environmentResourceFile" "MSASDOCUser01"
$MSASDOCSharedFolder                 = ReadConfigFileNode "$environmentResourceFile" "MSASDOCSharedFolder"
$MSASDOCVisibleFolder                = ReadConfigFileNode "$environmentResourceFile" "MSASDOCVisibleFolder"
$MSASDOCHiddenFolder                 = ReadConfigFileNode "$environmentResourceFile" "MSASDOCHiddenFolder"
$MSASDOCVisibleDocument              = ReadConfigFileNode "$environmentResourceFile" "MSASDOCVisibleDocument"
$MSASDOCHiddenDocument               = ReadConfigFileNode "$environmentResourceFile" "MSASDOCHiddenDocument"

$MSASEMAILUser01                     = ReadConfigFileNode "$environmentResourceFile" "MSASEMAILUser01"
$MSASEMAILUser02                     = ReadConfigFileNode "$environmentResourceFile" "MSASEMAILUser02"
$MSASEMAILUser03                     = ReadConfigFileNode "$environmentResourceFile" "MSASEMAILUser03"
$MSASEMAILUser04                     = ReadConfigFileNode "$environmentResourceFile" "MSASEMAILUser04"
$MSASEMAILUser05                     = ReadConfigFileNode "$environmentResourceFile" "MSASEMAILUser05"

$MSASHTTPUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASHTTPUser01"
$MSASHTTPUser02                      = ReadConfigFileNode "$environmentResourceFile" "MSASHTTPUser02"
$MSASHTTPUser03                      = ReadConfigFileNode "$environmentResourceFile" "MSASHTTPUser03"
$MSASHTTPUser04                      = ReadConfigFileNode "$environmentResourceFile" "MSASHTTPUser04"

$MSASNOTEUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASNOTEUser01"

$MSASPROVUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASPROVUser01"
$MSASPROVUser02                      = ReadConfigFileNode "$environmentResourceFile" "MSASPROVUser02"
$MSASPROVUser03                      = ReadConfigFileNode "$environmentResourceFile" "MSASPROVUser03"

$MSASRMUser01                        = ReadConfigFileNode "$environmentResourceFile" "MSASRMUser01"
$MSASRMUser02                        = ReadConfigFileNode "$environmentResourceFile" "MSASRMUser02"
$MSASRMUser03                        = ReadConfigFileNode "$environmentResourceFile" "MSASRMUser03"
$MSASRMUser04                        = ReadConfigFileNode "$environmentResourceFile" "MSASRMUser04"

$MSASTASKUser01                      = ReadConfigFileNode "$environmentResourceFile" "MSASTASKUser01"

#-----------------------------------------------------
# Paths for all PTF configuration files.
#-----------------------------------------------------
$commonDeploymentFile     = resolve-path "..\..\Source\Common\ExchangeCommonConfiguration.deployment.ptfconfig"
$MSASAIRSDeploymentFile   = resolve-path "..\..\Source\MS-ASAIRS\TestSuite\MS-ASAIRS_TestSuite.deployment.ptfconfig"
$MSASCALDeploymentFile    = resolve-path "..\..\Source\MS-ASCAL\TestSuite\MS-ASCAL_TestSuite.deployment.ptfconfig"
$MSASCMDDeploymentFile    = resolve-path "..\..\Source\MS-ASCMD\TestSuite\MS-ASCMD_TestSuite.deployment.ptfconfig"
$MSASCNTCDeploymentFile   = resolve-path "..\..\Source\MS-ASCNTC\TestSuite\MS-ASCNTC_TestSuite.deployment.ptfconfig"
$MSASCONDeploymentFile    = resolve-path "..\..\Source\MS-ASCON\TestSuite\MS-ASCON_TestSuite.deployment.ptfconfig"
$MSASDOCDeploymentFile    = resolve-path "..\..\Source\MS-ASDOC\TestSuite\MS-ASDOC_TestSuite.deployment.ptfconfig"
$MSASEMAILDeploymentFile  = resolve-path "..\..\Source\MS-ASEMAIL\TestSuite\MS-ASEMAIL_TestSuite.deployment.ptfconfig"
$MSASHTTPDeploymentFile   = resolve-path "..\..\Source\MS-ASHTTP\TestSuite\MS-ASHTTP_TestSuite.deployment.ptfconfig"
$MSASNOTEDeploymentFile   = resolve-path "..\..\Source\MS-ASNOTE\TestSuite\MS-ASNOTE_TestSuite.deployment.ptfconfig"
$MSASPROVDeploymentFile   = resolve-path "..\..\Source\MS-ASPROV\TestSuite\MS-ASPROV_TestSuite.deployment.ptfconfig"
$MSASRMDeploymentFile     = resolve-path "..\..\Source\MS-ASRM\TestSuite\MS-ASRM_TestSuite.deployment.ptfconfig"
$MSASTASKDeploymentFile   = resolve-path "..\..\Source\MS-ASTASK\TestSuite\MS-ASTASK_TestSuite.deployment.ptfconfig"
#-----------------------------------------------------
# Check and make sure that the SUT configuration is finished before running the client setup script.
#-----------------------------------------------------
OutputQuestion "The SUT must be configured before running the client setup script."
OutputQuestion "Did you either run the SUT setup script or configure the SUT as described by the Test Suite Deployment Guide? (Y/N)"
$isSutConfiguredChoices = @("Y","N")
$isSutConfigured = ReadUserChoice $isSutConfiguredChoices "isSutConfigured"
if($isSutConfigured -eq "N")
{
    OutputText "You input `"N`"."
    OutputWarning "Exiting the client setup script now."
    OutputWarning "Configure the SUT and run the client setup script again."
    Stop-Transcript
    exit 0
}
#-----------------------------------------------------
# Check whether the unattended client configuration XML is available if run in unattended mode.
#-----------------------------------------------------
if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
{    
    OutputText "The client setup script will run in attended mode." 
}
else
{
    While($unattendedXmlName -ne "" -and $unattendedXmlName -ne $null)
    {   
        if(Test-Path $unattendedXmlName -PathType Leaf)
        {
            OutputText "The client setup script will run in unattended mode with information provided by the client configuration XML `"$unattendedXmlName`"." 
            $unattendedXmlName = Resolve-Path $unattendedXmlName
            break
        }
        else
        {
            OutputWarning "The client configuration XML path `"$unattendedXmlName`" is not correct." 
            OutputQuestion "Retry with the correct file path or press `"Enter`" if you want client setup script to run in attended mode?"
            $unattendedXmlName = Read-Host
        }
    }
}

#-----------------------------------------------------
# Check the application environment
#-----------------------------------------------------
OutputText "Check whether the required applications have been installed ..." 
$vsInstalledStatus = CheckVSVersion "12.0"
$ptfInstalledStatus = CheckPTFVersion "1.0.2220.0"
if(!$vsInstalledStatus -or !$ptfInstalledStatus)
{
    OutputQuestion "Would you like to exit and install the application(s) that highlighted as yellow in above or continue without installing the application(s)?" 
    OutputQuestion "1: CONTINUE (Without installing the recommended application(s) , it may cause some risk on running the test cases)." 
    OutputQuestion "2: EXIT."     
    $runWithoutRequiredAppInstalledChoices = @('1','2')
    $runWithoutRequiredAppInstalled = ReadUserChoice $runWithoutRequiredAppInstalledChoices "runWithoutRequiredAppInstalled"
    if($runWithoutRequiredAppInstalled -eq "2")
    {
        Stop-Transcript
        cmd.exe /c ECHO CONFIG FINISHED>C:\config.finished.signal
    }
}

#-----------------------------------------------------
# Check the Operating System (OS) version
#-----------------------------------------------------
OutputText "Check the Operating System (OS) version of the local machine ..."
CheckOSVersion -computer localhost

#-----------------------------------------------------
# Configuration for common ptfconfig file.
#-----------------------------------------------------
OutputText "Configure the ExchangeCommonConfiguration.deployment.ptfconfig file ..." 
OutputQuestion "Enter the computer name of the SUT:" 
OutputWarning "The computer name must be valid. Fully qualified domain name (FQDN) or IP address is not supported." 
$sutComputerName = ReadComputerName $false "sutComputerName"
OutputText "Name of the SUT you entered: $sutcomputerName" 

OutputQuestion "Enter the domain name of SUT(for example: contoso.com):" 
$dnsDomain = CheckForEmptyUserInput "Domain name" "dnsDomain"
OutputText "The domain name you entered: $dnsDomain" 

OutputQuestion "Select the Microsoft Exchange Server version" 
OutputQuestion "If you are running your own server implementation, choose the closest exchange server version which matches your implementation." 
OutputQuestion "1: Microsoft Exchange Server 2007" 
OutputQuestion "2: Microsoft Exchange Server 2010" 
OutputQuestion "3: Microsoft Exchange Server 2013" 
OutputQuestion "4: Microsoft Exchange Server 2016" 
OutputQuestion "5: Microsoft Exchange Server 2019"

$sutVersions =@('1','2','3','4','5')
$sutVersion = ReadUserChoice $sutVersions "sutVersion" 
Switch ($sutVersion)
{
    "1" { $sutVersion = "ExchangeServer2007"; $protocolVersion ="12.1"; break }
    "2" { $sutVersion = "ExchangeServer2010"; break }
    "3" { $sutVersion = "ExchangeServer2013"; break }
    "4" { $sutVersion = "ExchangeServer2016"; break }
    "5" { $sutVersion = "ExchangeServer2019"; break }
}
OutputText "The SUT version you selected is $sutVersion." 

if($sutVersion -ge "ExchangeServer2010")
{
    OutputQuestion "Select ActiveSync protocol version. Test suites will use this version while sending requests." 
    OutputQuestion "1: Protocol version is 12.1" 
    OutputQuestion "2: Protocol version is 14.0" 
    OutputQuestion "3: Protocol version is 14.1" 
    OutputQuestion "4: Protocol version is 16.0" 
    OutputQuestion "5: Protocol version is 16.1" 
    OutputQuestion "For Exchange 2010 and 2013, the supported values are 12.1,14.0 and 14.1."
    $protocolVersions =@('1','2','3','4','5')
    $protocolVersion = ReadUserChoice $protocolVersions "protocolVersion"
    Switch ($protocolVersion)
    {
        "1" {$protocolVersion = "12.1"; break}
        "2" {$protocolVersion = "14.0"; break}
        "3" {$protocolVersion = "14.1"; break}
        "4" {$protocolVersion = "16.0"; break}
        "5" {$protocolVersion = "16.1"; break}
    }
}
OutputText "The ActiveSync protocol version you selected is $protocolVersion." 

OutputQuestion "Select the transport type" 
OutputQuestion "1: HTTP" 
OutputQuestion "2: HTTPS" 

$transportTypes =@('1','2')
$transportType = ReadUserChoice $transportTypes "transportType"
Switch ($transportType)
{
    "1" { $transportType = "HTTP";  break }
    "2" { $transportType = "HTTPS"; break }
}
OutputText "Transport type you entered: $transportType" 

OutputQuestion "Select encoding scheme for the URL query string." 
OutputQuestion "1: Test suites will use base64 encoding for the URL query string" 
OutputQuestion "2: Test suites will use plaintext encoding for the URL query string" 
$headerEncodingTypes =@('1','2')
$headerEncodingType = ReadUserChoice $headerEncodingTypes "headerEncodingType"
Switch ($headerEncodingType)
{
    "1" {$headerEncodingType = "Base64"; break}
    "2" {$headerEncodingType = "PlainText"; break}
}
OutputText "Head encoding type you selected is $headerEncodingType." 

OutputWarning "Add SUT machine to the TrustedHosts configuration setting to ensure WinRM client can process remote calls against SUT machine." 
$service = "WinRM"
$serviceStatus = (Get-Service $service).Status
if($serviceStatus -ne "Running")
{
    Start-Service $service
}
$originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -Force).Value
if ($originalTrustedHosts -ne "*")
{
    if ($originalTrustedHosts -eq "")
    {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$sutComputerName" -Force
    }
    elseif (!($originalTrustedHosts.split(',') -icontains $sutComputerName))
    {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$originalTrustedHosts,$sutComputerName" -Force
    }
}

OutputText "Modify the properties as necessary in the ExchangeCommonConfiguration.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $commonDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"Domain`", and set the value as $dnsDomain" 
$step++
OutputWarning "$step.Find the property `"SutComputerName`", and set the value as $sutcomputerName" 
$step++
OutputWarning "$step.Find the property `"SutVersion`", and set the value as $sutVersion" 
$step++
OutputWarning "$step.Find the property `"TransportType`", and set the value as $transportType" 
$step++
OutputWarning "$step.Find the property `"ActiveSyncProtocolVersion`", and set the value as $protocolVersion" 
$step++
OutputWarning "$step.Find the property `"HeaderEncodingType`", and set the value as $headerEncodingType" 

ModifyConfigFileNode $commonDeploymentFile "Domain"                      $dnsDomain
ModifyConfigFileNode $commonDeploymentFile "SutComputerName"             $sutComputerName
ModifyConfigFileNode $commonDeploymentFile "SutVersion"                  $sutVersion
ModifyConfigFileNode $commonDeploymentFile "TransportType"               $transportType
ModifyConfigFileNode $commonDeploymentFile "ActiveSyncProtocolVersion"   $protocolVersion
ModifyConfigFileNode $commonDeploymentFile "HeaderEncodingType"          $headerEncodingType

OutputSuccess "Configuration for the ExchangeCommonConfiguration.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASAIRS ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASAIRS_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:"
OutputWarning "$step.Open $MSASAIRSDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASAIRSUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASAIRSUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User1Name"             "$MSASAIRSUser01"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User1Password"         "$userPassword"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User2Name"             "$MSASAIRSUser02"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User2Password"         "$userPassword"
OutputSuccess "Configuration for the MS-ASAIRS_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASCAL ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASCAL_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASCALDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"OrganizerUserName`", and set the value as $MSASCALUser01" 
$step++
OutputWarning "$step.Find the property `"OrganizerUserPassword`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"AttendeeUserName`", and set the value as $MSASCALUser02" 
$step++
OutputWarning "$step.Find the property `"AttendeeUserPassword`", and set the value as $userPassword" 
ModifyConfigFileNode $MSASCALDeploymentFile    "OrganizerUserName"            "$MSASCALUser01"
ModifyConfigFileNode $MSASCALDeploymentFile    "OrganizerUserPassword"        "$userPassword"
ModifyConfigFileNode $MSASCALDeploymentFile    "AttendeeUserName"             "$MSASCALUser02"
ModifyConfigFileNode $MSASCALDeploymentFile    "AttendeeUserPassword"         "$userPassword"

OutputSuccess "Configuration for the MS-ASCAL_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASCMD ptfconfig file.
#-------------------------------------------------------
$MSASCMDSharedFolderPath = "\\[SutComputerName]\$MSASCMDSharedFolder"
$MSASCMDNonEmptyDocumentPath= "[SharedFolder]\$MSASCMDNonEmptyDocument"
$MSASCMDEmptyDocumentPath= "[SharedFolder]\$MSASCMDEmptyDocument"

OutputText "Modify the properties as necessary in the MS-ASCMD_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASCMDDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASCMDUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASCMDUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASCMDUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User4Name`", and set the value as $MSASCMDUser04" 
$step++
OutputWarning "$step.Find the property `"User4Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User7Name`", and set the value as $MSASCMDUser07" 
$step++
OutputWarning "$step.Find the property `"User7Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User8Name`", and set the value as $MSASCMDUser08" 
$step++
OutputWarning "$step.Find the property `"User8Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User9Name`", and set the value as $MSASCMDUser09" 
$step++
OutputWarning "$step.Find the property `"User9Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"GroupDisplayName`", and set the value as $MSASCMDTestGroup " 
$step++
OutputWarning "$step.Find the property `"LargeGroupDisplayName`", and set the value as $MSASCMDLargeGroup " 
$step++
OutputWarning "$step.Find the property `"SharedFolder`", and set the value as $MSASCMDSharedFolderPath " 
$step++
OutputWarning "$step.Find the property `"SharedDocument1`", and set the value as $MSASCMDNonEmptyDocumentPath " 
$step++
OutputWarning "$step.Find the property `"SharedDocument2`", and set the value as $MSASCMDEmptyDocumentPath" 
$step++
OutputWarning "$step.Find the property `"MIMEMailSubject`", and set the value as $MSASCMDEmailSubjectName" 

ModifyConfigFileNode $MSASCMDDeploymentFile    "User1Name"                    "$MSASCMDUser01"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User1Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User2Name"                    "$MSASCMDUser02"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User2Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User3Name"                    "$MSASCMDUser03"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User3Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User4Name"                    "$MSASCMDUser04"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User4Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User7Name"                    "$MSASCMDUser07"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User7Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User8Name"                    "$MSASCMDUser08"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User8Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User9Name"                    "$MSASCMDUser09"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User9Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "GroupDisplayName"             "$MSASCMDTestGroup"
ModifyConfigFileNode $MSASCMDDeploymentFile    "LargeGroupDisplayName"        "$MSASCMDLargeGroup"
ModifyConfigFileNode $MSASCMDDeploymentFile    "SharedFolder"                 "$MSASCMDSharedFolderPath"
ModifyConfigFileNode $MSASCMDDeploymentFile    "SharedDocument1"              "$MSASCMDNonEmptyDocumentPath"
ModifyConfigFileNode $MSASCMDDeploymentFile    "SharedDocument2"              "$MSASCMDEmptyDocumentPath"
ModifyConfigFileNode $MSASCMDDeploymentFile    "MIMEMailSubject"              "$MSASCMDEmailSubjectName"

OutputSuccess "Configuration for the MS-ASCMD_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASCNTC ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASCNTC_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASCNTCDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASCNTCUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASCNTCUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASCNTCDeploymentFile    "User1Name"            $MSASCNTCUser01
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User1Password"        $userPassword
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User2Name"            $MSASCNTCUser02
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User2Password"        $userPassword

OutputSuccess "Configuration for the MS-ASCNTC_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASCON ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASCON_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASCONDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASCONUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASCONUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASCONUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASCONDeploymentFile "User1Name"      $MSASCONUser01
ModifyConfigFileNode $MSASCONDeploymentFile "User1Password"  $userPassword
ModifyConfigFileNode $MSASCONDeploymentFile "User2Name"      $MSASCONUser02
ModifyConfigFileNode $MSASCONDeploymentFile "User2Password"  $userPassword
ModifyConfigFileNode $MSASCONDeploymentFile "User3Name"      $MSASCONUser03
ModifyConfigFileNode $MSASCONDeploymentFile "User3Password"  $userPassword

OutputSuccess "Configuration for the MS-ASCON_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASDOC ptfconfig file.
#-------------------------------------------------------
# Get the property value of MS-ASDOC ptfconfig file.
$MSASDOCSharedFolderPath = "\\[SutComputerName]\$MSASDOCSharedFolder"
$visibleDocumentPath = "[SharedFolder]\$MSASDOCVisibleDocument"
$hiddenDocumentPath = "[SharedFolder]\$MSASDOCHiddenDocument"
$hiddenFolderPath = "[SharedFolder]\$MSASDOCHiddenFolder"
$visibleFolderPath = "[SharedFolder]\$MSASDOCVisibleFolder"

OutputText "Modify the properties as necessary in the MS-ASDOC_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASDOCDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSASDOCUser01" 
$step++
OutputWarning "$step.Find the property `"UserPassword`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"SharedFolder`", and set the value as $MSASDOCSharedFolderPath" 
$step++
OutputWarning "$step.Find the property `"SharedHiddenDocument`", and set the value as $hiddenDocumentPath" 
$step++
OutputWarning "$step.Find the property `"SharedVisibleDocument`", and set the value as $visibleDocumentPath" 
$step++
OutputWarning "$step.Find the property `"SharedHiddenFolder`", and set the value as $hiddenFolderPath" 
$step++
OutputWarning "$step.Find the property `"SharedVisibleFolder`", and set the value as $visibleFolderPath" 

ModifyConfigFileNode $MSASDOCDeploymentFile "UserName"                   "$MSASDOCUser01"
ModifyConfigFileNode $MSASDOCDeploymentFile "UserPassword"               "$userPassword"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedFolder"               "$MSASDOCSharedFolderPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedHiddenDocument"       "$hiddenDocumentPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedVisibleDocument"      "$visibleDocumentPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedHiddenFolder"         "$hiddenFolderPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedVisibleFolder"        "$visibleFolderPath"

OutputSuccess "Configuration for the MS-ASDOC_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASEMAIL ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASEMAIL_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASEMAILDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASEMAILUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASEMAILUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASEMAILUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword"
$step++
OutputWarning "$step.Find the property `"User4Name`", and set the value as $MSASEMAILUser04" 
$step++
OutputWarning "$step.Find the property `"User4Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User5Name`", and set the value as $MSASEMAILUser05" 
$step++
OutputWarning "$step.Find the property `"User5Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASEMAILDeploymentFile "User1Name"                   $MSASEMAILUser01
ModifyConfigFileNode $MSASEMAILDeploymentFile "User1Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User2Name"                   $MSASEMAILUser02
ModifyConfigFileNode $MSASEMAILDeploymentFile "User2Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User3Name"                   $MSASEMAILUser03
ModifyConfigFileNode $MSASEMAILDeploymentFile "User3Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User4Name"                   $MSASEMAILUser04
ModifyConfigFileNode $MSASEMAILDeploymentFile "User4Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User5Name"                   $MSASEMAILUser05
ModifyConfigFileNode $MSASEMAILDeploymentFile "User5Password"               $userPassword

OutputSuccess "Configuration for the MS-ASEMAIL_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASHTTP ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASHTTP_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASHTTPDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASHTTPUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASHTTPUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASHTTPUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User4Name`", and set the value as $MSASHTTPUser04" 
$step++
OutputWarning "$step.Find the property `"User4Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASHTTPDeploymentFile "User1Name"                  $MSASHTTPUser01
ModifyConfigFileNode $MSASHTTPDeploymentFile "User1Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User2Name"                  $MSASHTTPUser02
ModifyConfigFileNode $MSASHTTPDeploymentFile "User2Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User3Name"                  $MSASHTTPUser03
ModifyConfigFileNode $MSASHTTPDeploymentFile "User3Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User4Name"                  $MSASHTTPUser04
ModifyConfigFileNode $MSASHTTPDeploymentFile "User4Password"              $userPassword

OutputSuccess "Configuration for the MS-ASHTTP_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASNOTE ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASNOTE_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASNOTEDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSASNOTEUser01" 
$step++
OutputWarning "$step.Find the property `"UserPassword`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASNOTEDeploymentFile    "UserName"             "$MSASNOTEUser01"
ModifyConfigFileNode $MSASNOTEDeploymentFile    "UserPassword"         "$userPassword"
OutputSuccess "Configuration for the MS-ASNOTE_TestSuite.deployment.ptfconfig file is complete." 
#-------------------------------------------------------
# Configuration for MS-ASPROV ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASPROV_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASPROVDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASPROVUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASPROVUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASPROVUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASPROVDeploymentFile "User1Name"                $MSASPROVUser01
ModifyConfigFileNode $MSASPROVDeploymentFile "User1Password"            $userPassword
ModifyConfigFileNode $MSASPROVDeploymentFile "User2Name"                $MSASPROVUser02
ModifyConfigFileNode $MSASPROVDeploymentFile "User2Password"            $userPassword
ModifyConfigFileNode $MSASPROVDeploymentFile "User3Name"                $MSASPROVUser03
ModifyConfigFileNode $MSASPROVDeploymentFile "User3Password"            $userPassword

OutputSuccess "Configuration for the MS-ASPROV_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASRM ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASRM_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASRMDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"User1Name`", and set the value as $MSASRMUser01" 
$step++
OutputWarning "$step.Find the property `"User1Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User2Name`", and set the value as $MSASRMUser02" 
$step++
OutputWarning "$step.Find the property `"User2Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User3Name`", and set the value as $MSASRMUser03" 
$step++
OutputWarning "$step.Find the property `"User3Password`", and set the value as $userPassword" 
$step++
OutputWarning "$step.Find the property `"User4Name`", and set the value as $MSASRMUser04" 
$step++
OutputWarning "$step.Find the property `"User4Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASRMDeploymentFile "User1Name"                $MSASRMUser01
ModifyConfigFileNode $MSASRMDeploymentFile "User1Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User2Name"                $MSASRMUser02
ModifyConfigFileNode $MSASRMDeploymentFile "User2Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User3Name"                $MSASRMUser03
ModifyConfigFileNode $MSASRMDeploymentFile "User3Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User4Name"                $MSASRMUser04
ModifyConfigFileNode $MSASRMDeploymentFile "User4Password"            $userPassword

OutputSuccess "Configuration for the MS-ASRM_TestSuite.deployment.ptfconfig file is complete." 

#-------------------------------------------------------
# Configuration for MS-ASTASK ptfconfig file.
#-------------------------------------------------------
OutputText "Modify the properties as necessary in the MS-ASTASK_TestSuite.deployment.ptfconfig file..." 
$step=1
OutputWarning "Steps for manual configuration:" 
OutputWarning "$step.Open $MSASTASKDeploymentFile" 
$step++
OutputWarning "$step.Find the property `"UserName`", and set the value as $MSASTASKUser01" 
$step++
OutputWarning "$step.Find the property `"Password`", and set the value as $userPassword" 

ModifyConfigFileNode $MSASTASKDeploymentFile "UserName"           $MSASTASKUser01
ModifyConfigFileNode $MSASTASKDeploymentFile "Password"           $userPassword

OutputSuccess "Configuration for the MS-ASTASK_TestSuite.deployment.ptfconfig file is complete." 

#----------------------------------------------------------------------------
# End script
#----------------------------------------------------------------------------

OutputSuccess "Client configuration script was executed successfully." 
Stop-Transcript
cmd.exe /c ECHO CONFIG FINISHED>C:\config.finished.signal
