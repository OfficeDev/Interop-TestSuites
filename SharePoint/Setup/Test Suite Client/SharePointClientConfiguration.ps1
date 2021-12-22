#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

#----------------------------------------------------------------------------
# <param name="unattendedXmlName">The unattended client configuration XML.</param>
#----------------------------------------------------------------------------
param(
[string]$unattendedXmlName
)

#-----------------------------------------------------
# Import the Common Function Library File
#-----------------------------------------------------
$scriptDirectory = Split-Path $MyInvocation.Mycommand.Path 
$commonScriptDirectory = $scriptDirectory.SubString(0,$scriptDirectory.LastIndexOf("\")+1) +"Common"       
.("$commonScriptDirectory\SharePointCommonConfiguration.ps1")
.("$commonScriptDirectory\CommonConfiguration.ps1")

#----------------------------------------------------------------------------
# Start script
#----------------------------------------------------------------------------
# Set ErrorActionPreference variable to stop script when error occurs. 
$script:ErrorActionPreference = "Stop"
[String]$containerPath = Get-Location
$logPath        = $containerPath + "\SetupLogs"
$logFile        = $logPath + "\SharePointClientConfiguration.ps1.log"
$debugLogFile   = $logPath + "\SharePointClientConfiguration.ps1.debug.log"

if(!(Test-Path $logPath))
{
    New-Item $logPath -ItemType directory
}elseif([System.IO.File]::Exists($logFile))
{
    Remove-Item $logFile -Force
}
Start-Transcript $debugLogFile -force
AddTimesStampsToLogFile "Start" "$logFile"

#----------------------------------------------------------------------------
# Default Values of Configuration. 
#----------------------------------------------------------------------------
$environmentResourceFile                     = "$commonScriptDirectory\SharePointTestSuite.config"

$MSSITESSSiteCollectionName                  = ReadConfigFileNode "$environmentResourceFile" "MSSITESSSiteCollectionName"
$MSSITESSSite                                = ReadConfigFileNode "$environmentResourceFile" "MSSITESSSite"
$MSSITESSNormalSubSite                       = ReadConfigFileNode "$environmentResourceFile" "MSSITESSNormalSubSite"
$MSSITESSSpecialSubSite                      = ReadConfigFileNode "$environmentResourceFile" "MSSITESSSpecialSubSite"
$MSSITESSDocumentLibrary                     = ReadConfigFileNode "$environmentResourceFile" "MSSITESSDocumentLibrary"
$MSSITESSSubSiteDocumentLibrary              = ReadConfigFileNode "$environmentResourceFile" "MSSITESSSubSiteDocumentLibrary"
$MSSITESSTestData                            = ReadConfigFileNode "$environmentResourceFile" "MSSITESSTestData"
$MSSITESSCustomPage                          = ReadConfigFileNode "$environmentResourceFile" "MSSITESSCustomPage"

$MSVERSSSiteCollectionName                   = ReadConfigFileNode "$environmentResourceFile" "MSVERSSSiteCollectionName"

$MSDWSSSiteCollectionName                    = ReadConfigFileNode "$environmentResourceFile" "MSDWSSSiteCollectionName"
$MSDWSSSiteDocumentWorkpace                  = ReadConfigFileNode "$environmentResourceFile" "MSDWSSSiteDocumentWorkpace"
$MSDWSSSite                                  = ReadConfigFileNode "$environmentResourceFile" "MSDWSSSite"
$MSDWSSTestFolder                            = ReadConfigFileNode "$environmentResourceFile" "MSDWSSTestFolder"
$MSDWSSTestData                              = ReadConfigFileNode "$environmentResourceFile" "MSDWSSTestData"
$MSDWSSInheritPermissionSite                 = ReadConfigFileNode "$environmentResourceFile" "MSDWSSInheritPermissionSite"
$MSDWSSNoneRoleUser                          = ReadConfigFileNode "$environmentResourceFile" "MSDWSSNoneRoleUser"
$MSDWSSNoneRoleUserPassword                  = ReadConfigFileNode "$environmentResourceFile" "MSDWSSNoneRoleUserPassword"
$MSDWSSReaderRoleUser                        = ReadConfigFileNode "$environmentResourceFile" "MSDWSSReaderRoleUser"
$MSDWSSReaderRoleUserPassword                = ReadConfigFileNode "$environmentResourceFile" "MSDWSSReaderRoleUserPassword"
$MSDWSSDocumentLibrary                       = ReadConfigFileNode "$environmentResourceFile" "MSDWSSDocumentLibrary"
$MSDWSSGroupName                             = ReadConfigFileNode "$environmentResourceFile" "MSDWSSGroupName"
$MSDWSSGroupOwner                            = ReadConfigFileNode "$environmentResourceFile" "MSDWSSGroupOwner"
$MSDWSSGroupOwnerPassword                    = ReadConfigFileNode "$environmentResourceFile" "MSDWSSGroupOwnerPassword"

$MSLISTSWSSiteCollectionName                 = ReadConfigFileNode "$environmentResourceFile" "MSLISTSWSSiteCollectionName"
$MSLISTSWSDocumentLibrary                    = ReadConfigFileNode "$environmentResourceFile" "MSLISTSWSDocumentLibrary"

$MSMEETSUser                                 = ReadConfigFileNode "$environmentResourceFile" "MSMEETSUser"
$MSMEETSUserPassword                         = ReadConfigFileNode "$environmentResourceFile" "MSMEETSUserPassword"
$MSMEETSSiteCollectionName                   = ReadConfigFileNode "$environmentResourceFile" "MSMEETSSiteCollectionName"

$MSWEBSSSiteCollectionName                   = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSSiteCollectionName"
$MSWEBSSDocumentLibrary                      = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSDocumentLibrary"
$MSWEBSSSite                                 = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSSite"
$MSWEBSSSiteDescription                      = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSSiteDescription"
$MSWEBSSSiteTitle                            = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSSiteTitle"
$MSWEBSSTestData                             = ReadConfigFileNode "$environmentResourceFile" "MSWEBSSTestData"

$MSWDVMODUUSiteCollectionName                = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUSiteCollectionName"
$MSWDVMODUUDocumentLibrary1                  = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUDocumentLibrary1"
$MSWDVMODUUDocumentLibrary2                  = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUDocumentLibrary2"
$MSWDVMODUUTestFolder                        = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUTestFolder"
$MSWDVMODUUTestData1                         = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUTestData1"
$MSWDVMODUUTestData2                         = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUTestData2"
$MSWDVMODUUTestData3                         = ReadConfigFileNode "$environmentResourceFile" "MSWDVMODUUTestData3"

$MSOUTSPSSiteCollectionName                  = ReadConfigFileNode "$environmentResourceFile" "MSOUTSPSSiteCollectionName"

$MSWWSPSiteCollectionName                    = ReadConfigFileNode "$environmentResourceFile" "MSWWSPSiteCollectionName"
$MSWWSPWorkflowName                          = ReadConfigFileNode "$environmentResourceFile" "MSWWSPWorkflowName"
$MSWWSPWorkflowHistoryList                   = ReadConfigFileNode "$environmentResourceFile" "MSWWSPWorkflowHistoryList"
$MSWWSPWorkflowTaskList                      = ReadConfigFileNode "$environmentResourceFile" "MSWWSPWorkflowTaskList"
$MSWWSPDocumentLibrary                       = ReadConfigFileNode "$environmentResourceFile" "MSWWSPDocumentLibrary"
$MSWWSPUserGroupName                         = ReadConfigFileNode "$environmentResourceFile" "MSWWSPUserGroupName"
$MSWWSPUser                                  = ReadConfigFileNode "$environmentResourceFile" "MSWWSPUser"
$MSWWSPUserPassword                          = ReadConfigFileNode "$environmentResourceFile" "MSWWSPUserPassword"

$MSSHDACCWSSiteCollectionName                = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSSiteCollectionName"
$MSSHDACCWSSite                              = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSSite"
$MSSHDACCWSDocumentLibrary                   = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSDocumentLibrary"
$MSSHDACCWSLockedTestData                    = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSLockedTestData"
$MSSHDACCWSCoStatusTestData                  = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSCoStatusTestData"
$MSSHDACCWSTestData                          = ReadConfigFileNode "$environmentResourceFile" "MSSHDACCWSTestData"

$MSVIEWSSSiteCollectionName                  = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSSiteCollectionName"
$MSVIEWSSViewListName                        = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSViewListName"
$MSVIEWSSListItem1                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem1"
$MSVIEWSSListItem2                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem2"
$MSVIEWSSListItem3                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem3"
$MSVIEWSSListItem4                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem4"
$MSVIEWSSListItem5                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem5"
$MSVIEWSSListItem6                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem6"
$MSVIEWSSListItem7                           = ReadConfigFileNode "$environmentResourceFile" "MSVIEWSSListItem7"

$MSOFFICIALFILESiteCollectionName            = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILESiteCollectionName"
$MSOFFICIALFILERoutingRepositorySite         = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILERoutingRepositorySite"
$MSOFFICIALFILENoRoutingRepositorySite       = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILENoRoutingRepositorySite"
$MSOFFICIALFILEEnabledParsingRepositorySite  = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEEnabledParsingRepositorySite"
$MSOFFICIALFILEDropOffLibrary                = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEDropOffLibrary"
$MSOFFICIALFILEDocumentRuleLocationLibrary   = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEDocumentRuleLocationLibrary"      
$MSOFFICIALFILENoEnforceLibrary              = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILENoEnforceLibrary"
$MSOFFICIALFILEDocumentSetLocationLibrary    = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEDocumentSetLocationLibrary"
$MSOFFICIALFILEDocumentSetName               = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEDocumentSetName"
$MSOFFICIALFILEReadUser                      = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEReadUser"
$MSOFFICIALFILEReadUserPassword              = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEReadUserPassword"
$MSOFFICIALFILEHolds                         = ReadConfigFileNode "$environmentResourceFile" "MSOFFICIALFILEHolds"

$MSAUTHWSFormsWebAPPName                     = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSFormsWebAPPName"
$MSAUTHWSFormsWebAPPPort                     = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSFormsWebAPPPort"
$MSAUTHWSFormsWebAPPHTTPSPort                = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSFormsWebAPPHTTPSPort"
$MSAUTHWSPassportWebAPPName                  = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSPassportWebAPPName"
$MSAUTHWSPassportWebAPPPort                  = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSPassportWebAPPPort"
$MSAUTHWSPassportWebAPPHTTPSPort             = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSPassportWebAPPHTTPSPort"
$MSAUTHWSNoneWebAPPName                      = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSNoneWebAPPName"
$MSAUTHWSNoneWebAPPPort                      = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSNoneWebAPPPort"
$MSAUTHWSNoneWebAPPHTTPSPort                 = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSNoneWebAPPHTTPSPort"
$MSAUTHWSWindowsWebAPPName                   = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSWindowsWebAPPName"
$MSAUTHWSWindowsWebAPPPort                   = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSWindowsWebAPPPort"
$MSAUTHWSWindowsWebAPPHTTPSPort              = ReadConfigFileNode "$environmentResourceFile" "MSAUTHWSWindowsWebAPPHTTPSPort"

$httpsPortNumberOnAdminSite                  = ReadConfigFileNode "$environmentResourceFile" "httpsPortNumberOnAdminSite"
$httpsPortNumberOnSUTWebSite                 = ReadConfigFileNode "$environmentResourceFile" "httpsPortNumberOnSUTWebSite"

$MSCPSWSUser                                 = ReadConfigFileNode "$environmentResourceFile" "MSCPSWSUser"
$MSCPSWSUserPassword                         = ReadConfigFileNode "$environmentResourceFile" "MSCPSWSUserPassword"

$MSWSSRESTSiteCollectionName                 = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTSiteCollectionName"
$MSWSSRESTCalendar                           = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTCalendar"
$MSWSSRESTDocumentLibrary                    = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTDocumentLibrary"
$MSWSSRESTDiscussionBoard                    = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTDiscussionBoard"
$MSWSSRESTGenericList                        = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTGenericList"
$MSWSSRESTSurvey                             = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTSurvey"
$MSWSSRESTWorkflowHistoryList                = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTWorkflowHistoryList"
$MSWSSRESTWorkflowTaskList                   = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTWorkflowTaskList"
$MSWSSRESTWorkflowName                       = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTWorkflowName"
$MSWSSRESTBooleanFieldName                   = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTBooleanFieldName"
$MSWSSRESTChoiceFieldName                    = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTChoiceFieldName"
$MSWSSRESTCurrencyFieldName                  = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTCurrencyFieldName"
$MSWSSRESTGridChoiceFieldName                = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTGridChoiceFieldName"
$MSWSSRESTIntegerFieldName                   = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTIntegerFieldName"
$MSWSSRESTMultiChoiceFieldName               = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTMultiChoiceFieldName"
$MSWSSRESTNumberFieldName                    = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTNumberFieldName"
$MSWSSRESTUrlFieldName                       = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTUrlFieldName"
$MSWSSRESTPageSeparatorFieldName             = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTPageSeparatorFieldName"
$MSWSSRESTWorkFlowEventTypeFieldName         = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTWorkFlowEventTypeFieldName"
$MSWSSRESTLookupFieldName                    = ReadConfigFileNode "$environmentResourceFile" "MSWSSRESTLookupFieldName"
$MSWSSREST_SingleChoiceOptions               = ReadConfigFileNode "$environmentResourceFile" "MSWSSREST_SingleChoiceOptions"
$MSWSSREST_SingleChoiceOptions               = $MSWSSREST_SingleChoiceOptions.Split(",")
$MSWSSREST_MultiChoiceOptions                = ReadConfigFileNode "$environmentResourceFile" "MSWSSREST_MultiChoiceOptions"
$MSWSSREST_MultiChoiceOptions                = $MSWSSREST_MultiChoiceOptions.Split(",")

$MSCOPYSSiteCollectionName                   = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSSiteCollectionName"
$MSCOPYSSubSite                              = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSSubSite"
$MSCOPYSSubSiteDocumentLibrary               = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSSubSiteDocumentLibrary"
$MSCOPYSSourceDocumentLibrary                = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSSourceDocumentLibrary"
$MSCOPYSDestinationDocumentLibrary           = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSDestinationDocumentLibrary"
$MSCOPYSTextFieldName                        = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSTextFieldName"
$MSCOPYSWorkFlowEventFieldName               = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSWorkFlowEventFieldName"
$MSCOPYSSourceLibraryFieldValue              = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSSourceLibraryFieldValue"
$MSCOPYSDestinationLibraryFieldValue         = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSDestinationLibraryFieldValue"
$MSCOPYSTestData                             = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSTestData"
$MSCOPYSTestContent                          = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSTestContent"
$MSCOPYSEditUser                             = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSEditUser"
$MSCOPYSEditUserPassword                     = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSEditUserPassword"
$MSCOPYSNoPermissionUser                     = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSNoPermissionUser"
$MSCOPYSNoPermissionUserPassword             = ReadConfigFileNode "$environmentResourceFile" "MSCOPYSNoPermissionUserPassword"

#-----------------------------------------------------
# Paths for all PTF configuration files.
#-----------------------------------------------------
$CommonDeploymentFile = resolve-path "..\..\Source\Common\SharePointCommonConfiguration.deployment.ptfconfig"
$MSSITESSDeploymentFile = resolve-path "..\..\Source\MS-SITESS\TestSuite\MS-SITESS_TestSuite.deployment.ptfconfig"
$MSVERSSDeploymentFile = resolve-path "..\..\Source\MS-VERSS\TestSuite\MS-VERSS_TestSuite.deployment.ptfconfig"
$MSLISTSWSDeploymentFile = resolve-path "..\..\Source\MS-LISTSWS\TestSuite\MS-LISTSWS_TestSuite.deployment.ptfconfig"
$MSDWSSDeploymentFile = resolve-path "..\..\Source\MS-DWSS\TestSuite\MS-DWSS_TestSuite.deployment.ptfconfig"
$MSADMINSDeploymentFile = resolve-path "..\..\Source\MS-ADMINS\TestSuite\MS-ADMINS_TestSuite.deployment.ptfconfig"
$MSMEETSDeploymentFile = resolve-path "..\..\Source\MS-MEETS\TestSuite\MS-MEETS_TestSuite.deployment.ptfconfig"
$MSOUTSPSDeploymentFile = resolve-path "..\..\Source\MS-OUTSPS\TestSuite\MS-OUTSPS_TestSuite.deployment.ptfconfig"
$MSWDVMODUUDeploymentFile = resolve-path "..\..\Source\MS-WDVMODUU\TestSuite\MS-WDVMODUU_TestSuite.deployment.ptfconfig"
$MSWEBSSDeploymentFile = resolve-path "..\..\Source\MS-WEBSS\TestSuite\MS-WEBSS_TestSuite.deployment.ptfconfig"
$MSWWSPDeploymentFile = resolve-path "..\..\Source\MS-WWSP\TestSuite\MS-WWSP_TestSuite.deployment.ptfconfig"
$MSSHDACCWSDeploymentFile = resolve-path "..\..\Source\MS-SHDACCWS\TestSuite\MS-SHDACCWS_TestSuite.deployment.ptfconfig"
$MSAUTHWSDeploymentFile = resolve-path "..\..\Source\MS-AUTHWS\TestSuite\MS-AUTHWS_TestSuite.deployment.ptfconfig"
$MSCPSWSDeploymentFile = resolve-path "..\..\Source\MS-CPSWS\TestSuite\MS-CPSWS_TestSuite.deployment.ptfconfig"
$MSWSSRESTDeploymentFile = resolve-path "..\..\Source\MS-WSSREST\TestSuite\MS-WSSREST_TestSuite.deployment.ptfconfig"
$MSOFFICIALFILEDeploymentFile = resolve-path "..\..\Source\MS-OFFICIALFILE\TestSuite\MS-OFFICIALFILE_TestSuite.deployment.ptfconfig"
$MCOPYSDeploymentFile = resolve-path "..\..\Source\MS-COPYS\TestSuite\MS-COPYS_TestSuite.deployment.ptfconfig"
$MSVIEWSSDeploymentFile = resolve-path "..\..\Source\MS-VIEWSS\TestSuite\MS-VIEWSS_TestSuite.deployment.ptfconfig"

#-----------------------------------------------------
# Check whether the unattended client configuration XML is available if run in unattended mode.
#-----------------------------------------------------
if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
{    
    Output "The client setup script will run in attended mode." "White"
}
else
{
    While($unattendedXmlName -ne "" -and $unattendedXmlName -ne $null)
    {   
        if(Test-Path $unattendedXmlName -PathType Leaf)
        {
            Output "The client setup script will run in unattended mode with information provided by the `"$unattendedXmlName`" file." "White"
            $unattendedXmlName = Resolve-Path $unattendedXmlName
            break
        }
        else
        {
            Output "The client configuration XML path `"$unattendedXmlName`" is not correct." "Yellow"
            Output "Retry with the correct file path or press `"Enter`" if you want the client setup script to run in attended mode." "Cyan"
            $unattendedXmlName = Read-Host
        }
    }
}

#-----------------------------------------------------
# Check and make sure that the SUT configuration is finished before running the client setup script.
#-----------------------------------------------------
Output "The SUT must be configured before running the client setup script." "Cyan"
Output "Did you either run the SUT setup script or configure the SUT as described by the Test Suite Deployment Guide? (Y/N)" "Cyan"
$isSutConfiguredChoices = @("Y","N")
$isSutConfigured = ReadUserChoice $isSutConfiguredChoices "isSutConfigured"
if($isSutConfigured -eq "N")
{
    Output "Exiting the client setup script now." "Yellow"
    Output "Configure the SUT and run the client setup script again." "Yellow"
    Stop-Transcript
    exit 0
}

#-----------------------------------------------------
# Check the Operating System (OS) version.
#-----------------------------------------------------
Output "Check the Operating System (OS) version of the local machine ..." "White"
CheckOSVersion -computer localhost
#-----------------------------------------------------
# Check the Application environment.
#-----------------------------------------------------
Output "Check whether the machine has installed the prerequisite applications..." "White"
$vsInstalledStatus = CheckVSVersion "12.0"
$ptfInstalledStatus = CheckPTFVersion "1.0.2220.0"
if(!$vsInstalledStatus -or !$ptfInstalledStatus)
{
    Output "Would you like to continue without installing the application(s) or exit and install the application(s) (highlighted in yellow above)?" "Cyan"    
    Output "1: CONTINUE (Without installing the recommended application(s) , it may cause some risk on running the test cases)." "Cyan"
    Output "2: EXIT." "Cyan"    
    $runWithoutRequiredAppInstalledChoices = @('1: CONTINUE','2: EXIT')
    $runWithoutRequiredAppInstalled = ReadUserChoice $runWithoutRequiredAppInstalledChoices "runWithoutRequiredAppInstalled"
    if($runWithoutRequiredAppInstalled -eq "2")
    {
        Stop-Transcript
        exit 0
    }
}

#-----------------------------------------------------
# Configuration for common ptfconfig file.
#-----------------------------------------------------
Output "Enter the computer name of the first SUT:" "Cyan"
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported." "Cyan"
$sutComputerName = ReadComputerName $false "sutComputerName"
Output "The computer name of SUT that you entered: $sutComputerName" "White"

Output "Enter the computer name of the second SUT. Press `"Enter`" if it doesn't exist." "Cyan"    
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP addresses are not supported." "Cyan"    
$sut2ComputerName = ReadComputerName $true "sut2ComputerName"
if($sut2ComputerName -ne "")
{
    Output "The computer name of the second SUT you entered: $sut2ComputerName" "White"
}

Output "Check the status of the Windows Remote Management (WinRM) service to make sure that the service is running." "Yellow"
$service = "WinRM"
$serviceStatus = (Get-Service $service).Status
if($serviceStatus -ne "Running")
{
    try
	{
	    Start-Service $service
	}
	catch
	{	
	    if( $error[0].Exception -match "Microsoft.PowerShell.Commands.ServiceCommandException")
		{
		    Output "Failed to start service $service. Start it manually and then press any key to continue ..." "Red"
		    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") | Out-Null	
		}
	}
	
}
else
{
     Output "The status of service is $serviceStatus" "Green"
}

Output "Steps for manual configuration:" "Yellow" 
Output "Add the SUT to the TrustedHosts configuration setting to ensure that the WinRM client can process remote calls against the SUT." "Yellow"
$originalTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -Force).Value

if($originalTrustedHosts -ne "*")
{
    if($originalTrustedHosts -eq "")
    {
        if($sut2ComputerName -eq "")
        {
            $newTrustedHosts = "$sutComputerName"
        }
        else
        {
            $newTrustedHosts = "$sutComputerName,$sut2ComputerName"         
        }        
    }
    else
    {
        $newTrustedHosts = $originalTrustedHosts
        if(!($originalTrustedHosts.split(',') -icontains $sutComputerName))
        {
            $newTrustedHosts = "$newTrustedHosts,$sutComputerName"
        }
        if(($sut2ComputerName -ne "") -and !($originalTrustedHosts.split(',') -icontains $sut2ComputerName))
        {
            $newTrustedHosts = "$newTrustedHosts,$sut2ComputerName"
        }
    }
	
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$newTrustedHosts" -Force
}

Output "Enter the user who will call protocol methods in the test suite and remotely configure the SUT if the SUT control adapter is set to Powershell mode." "Cyan"
Output "The user should be able to create users in Active Directory directory service, be a part of the local admin group on the server, and also be the SUT administrator." "Cyan"
if(($Env:USERDNSDOMAIN -ne $null) -and ($Env:USERDNSDOMAIN -ine $ENV:COMPUTERNAME) -and ($ENV:USERNAME -ne $null))
{
    Output "Current logon user:" "Cyan"
    Output "Domain: $Env:USERDNSDOMAIN" "Cyan"
    Output "Name:   $ENV:USERNAME" "Cyan"
    Output "Would you like to use this user? (Y/N)" "Cyan"
    $useCurrentUserChoices = @("Y","N")
    $useCurrentUser = ReadUserChoice $useCurrentUserChoices "useCurrentUser"   
    if($useCurrentUser -eq "Y")
    {
        $dnsDomain = $Env:USERDNSDOMAIN
        $userName = $ENV:USERNAME
        $useCurrentUser = $true
    }
    else
    {
        $useCurrentUser = $false
    }
}

if(!$useCurrentUser)
{
    Output "Enter the domain name of the SUT (for example: contoso.com):" "Cyan"
    [String]$dnsDomain = CheckForEmptyUserInput "Domain name" "dnsDomain"
	
    Output "The domain name you entered: $dnsDomain" "White"

    Output "Enter the user name:" "Cyan"
    $userName = CheckForEmptyUserInput "User name" "userName"

    Output "The user name you entered: $userName" "White"
}

Output "Enter password:" "Cyan"
$password = CheckForEmptyUserInput "Password" "password"
Output "Password you entered: $password" "White"

Output "Try to get the SharePoint version on the selected server ..." "White"
$sutVersionInfo = GetSharePointServerVersion $sutComputerName ($dnsDomain.split(".")[0]+ "\" + $userName) $password
if($sutVersionInfo -ne $null -and $sutVersionInfo -ne "" -and $sutVersionInfo -ne "Unknown Version")
{
  $sutVersion = $sutVersionInfo[0]
  Output ("The SharePoint version installed on the server is " + $sutVersionInfo[1] +" " + $sutVersionInfo[2]+ ".") "Green"
}
else
{
    $sutVersioninfo = GetSharePointVersionManually
    $sutVersion = $sutVersionInfo[0]
}

Output "Select the transport type: " "Cyan"
Output "1: HTTP" "Cyan"
Output "2: HTTPS" "Cyan"

$transportType = @('1: HTTP','2: HTTPS')
[String]$transportType = ReadUserChoice $transportType "transportType"
if($transportType -eq "1")
{
    $transportType = "HTTP"
}
else
{
    $transportType = "HTTPS"
}

Output "Select the SOAP version: " "Cyan"
Output "1: SOAP11" "Cyan"
Output "2: SOAP12" "Cyan"

$soapVersion = @('1: SOAP11','2: SOAP12')
[String]$soapVersion = ReadUserChoice $soapVersion "soapVersion"

if($soapVersion -eq "1")
{
    $soapVersion = "SOAP11"
}
else
{
    $soapVersion = "SOAP12"
}

Output "Configure the SharePointCommonConfiguration.deployment.ptfconfig file ..." "White"
Output "Modify the properties as necessary in the SharePointCommonConfiguration.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $CommonDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SutComputerName`", and set the value to $sutComputerName." "Yellow"
$step++
Output "$step.Find the property `"Domain`", and set the value to $dnsDomain." "Yellow"
$step++
Output "$step.Find the property `"UserName`", and set the value to $userName." "Yellow"
$step++
Output "$step.Find the property `"Password`", and set the value to $password." "Yellow"
$step++
Output "$step.Find the property `"SutVersion`", and set the value to $sutVersion." "Yellow"
$step++
Output "$step.Find the property `"TransportType`", and set the value to $transportType." "Yellow"
$step++
Output "$step.Find the property `"SoapVersion`", and set the value to $soapVersion." "Yellow"

ModifyConfigFileNode $CommonDeploymentFile "SutComputerName"             $sutComputerName
ModifyConfigFileNode $CommonDeploymentFile "Domain"                      $dnsDomain
ModifyConfigFileNode $CommonDeploymentFile "UserName"                    $userName
ModifyConfigFileNode $CommonDeploymentFile "Password"                    $password
ModifyConfigFileNode $CommonDeploymentFile "SutVersion"                  $sutVersion
ModifyConfigFileNode $CommonDeploymentFile "TransportType"               $transportType
ModifyConfigFileNode $CommonDeploymentFile "SoapVersion"                 $soapVersion

Output "Configuration for the SharePointCommonConfiguration.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-LISTSWS ptfconfig file.
#-------------------------------------------------------
Output "Configure the MS-LISTSWS_TestSuite.deployment.ptfconfig file ..." "White"

# The fully qualified Url of the protocol server endpoint of MS-LISTSWS.
# This endpoint is under site collection $MSLISTSWSSiteCollectionName.
$listswsServiceUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/_vti_bin/lists.asmx"

Output "Modify the properties as necessary in the MS-LISTSWS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSLISTSWSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSLISTSWSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"TargetServiceUrl`", and set the value to $listswsServiceUrl." "Yellow"

ModifyConfigFileNode $MSLISTSWSDeploymentFile "SiteCollectionName"                   $MSLISTSWSSiteCollectionName
ModifyConfigFileNode $MSLISTSWSDeploymentFile "TargetServiceUrl"                     $listswsServiceUrl

#-------------------------------------------------------
# Configuration for MS-SITESS ptfconfig file.
#-------------------------------------------------------
Output "Configure the MS-SITESS_TestSuite.deployment.ptfconfig file ..." "White"

# The URL of the major site collection used by MS-SITESS test suite.
$sitessCollectionUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]"

# The server-relative URL of the major site collection used by MS-SITESS test suite.
$siteCollectionServerRelativeUrl = "sites/[SiteCollectionName]"

# The URL of the protocol server endpoint to invoke MS-SITESS web service.
# This endpoint is under site collection $MSSITESSSiteCollectionName.
$serviceUrl =  "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/_vti_bin/sites.asmx"

$lcids = GetSharePointLCIDs $sutComputerName ($dnsDomain.split(".")[0]+ "\" + $userName) $password "$transportType`://$sutComputerName/sites/$MSSITESSSiteCollectionName"
$defaultLCID = $lcids[0]
$validLCID = $lcids[1]

Output "Enter a language code identifier (LCID) of the language which is not installed on the server:" "Cyan"
Output "For example, if the Chinese (Simplified) package is not installed on the server, enter 2052." "Cyan"
$notInstalledLCID = CheckForEmptyUserInput "language code identifier (LCID)" "notInstalledLCID"
Output "You entered an invalid LCID: $notInstalledLCID" "White"

# The URL of the site $MSSITESSSite created under major site collection directly for MS-SITESS test suite.
$siteUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/[SiteName]"

# The URL of a subsite of site $MSSITESSSite to be exported. On the subsite there is no file uploaded.
$normalSubSiteUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/[SiteName]/$MSSITESSNormalSubSite"

# The URL of a subsite of site $MSSITESSSite to be exported. On the subsite a 24MB txt file and a custom web page file have been uploaded.
$specialSubSiteUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/[SiteName]/[SpecialSubsiteName]"

# The url of the document library with name $MSSITESSDocumentLibrary which is used as the store path of the files exported by the ExportWeb and ExportWorkflowTemplate operations defined in MS-SITESS.
$dataStoreLibraryUrl = $sitessCollectionUrl +"/[ValidLibraryName]"

# The URL for the uploaded custom web page that contains a form to be posted. It is uploaded on the document library $subSiteLibraryName.
$webPageUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/[SiteName]/[SpecialSubsiteName]/$MSSITESSSubSiteDocumentLibrary/$MSSITESSCustomPage"

# The site-collection relative URL of the document library to store the solution files exported by the ExportSolution operation is defined in MS-SITESS.
# The relative Url is a fixed value "Solution Gallery" for SharePoint servers, and should be changed for other products.
$solutionGalleryName = "Solution Gallery"

# The name of a declarative workflow template is on the server, which can be used to create a new workflow by the ExportWorkflowTemplate operation defined in MS-SITESS.
# The workflow template is only applicable for SharePoint Server 2010 and SharePoint Server 2013. 
# The name is a fixed value "Approval - SharePoint 2010" for SharePoint servers, and should be changed for other products.
$workflowTemplateName = "Approval - SharePoint 2010"

Output "Modify the properties as necessary in the MS-SITESS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSSITESSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSSITESSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionUrl`", and set the value to $sitessCollectionUrl." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionPath`", and set the value to $siteCollectionServerRelativeUrl." "Yellow"
$step++
Output "$step.Find the property `"ServiceUrl`", and set the value to $serviceUrl." "Yellow"
$step++
Output "$step.Find the property `"ValidLCID`", and set the value to $validLCID." "Yellow"
$step++
Output "$step.Find the property `"DefaultLCID`", and set the value to $defaultLCID." "Yellow"
$step++
Output "$step.Find the property `"NotInstalledLCID`", and set the value to $notInstalledLCID." "Yellow"
$step++
Output "$step.Find the property `"SiteName`", and set the value to $MSSITESSSite." "Yellow"
$step++
Output "$step.Find the property `"SiteUrl`", and set the value to $siteUrl." "Yellow"
$step++
Output "$step.Find the property `"NormalSubSiteUrl`", and set the value to $normalSubSiteUrl." "Yellow"
$step++
Output "$step.Find the property `"SpecialSubsiteName`", and set the value to $MSSITESSSpecialSubSite." "Yellow"
$step++
Output "$step.Find the property `"SpecialSubSiteUrl`", and set the value to $specialSubSiteUrl." "Yellow"
$step++
Output "$step.Find the property `"ValidLibraryName`", and set the value to $MSSITESSDocumentLibrary." "Yellow"
$step++
Output "$step.Find the property `"DataPath`", and set the value to $dataStoreLibraryUrl." "Yellow"
$step++
Output "$step.Find the property `"WebPageUrl`", and set the value to $webPageUrl." "Yellow"
$step++
Output "$step.Find the property `"SolutionGalleryName`", and set the value to $solutionGalleryName." "Yellow"
$step++
Output "$step.Find the property `"WorkflowTemplateName`", and set the value to $workflowTemplateName." "Yellow"

ModifyConfigFileNode $MSSITESSDeploymentFile "SiteCollectionName"              $MSSITESSSiteCollectionName
ModifyConfigFileNode $MSSITESSDeploymentFile "SiteCollectionUrl"               $sitessCollectionUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "SiteCollectionPath"              $siteCollectionServerRelativeUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "ServiceUrl"                      $serviceUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "ValidLCID"                       $validLCID
ModifyConfigFileNode $MSSITESSDeploymentFile "DefaultLCID"                     $defaultLCID
ModifyConfigFileNode $MSSITESSDeploymentFile "NotInstalledLCID"                $notInstalledLCID
ModifyConfigFileNode $MSSITESSDeploymentFile "SiteName"                        $MSSITESSSite
ModifyConfigFileNode $MSSITESSDeploymentFile "SiteUrl"                         $siteUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "NormalSubSiteUrl"                $normalSubSiteUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "SpecialSubsiteName"              $MSSITESSSpecialSubSite
ModifyConfigFileNode $MSSITESSDeploymentFile "SpecialSubSiteUrl"               $specialSubSiteUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "ValidLibraryName"                $MSSITESSDocumentLibrary
ModifyConfigFileNode $MSSITESSDeploymentFile "DataPath"                        $dataStoreLibraryUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "WebPageUrl"                      $webPageUrl
ModifyConfigFileNode $MSSITESSDeploymentFile "SolutionGalleryName"             $solutionGalleryName
ModifyConfigFileNode $MSSITESSDeploymentFile "WorkflowTemplateName"            $workflowTemplateName

Output "Configuration for the MS-SITESS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-VERSS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-VERSS_TestSuite.deployment.ptfconfig file ..." "White"

# The absolute URL of the site collection which is used by MS-VERSS test suite.
$siteCollectionUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/"

# The URL of the protocol server endpoint used by MS-VERSS test suite.
# This endpoint is under site collection $MSVERSSSiteCollectionName.
$MSVERSSServiceUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/_vti_bin/versions.asmx"

# The URL of the MS-LISTSWS protocol server endpoint.
# This MS-LISTSWS protocol server endpoint is under site collection $MSVERSSSiteCollectionName.
$MSLISTSWSServiceURL = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/_vti_bin/lists.asmx"

Output "Modify the properties as necessary in the MS-VERSS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSVERSSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSVERSSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"RequestUrl`", and set the value to $siteCollectionUrl." "Yellow"
$step++
Output "$step.Find the property `"MSVERSSServiceUrl`", and set the value to $MSVERSSServiceUrl." "Yellow"
$step++
Output "$step.Find the property `"MSLISTSWSServiceURL`", and set the value to $MSLISTSWSServiceURL." "Yellow"

ModifyConfigFileNode $MSVERSSDeploymentFile "SiteCollectionName"         $MSVERSSSiteCollectionName
ModifyConfigFileNode $MSVERSSDeploymentFile "RequestUrl"                 $siteCollectionUrl
ModifyConfigFileNode $MSVERSSDeploymentFile "MSVERSSServiceUrl"          $MSVERSSServiceUrl
ModifyConfigFileNode $MSVERSSDeploymentFile "MSLISTSWSServiceURL"        $MSLISTSWSServiceURL

Output "Configuration for the MS-VERSS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-DWSS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-DWSS_TestSuite.deployment.ptfconfig file ..." "White"

# The URL of the protocol sever endpoint of the site, which does not inherit permission from parent site, is created for MS-DWSS test suite.
# This endpoint is under site collection $MSSITESSSiteCollectionName.
$dwssWebSite = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/$MSDWSSSite/_vti_bin/dws.asmx"

# The URL of the protocol sever endpoint of the site, which inherits permission from parent site, is created for MS-DWSS test suite.
$inheritPermissionSite = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/$MSDWSSInheritPermissionSite/_vti_bin/dws.asmx"

# A valid site-relative URL of the folder on the SUT, the folder has been created for MS-DWSS test suite.
# For example, "MSDWSSDocLibrary/MSDWSSFolder", "MSDWSSDocLibrary" is the document library list under a site, "MSDWSSFolder" is the folder.
$validFolderUrl = "[ValidDocumentLibraryName]/[ValidFolderName]"

# The document name which has been uploaded to the folder of the document library is created for MS-DWSS test suite.
$documentName = [System.IO.Path]::GetFileNameWithoutExtension($MSDWSSTestData) 

# The file extension of document which has been uploaded to the folder of the document library is created for MS-DWSS test suite.
$documentExtension = [System.IO.Path]::GetExtension($MSDWSSTestData)

# The site-relative url of the document $documentName is specified in previous step.
# For example, "MSDWSSDocLibrary/MSDWSSFolder/MSDWSSFile.txt", "MSDWSSDocLibrary" is the document library list under a site.
$validDocumentUrl = "[ValidDocumentLibraryName]/[ValidFolderName]/[DocumentsName]$documentExtension"

# The url of protocol server endpoint is used to get an empty string of WorkspaceType element in response of the GetDwsMetaData operation.
# For example, "http://SUT01/sites/xxxx/_vti_bin/dws.asmx".
$siteCollectionUrl = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/_vti_bin/dws.asmx"

# A site-relative URL of the folder with a new folder in a valid document library. For example: "MSDWSS_DocumentLibrary/NewFolder".
$newFolderUrl = "[ValidDocumentLibraryName]/NewFolder"

# A sever endpoint URL of an invalid site under the default site collection. For example: http://SUT01/NonExistentDWSSWebSiteUrl/_vti_bin/dws.asmx
$nonExistentDWSSWebSiteUrl = "[TransportType]://[SutComputerName]/NonExistentDWSSWebSiteUrl/_vti_bin/dws.asmx"

# A url of protocol server endpoint, it must be under the site collection without subsite. And the web template of the site collection is document workspace. For example: http://SUT01/sites/MSDWSS_SiteCollection_DocumentWorkspace/_vti_bin/dws.asmx
$siteCollectionWithoutSubSite = "[TransportType]://[SutComputerName]/sites/$MSDWSSSiteDocumentWorkpace/_vti_bin/dws.asmx"

Output "Modify the properties as necessary in the MS-DWSS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSDWSSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"ReaderRoleUser`", and set the value to $MSDWSSReaderRoleUser." "Yellow"
$step++
Output "$step.Find the property `"ReaderRoleUserPassword`", and set the value to $MSDWSSReaderRoleUserPassword." "Yellow"
$step++
Output "$step.Find the property `"NoneRoleUser`", and set the value to $MSDWSSNoneRoleUser." "Yellow"
$step++
Output "$step.Find the property `"NoneRoleUserPassword`", and set the value to $MSDWSSNoneRoleUserPassword." "Yellow"
$step++
Output "$step.Find the property `"TestDWSSWebSite`", and set the value to $dwssWebSite." "Yellow"
$step++
Output "$step.Find the property `"InheritPermissionSite`", and set the value to $inheritPermissionSite." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSDWSSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"ValidDocumentLibraryName`", and set the value to $MSDWSSDocumentLibrary." "Yellow"
$step++
Output "$step.Find the property `"ValidFolderName`", and set the value to $MSDWSSTestFolder." "Yellow"
$step++
Output "$step.Find the property `"ValidFolderUrl`", and set the value to $validFolderUrl." "Yellow"
$step++
Output "$step.Find the property `"DocumentsName`", and set the value to $documentName." "Yellow"
$step++
Output "$step.Find the property `"ValidDocumentUrl`", and set the value to $validDocumentUrl." "Yellow"
$step++
Output "$step.Find the property `"SiteCollection`", and set the value to $siteCollectionUrl." "Yellow"
$step++
Output "$step.Find the property `"RegisteredUsersEmail`", and set the value to $UserName@$dnsDomain." "Yellow"
$step++
Output "$step.Find the property `"NewFolderUrl`", and set the value to $newFolderUrl." "Yellow"
$step++
Output "$step.Find the property `"NonExistentDWSSWebSiteUrl`", and set the value to $nonExistentDWSSWebSiteUrl." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionWithoutSubSite`", and set the value to $siteCollectionWithoutSubSite." "Yellow"

ModifyConfigFileNode $MSDWSSDeploymentFile "ReaderRoleUser"                             $MSDWSSReaderRoleUser
ModifyConfigFileNode $MSDWSSDeploymentFile "ReaderRoleUserPassword"                     $MSDWSSReaderRoleUserPassword
ModifyConfigFileNode $MSDWSSDeploymentFile "NoneRoleUser"                               $MSDWSSNoneRoleUser
ModifyConfigFileNode $MSDWSSDeploymentFile "NoneRoleUserPassword"                       $MSDWSSNoneRoleUserPassword
ModifyConfigFileNode $MSDWSSDeploymentFile "TestDWSSWebSite"                            $dwssWebSite
ModifyConfigFileNode $MSDWSSDeploymentFile "InheritPermissionSite"                      $inheritPermissionSite
ModifyConfigFileNode $MSDWSSDeploymentFile "SiteCollectionName"                         $MSDWSSSiteCollectionName
ModifyConfigFileNode $MSDWSSDeploymentFile "ValidDocumentLibraryName"                   $MSDWSSDocumentLibrary
ModifyConfigFileNode $MSDWSSDeploymentFile "ValidFolderName"                            $MSDWSSTestFolder
ModifyConfigFileNode $MSDWSSDeploymentFile "ValidFolderUrl"                             $validFolderUrl
ModifyConfigFileNode $MSDWSSDeploymentFile "DocumentsName"                              $documentName
ModifyConfigFileNode $MSDWSSDeploymentFile "ValidDocumentUrl"                           $validDocumentUrl
ModifyConfigFileNode $MSDWSSDeploymentFile "SiteCollection"                             $siteCollectionUrl
ModifyConfigFileNode $MSDWSSDeploymentFile "RegisteredUsersEmail"                       $userName@$dnsDomain
ModifyConfigFileNode $MSDWSSDeploymentFile "NewFolderUrl"                               $newFolderUrl
ModifyConfigFileNode $MSDWSSDeploymentFile "SiteCollectionWithoutSubSite"               $siteCollectionWithoutSubSite
ModifyConfigFileNode $MSDWSSDeploymentFile "NonExistentDWSSWebSiteUrl"                  $nonExistentDWSSWebSiteUrl

Output "Configuration for the MS-DWSS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-MEETS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-MEETS_TestSuite.deployment.ptfconfig file ..." "White"

# The URL of the protocol server endpoint to invoke MS-MEETS web service.
# This endpoint is under site collection $MSMEETSSiteCollectionName.
$serviceUrl =  "[TransportType]://[SutComputerName]/sites/[SiteCollectionName][EntryUrl]"

# The relative url of the meeting web service endpoint.
$entryUrl = "/_vti_bin/meetings.asmx"

#The e-mail address of the meeting organizer.
$organizerEmail = "$userName@$dnsDomain"

#The e-mail address of the meeting attendee
$attendeeEmail = "$MSMEETSUser@$dnsDomain"

Output "Modify the properties as necessary in the MS-MEETS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSMEETSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"TargetServiceUrl`", and set the value to $serviceUrl." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSMEETSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"EntryUrl`", and set the value to $entryUrl." "Yellow"
$step++
Output "$step.Find the property `"OrganizerEmail`", and set the value to $organizerEmail." "Yellow"
$step++
Output "$step.Find the property `"AttendeeEmail`", and set the value to $attendeeEmail." "Yellow"

ModifyConfigFileNode $MSMEETSDeploymentFile "TargetServiceUrl"                  $serviceUrl
ModifyConfigFileNode $MSMEETSDeploymentFile "SiteCollectionName"                $MSMEETSSiteCollectionName
ModifyConfigFileNode $MSMEETSDeploymentFile "EntryUrl"                          $entryUrl
ModifyConfigFileNode $MSMEETSDeploymentFile "OrganizerEmail"                    $organizerEmail
ModifyConfigFileNode $MSMEETSDeploymentFile "AttendeeEmail"                     $attendeeEmail

Output "Configuration for the MS-MEETS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-WWSP ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-WWSP_TestSuite.deployment.ptfconfig file ..." "White"

# The URL of the protocol server endpoint to invoke MS-WWSP web service.
# This endpoint is under site collection $MSWWSPSiteCollectionName.
$targetServiceUrl = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/_vti_bin/workflow.asmx"

Output "Modify the properties as necessary in the MS-WWSP_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSWWSPDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"TargetServiceUrl`", and set the value to $targetServiceUrl." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSWWSPSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"WorkflowAssociationName`", and set the value to $MSWWSPWorkflowName." "Yellow"
$step++
Output "$step.Find the property `"UserGroupOnSUT`", and set the value to $MSWWSPUserGroupName." "Yellow"
$step++
Output "$step.Find the property `"CurrentDocLibraryListName`", and set the value to $MSWWSPDocumentLibrary." "Yellow"
$step++
Output "$step.Find the property `"CurrentTaskListName`", and set the value to $MSWWSPWorkflowTaskList." "Yellow"
$step++
Output "$step.Find the property `"MSWWSPTestAccount`", and set the value to $MSWWSPUser." "Yellow"
$step++
Output "$step.Find the property `"MSWWSPTestAccountPassword`", and set the value to $isStandaloneInstallation." "Yellow"
$step++
Output "$step.Find the property `"KeyWordForAssignedToField`", and set the value to $MSWWSPUser." "Yellow"

ModifyConfigFileNode $MSWWSPDeploymentFile "TargetServiceUrl"                $targetServiceUrl
ModifyConfigFileNode $MSWWSPDeploymentFile "SiteCollectionName"              $MSWWSPSiteCollectionName
ModifyConfigFileNode $MSWWSPDeploymentFile "WorkflowAssociationName"         $MSWWSPWorkflowName
ModifyConfigFileNode $MSWWSPDeploymentFile "UserGroupOnSUT"                  $MSWWSPUserGroupName
ModifyConfigFileNode $MSWWSPDeploymentFile "CurrentDocLibraryListName"        $MSWWSPDocumentLibrary
ModifyConfigFileNode $MSWWSPDeploymentFile "CurrentTaskListName"             $MSWWSPWorkflowTaskList
ModifyConfigFileNode $MSWWSPDeploymentFile "MSWWSPTestAccount"               $MSWWSPUser
ModifyConfigFileNode $MSWWSPDeploymentFile "MSWWSPTestAccountPassword"       $MSWWSPUserPassword
ModifyConfigFileNode $MSWWSPDeploymentFile "KeyWordForAssignedToField"       $MSWWSPUser

Output "Configuration for the MS-WWSP_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-ADMINS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-ADMINS_TestSuite.deployment.ptfconfig file ..." "White"

#The URL of SUT web site when transport type is HTTP.
$webSiteUrl = "${transportType}://$sutComputerName"

#The https port number used by Administration web service on the protocol server.
$adminHTTPSPortNumber = GetHttpsSharePointAdminSitePort $sutComputerName $userName $password

#The https port number used by SharePoint default site.
$httpsPortNumber = GetHttpsSUTWebSitePort $sutComputerName $userName $password $webSiteUrl

#The http port number used by Administration web service on the protocol server.
$adminHTTPPortNumber = GetSharePointAdminSitePort $sutComputerName $userName $password

#The https port number used by SharePoint default site.
$httpPortNumber   = GetSUTWebSitePort $sutComputerName $userName $password $webSiteUrl

Output "Modify the properties as necessary in the MS-ADMINS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSADMINSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"AdminHTTPPortNumber`", and set the value to $adminHTTPPortNumber." "Yellow"
$step++
Output "$step.Find the property `"AdminHTTPSPortNumber`", and set the value to $adminHTTPSPortNumber." "Yellow"
$step++
Output "$step.Find the property `"HTTPPortNumber`", and set the value to $httpPortNumber." "Yellow"
$step++
Output "$step.Find the property `"HTTPSPortNumber`", and set the value to $httpsPortNumber." "Yellow"
$step++
Output "$step.Find the property `"NotInstalledLCID`", and set the value to $notInstalledLCID." "Yellow"

ModifyConfigFileNode $MSADMINSDeploymentFile "AdminHTTPPortNumber"                 $adminHTTPPortNumber
ModifyConfigFileNode $MSADMINSDeploymentFile "AdminHTTPSPortNumber"                $adminHTTPSPortNumber
ModifyConfigFileNode $MSADMINSDeploymentFile "HTTPPortNumber"                      $httpPortNumber
ModifyConfigFileNode $MSADMINSDeploymentFile "HTTPSPortNumber"                     $httpsPortNumber
ModifyConfigFileNode $MSADMINSDeploymentFile "NotInstalledLCID"                    $notInstalledLCID

Output "Configuration for the MS-ADMINS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-OUTSPS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-OUTSPS_TestSuite.deployment.ptfconfig file ..." "White"

Output "Modify the properties as necessary in the MS-OUTSPS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSOUTSPSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSOUTSPSSiteCollectionName." "Yellow"

ModifyConfigFileNode $MSOUTSPSDeploymentFile "SiteCollectionName"                    $MSOUTSPSSiteCollectionName

Output "Configuration for the MS-OUTSPS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-WEBSS ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-WEBSS_TestSuite.deployment.ptfconfig file ..." "White"

#The computer name of the SUT.
$hostName = $sutComputerName

#The name of the document which will be added to test site.
$docName = $MSWEBSSTestData

#The folder name for the document which will be added to test site.
$foldName = $MSWEBSSDocumentLibrary

Output "Modify the properties as necessary in the MS-WEBSS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSWEBSSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSWEBSSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"webSiteName`", and set the value to $MSWEBSSSite." "Yellow"
$step++
Output "$step.Find the property `"TestSiteTitle`", and set the value to $MSWEBSSSiteTitle." "Yellow"
$step++
Output "$step.Find the property `"TestSiteDescription`", and set the value to $MSWEBSSSiteDescription." "Yellow"
$step++
Output "$step.Find the property `"DocName`", and set the value to $docName." "Yellow"
$step++
Output "$step.Find the property `"FoldName`", and set the value to $foldName." "Yellow"

ModifyConfigFileNode $MSWEBSSDeploymentFile "SiteCollectionName"                                 $MSWEBSSSiteCollectionName
ModifyConfigFileNode $MSWEBSSDeploymentFile "webSiteName"                                        $MSWEBSSSite
ModifyConfigFileNode $MSWEBSSDeploymentFile "TestSiteTitle"                                      $MSWEBSSSiteTitle
ModifyConfigFileNode $MSWEBSSDeploymentFile "TestSiteDescription"                                $MSWEBSSSiteDescription
ModifyConfigFileNode $MSWEBSSDeploymentFile "DocName"                                            $docName
ModifyConfigFileNode $MSWEBSSDeploymentFile "FoldName"                                           $foldName

Output "Configuration for the MS-WEBSS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-WDVMODUU ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-WDVMODUU_TestSuite.deployment.ptfconfig file ..." "White"

#The value of the property document library $MSWDVMODUUDocumentLibrary1 is under site collection MSWDVMODUU_SiteCollection.
$serverDefaultDocumentLibName = $MSWDVMODUUDocumentLibrary1

#The value of the property is the URI of the test file $MSWDVMODUUTestData1 under site collection MSWDVMODUU_SiteCollection.
$serverNewFile001Uri = "[Server_DefaultDocLibUri]$MSWDVMODUUTestData1"

#The value of the property is the URI of the test file $MSWDVMODUUTestData2 under site collection MSWDVMODUU_SiteCollection.
$serverNewFile002Uri = "[Server_DefaultDocLibUri]$MSWDVMODUUTestData2"

#The value of the property is the URI of the Sub-Folder $MSWDVMODUUTestFolder under the document library $MSWDVMODUUDocumentLibrary1.
$serverSubFolderUri = "[Server_DefaultDocLibUri]$MSWDVMODUUTestFolder/"

#The value of the property is the URI of the test file "$MSWDVMODUUTestData3.
$serverNewFile003Uri = "[Server_SubFolderUri]$MSWDVMODUUTestData3"

#The value of the property is the test document library name that under site collection MSWDVMODUU_SiteCollection.
$serverTestDocumentLibName = $MSWDVMODUUDocumentLibrary2

Output "Modify the properties as necessary in the MS-WDVMODUU_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSWDVMODUUDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSWDVMODUUSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"Server_DefaultDocumentLibName`", and set the value to $MSWDVMODUUDocumentLibrary1." "Yellow"
$step++
Output "$step.Find the property `"Server_NewFile001Uri`", and set the value to $serverNewFile001Uri." "Yellow"
$step++
Output "$step.Find the property `"Server_NewFile002Uri`", and set the value to $serverNewFile002Uri." "Yellow"
$step++
Output "$step.Find the property `"Server_SubFolderUri`", and set the value to $serverSubFolderUri." "Yellow"
$step++
Output "$step.Find the property `"Server_NewFile003Uri`", and set the value to $serverNewFile003Uri." "Yellow"
$step++
Output "$step.Find the property `"Server_TestDocumentLibName`", and set the value to $serverTestDocumentLibName." "Yellow"

ModifyConfigFileNode $MSWDVMODUUDeploymentFile "SiteCollectionName"                         $MSWDVMODUUSiteCollectionName  
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_DefaultDocumentLibName"              $MSWDVMODUUDocumentLibrary1  
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_NewFile001Uri"                       $serverNewFile001Uri  
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_NewFile002Uri"                       $serverNewFile002Uri  
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_SubFolderUri"                        $serverSubFolderUri
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_NewFile003Uri"                       $serverNewFile003Uri
ModifyConfigFileNode $MSWDVMODUUDeploymentFile "Server_TestDocumentLibName"                 $serverTestDocumentLibName

Output "Configuration for the MS-WDVMODUU_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-SHDACCWS ptfconfig file.
#-----------------------------------------------------
if($sutVersion -eq "SharePointFoundation2013" -or $sutVersion -eq "SharePointServer2013" -or $sutVersion -eq "SharePointFoundation2010" -or $sutVersion -eq "SharePointServer2010" -or $sutVersion -eq "SharePointServer2016" -or $sutVersion -eq "SharePointServer2019" -or $sutVersion -eq "SharePointServerSubscriptionEdition")
{
	Output "Configure the MS-SHDACCWS_TestSuite.deployment.ptfconfig file ..." "White"

	# The absolute URL of the site collection which is used by MS-SHDACCWS test suite.
	$MSSHDACCWSsiteCollectionUrl = "${transportType}://$sutComputerName/sites/$MSSHDACCWSSiteCollectionName"
	$targetServiceUrl = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/_vti_bin/sharedaccess.asmx" 

	#Get uploaded file id in SUT.
	$fileIdOfLock = GetFileId $sutComputerName $userName $password $MSSHDACCWSsiteCollectionUrl $MSSHDACCWSDocumentLibrary $MSSHDACCWSLockedTestData
	if($fileIdOfLock -eq $null -or $fileIdOfLock -eq "")
    {
        Output "Cannot get the GUID of the $MSSHDACCWSLockedTestData file automatically. Enter the GUID of the $MSSHDACCWSLockedTestData file on the server:" "Cyan"
        $fileIdOfLock = CheckForEmptyUserInput "The GUID of file $MSSHDACCWSLockedTestData" "fileIdOfLock"
        Output ("The GUID you entered: " + $fileIdOfLock) "White"
    }	
	
	$fileIdOfCoAuthoring = GetFileId $sutComputerName $userName $password $MSSHDACCWSsiteCollectionUrl $MSSHDACCWSDocumentLibrary $MSSHDACCWSCoStatusTestData
	if($fileIdOfCoAuthoring -eq $null -or $fileIdOfCoAuthoring -eq "")
    {
        Output "Cannot get the GUID of the $MSSHDACCWSCoStatusTestData file automatically. Enter the GUID of the $MSSHDACCWSCoStatusTestData file on the server:" "Cyan"
        $fileIdOfCoAuthoring = CheckForEmptyUserInput "The GUID of file $MSSHDACCWSCoStatusTestData" "fileIdOfCoAuthoring"
        Output ("The GUID you entered: " + $fileIdOfCoAuthoring) "White"
    }  
	
	$fileIdOfNormal = GetFileId $sutComputerName $userName $password $MSSHDACCWSsiteCollectionUrl $MSSHDACCWSDocumentLibrary $MSSHDACCWSTestData
	if($fileIdOfNormal -eq $null -or $fileIdOfNormal -eq "")
    {
        Output "Cannot get the GUID of the $MSSHDACCWSTestData file automatically. Enter the GUID of the $MSSHDACCWSTestData file on the server:" "Cyan"
        $fileIdOfNormal = CheckForEmptyUserInput "The GUID of file $MSSHDACCWSTestData" "fileIdOfNormal"
        Output ("The GUID you entered: " + $fileIdOfNormal) "White"
    }  

	Output "Modify the properties as necessary in the MS-SHDACCWS_TestSuite.deployment.ptfconfig file..." "White"
	$step = 1
	Output "Steps for manual configuration:" "Yellow"
	Output "$step.Open $MSSHDACCWSDeploymentFile" "Yellow"
	$step++
	Output "$step.Find the property `"TargetServiceUrl`", and set the value to $targetServiceUrl." "Yellow"
	$step++
	Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSSHDACCWSSiteCollectionName." "Yellow"
	$step++
	Output "$step.Find the property `"FileIdOfLock`", and set the value to $fileIdOfLock." "Yellow"
	$step++
	Output "$step.Find the property `"FileIdOfNormal`", and set the value to $fileIdOfCoAuthoring." "Yellow"
	$step++
	Output "$step.Find the property `"FileIdOfNormal`", and set the value to $fileIdOfNormal." "Yellow"

	ModifyConfigFileNode $MSSHDACCWSDeploymentFile "TargetServiceUrl"                 $targetServiceUrl 
	ModifyConfigFileNode $MSSHDACCWSDeploymentFile "SiteCollectionName"               $MSSHDACCWSSiteCollectionName
	ModifyConfigFileNode $MSSHDACCWSDeploymentFile "FileIdOfLock"                     $fileIdOfLock
	ModifyConfigFileNode $MSSHDACCWSDeploymentFile "FileIdOfCoAuthoring"              $fileIdOfCoAuthoring
	ModifyConfigFileNode $MSSHDACCWSDeploymentFile "FileIdOfNormal"                   $fileIdOfNormal
	
	Output "Configuration for the MS-SHDACCWS_TestSuite.deployment.ptfconfig file is complete." "Green"
	
}
#-----------------------------------------------------
# Configuration for MS-AUTHWS ptfconfig file.
#-----------------------------------------------------

Output "Configure the MS-AUTHWS_TestSuite.deployment.ptfconfig file ..." "White"

if($sutVersion -eq "WindowsSharePointServices3" -or $sutVersion -eq "SharePointServer2007" -or $sutVersion -eq "SharePointFoundation2010" -or $sutVersion -eq "SharePointServer2010")
{
    $MSAUTHWSWindowsWebAPPPort = $httpPortNumber
	$MSAUTHWSWindowsWebAPPHTTPSPort = $httpsPortNumber
}

Output "Modify the properties as necessary in the MS-AUTHWS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSAUTHWSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"WindowsAuthenticationPortForHTTP`", and set the value to $MSAUTHWSWindowsWebAPPPort." "Yellow"
$step++
Output "$step.Find the property `"WindowsAuthenticationPortForHTTPS`", and set the value to $MSAUTHWSWindowsWebAPPHTTPSPort." "Yellow"
$step++
Output "$step.Find the property `"FormsAuthenticationPortForHTTP`", and set the value to $MSAUTHWSFormsWebAPPPort." "Yellow"
$step++
Output "$step.Find the property `"FormsAuthenticationPortForHTTPS`", and set the value to $MSAUTHWSFormsWebAPPHTTPSPort." "Yellow"
$step++
Output "$step.Find the property `"PassportAuthenticationPortForHTTP`", and set the value to $MSAUTHWSPassportWebAPPPort." "Yellow"
$step++
Output "$step.Find the property `"PassportAuthenticationPortForHTTPS`", and set the value to $MSAUTHWSPassportWebAPPHTTPSPort." "Yellow"
$step++
Output "$step.Find the property `"NoneAuthenticationPortForHTTP`", and set the value to $MSAUTHWSNoneWebAPPPort." "Yellow"
$step++
Output "$step.Find the property `"WindowsAuthenticationPortForHTTP`", and set the value to $MSAUTHWSNoneWebAPPHTTPSPort." "Yellow"

ModifyConfigFileNode $MSAUTHWSDeploymentFile "WindowsAuthenticationPortForHTTP"          $MSAUTHWSWindowsWebAPPPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "WindowsAuthenticationPortForHTTPS"         $MSAUTHWSWindowsWebAPPHTTPSPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "FormsAuthenticationPortForHTTP"            $MSAUTHWSFormsWebAPPPort 
ModifyConfigFileNode $MSAUTHWSDeploymentFile "FormsAuthenticationPortForHTTPS"           $MSAUTHWSFormsWebAPPHTTPSPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "PassportAuthenticationPortForHTTP"         $MSAUTHWSPassportWebAPPPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "PassportAuthenticationPortForHTTPS"        $MSAUTHWSPassportWebAPPHTTPSPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "NoneAuthenticationPortForHTTP"             $MSAUTHWSNoneWebAPPPort
ModifyConfigFileNode $MSAUTHWSDeploymentFile "NoneAuthenticationPortForHTTPS"            $MSAUTHWSNoneWebAPPHTTPSPort 

Output "Configuration for the MS-AUTHWS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-CPSWS ptfconfig file.
#-----------------------------------------------------
if($sutVersion -eq "SharePointFoundation2013" -or $sutVersion -eq "SharePointServer2013" -or $sutVersion -eq "SharePointFoundation2010" -or $sutVersion -eq "SharePointServer2010" -or $sutVersion -eq "SharePointServer2016" -or $sutVersion -eq "SharePointServer2019")
{
    Output "Configure the MS-CPSWS_TestSuite.deployment.ptfconfig file ..." "White"
	$validUser = ($dnsDomain.split(".")[0]+ "\" + $MSCPSWSUser)
    Output "Modify the properties as necessary in the MS-CPSWS_TestSuite.deployment.ptfconfig file..." "White"
    $step = 1
    Output "Steps for manual configuration:" "Yellow"
    Output "$step.Open $MSCPSWSDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"ValidUser`", and set the value to $validUser." "Yellow"
	
    ModifyConfigFileNode $MSCPSWSDeploymentFile "ValidUser"           $validUser

    Output "Configuration for the MS-CPSWS_TestSuite.deployment.ptfconfig file is complete." "Green"

}

#-----------------------------------------------------
# Configuration for MS-WSSREST ptfconfig file.
#-----------------------------------------------------
if($sutVersion -eq "SharePointFoundation2013" -or $sutVersion -eq "SharePointServer2013" -or $sutVersion -eq "SharePointFoundation2010" -or $sutVersion -eq "SharePointServer2010" -or $sutVersion -eq "SharePointServer2016" -or $sutVersion -eq "SharePointServer2019")
{
    Output "Configure the MS-WSSREST_TestSuite.deployment.ptfconfig file ..." "White"
    Output "Modify the properties as necessary in the MS-WSSREST_TestSuite.deployment.ptfconfig file..." "White"
	$choiceFieldOptions = ($MSWSSREST_SingleChoiceOptions[0] + ","+ $MSWSSREST_SingleChoiceOptions[1])
	$multiChoiceFieldOptions = ($MSWSSREST_MultiChoiceOptions[0] + "," +$MSWSSREST_MultiChoiceOptions[1])
    $step = 1
    Output "Steps for manual configuration:" "Yellow"
    Output "$step.Open $MSWSSRESTDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSWSSRESTSiteCollectionName." "Yellow"
	$step++
    Output "$step.Find the property `"DoucmentLibraryListName`", and set the value to $MSWSSRESTDocumentLibrary." "Yellow"
	$step++
    Output "$step.Find the property `"GeneralListName`", and set the value to $MSWSSRESTGenericList." "Yellow"
	$step++
    Output "$step.Find the property `"SurveyListName`", and set the value to $MSWSSRESTSurvey." "Yellow"
	$step++
    Output "$step.Find the property `"TaskListName`", and set the value to $MSWSSRESTWorkflowTaskList." "Yellow"
	$step++
    Output "$step.Find the property `"DiscussionBoardListName`", and set the value to $MSWSSRESTDiscussionBoard." "Yellow"
	$step++
    Output "$step.Find the property `"CalendarListName`", and set the value to $MSWSSRESTCalendar." "Yellow"
	$step++
    Output "$step.Find the property `"ChoiceFieldName`", and set the value to $MSWSSRESTChoiceFieldName." "Yellow"
	$step++
    Output "$step.Find the property `"MultiChoiceFieldName`", and set the value to $MSWSSRESTMultiChoiceFieldName ." "Yellow"
	$step++
    Output "$step.Find the property `"WorkflowHistoryListName`", and set the value to $MSWSSRESTWorkflowHistoryList ." "Yellow"
	$step++
    Output "$step.Find the property `"ChoiceFieldOptions`", and set the value to $choiceFieldOptions ." "Yellow"
	$step++
    Output "$step.Find the property `"MultiChoiceFieldOptions`", and set the value to $multiChoiceFieldOptions ." "Yellow"
	
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "SiteCollectionName"                $MSWSSRESTSiteCollectionName
	ModifyConfigFileNode $MSWSSRESTDeploymentFile "DoucmentLibraryListName"           $MSWSSRESTDocumentLibrary
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "GeneralListName"                   $MSWSSRESTGenericList
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "SurveyListName"                    $MSWSSRESTSurvey
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "TaskListName"                      $MSWSSRESTWorkflowTaskList
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "DiscussionBoardListName"           $MSWSSRESTDiscussionBoard
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "CalendarListName"                  $MSWSSRESTCalendar
	ModifyConfigFileNode $MSWSSRESTDeploymentFile "ChoiceFieldName"                   $MSWSSRESTChoiceFieldName
    ModifyConfigFileNode $MSWSSRESTDeploymentFile "MultiChoiceFieldName"              $MSWSSRESTMultiChoiceFieldName 
	ModifyConfigFileNode $MSWSSRESTDeploymentFile "WorkflowHistoryListName"           $MSWSSRESTWorkflowHistoryList
	ModifyConfigFileNode $MSWSSRESTDeploymentFile "ChoiceFieldOptions"                $choiceFieldOptions
	ModifyConfigFileNode $MSWSSRESTDeploymentFile "MultiChoiceFieldOptions"           $multiChoiceFieldOptions 
    
	Output "Configuration for the MS-WSSREST_TestSuite.deployment.ptfconfig file is complete." "Green"
}
#-----------------------------------------------------
# Configuration for MS-OFFICIALFILE ptfconfig file.
#-----------------------------------------------------
if($sutVersion -eq "SharePointServer2013" -or $sutVersion -eq "SharePointServer2010" -or $sutVersion -eq "SharePointServer2007" -or $sutVersion -eq "SharePointServer2016" -or $sutVersion -eq "SharePointServer2019")
{
    Output "Configure the MS-OFFICIALFILE_TestSuite.deployment.ptfconfig file ..." "White"
    #The urls of site which are used by MS-OFFICIALFILE test suite
	$enableContentOrganizerRecordsCenterSite = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/$MSOFFICIALFILERoutingRepositorySite"
	$enableContentOrganizerRecordsCenterSiteUrl = "${transportType}://$sutComputerName/sites/$MSOFFICIALFILESiteCollectionName/$MSOFFICIALFILERoutingRepositorySite"
	$disableContentOrganizerRecordsCenterSite = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/$MSOFFICIALFILENoRoutingRepositorySite"
	$enableContentOrganizerDocumentsCenterSite = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/$MSOFFICIALFILEEnabledParsingRepositorySite"
	
	#The lists which are used by MS-OFFICIALFILE test suite
	$noEnforceLibraryUrl = "[EnableContentOrganizerRecordsCenterSite]/$MSOFFICIALFILENoEnforceLibrary "
	$documentSetUrl = "[EnableContentOrganizerRecordsCenterSite]/$MSOFFICIALFILEDocumentSetLocationLibrary/$MSOFFICIALFILEDocumentSetName"
	
	#The host info
	$holdInfo = GetHoldInfo $sutComputerName ($dnsDomain.split(".")[0]+ "\" + $userName) $password "$enableContentOrganizerRecordsCenterSiteUrl" "$MSOFFICIALFILEHolds"
    $holdId = $holdInfo[0]
    $holdUrl = $enableContentOrganizerRecordsCenterSiteUrl + "/" + $holdInfo[1]
	
	Output "Modify the properties as necessary in the MS-OFFICIALFILE_TestSuite.deployment.ptfconfig file..." "White"
	$step = 1
    Output "Steps for manual configuration:" "Yellow"
    Output "$step.Open $MSOFFICIALFILEDeploymentFile" "Yellow"
	$step++
    Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSOFFICIALFILESiteCollectionName." "Yellow"
	$step++
    Output "$step.Find the property `"EnableContentOrganizerRecordsCenterSite`", and set the value to $enableContentOrganizerRecordsCenterSite." "Yellow"
	$step++
    Output "$step.Find the property `"DisableContentOrganizerRecordsCenterSite`", and set the value to $disableContentOrganizerRecordsCenterSite." "Yellow"
	$step++
    Output "$step.Find the property `"EnableContentOrganizerDocumentsCenterSite`", and set the value to $enableContentOrganizerDocumentsCenterSite." "Yellow"
	$step++
    Output "$step.Find the property `"NoEnforceLibraryUrl`", and set the value to $noEnforceLibraryUrl." "Yellow"
	$step++
    Output "$step.Find the property `"DocumentSetUrl`", and set the value to $documentSetUrl." "Yellow"
	$step++
    Output "$step.Find the property `"NoRecordsCenterSubmittersPermissionUserName`", and set the value to $MSOFFICIALFILEReadUser." "Yellow"
	$step++
    Output "$step.Find the property `"NoRecordsCenterSubmittersPermissionPassword`", and set the value to $MSOFFICIALFILEReadUserPassword." "Yellow"
	$step++
    Output "$step.Find the property `"HoldName`", and set the value to $MSOFFICIALFILEHolds." "Yellow"
	$step++
    Output "$step.Find the property `"HoldUrl`", and set the value to $holdUrl." "Yellow"
	$step++
    Output "$step.Find the property `"HoldId`", and set the value to $holdId." "Yellow"
	$step++
    Output "$step.Find the property `"HoldSearchQuery`", and set the value to $sutComputerName." "Yellow"
	$step++
    Output "$step.Find the property `"HoldSearchContextUrl`", and set the value to $enableContentOrganizerRecordsCenterSite." "Yellow"
	$step++
    Output "$step.Find the property `"DocumentLibraryName`", and set the value to $MSOFFICIALFILEDocumentRuleLocationLibrary." "Yellow"
	$step++
    Output "$step.Find the property `"DefaultLibraryName`", and set the value to $MSOFFICIALFILEDropOffLibrary." "Yellow"
	$step++
    Output "$step.Find the property `"HoldSearchContextUrl`", and set the value to http://$sutComputerName." "Yellow"
    
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "SiteCollectionName"                                $MSOFFICIALFILESiteCollectionName
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "EnableContentOrganizerRecordsCenterSite"           $enableContentOrganizerRecordsCenterSite
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "DisableContentOrganizerRecordsCenterSite"          $disableContentOrganizerRecordsCenterSite
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "EnableContentOrganizerDocumentsCenterSite"         $enableContentOrganizerDocumentsCenterSite
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "NoEnforceLibraryUrl"                               $noEnforceLibraryUrl
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "DocumentSetUrl"                                    $documentSetUrl
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "NoRecordsCenterSubmittersPermissionUserName"       $MSOFFICIALFILEReadUser
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "NoRecordsCenterSubmittersPermissionPassword"       $MSOFFICIALFILEReadUserPassword
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldName"                                          $MSOFFICIALFILEHolds
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldUrl"                                           $holdUrl
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldId"                                            $holdId
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldSearchQuery"                                   $sutComputerName
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldSearchContextUrl"                              $enableContentOrganizerRecordsCenterSite
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "DocumentLibraryName"                               $MSOFFICIALFILEDocumentRuleLocationLibrary
	ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "DefaultLibraryName"                                $MSOFFICIALFILEDropOffLibrary
    ModifyConfigFileNode $MSOFFICIALFILEDeploymentFile "HoldSearchContextUrl"                              "http://$sutComputerName"
	
	Output "Configuration for the MS-OFFICIALFILE_TestSuite.deployment.ptfconfig file is complete." "Green"	
}

#-----------------------------------------------------
# Configuration for MS-COPYS ptfconfig file.
#-----------------------------------------------------

Output "Configure the MS-COPYS_TestSuite.deployment.ptfconfig file ..." "White"

$sourceFileUrlOnSourceSUT = "[TransportType]://[SourceSutComputerName]/sites/[SiteCollectionName]/[SourceDocLibraryName]/$MSCOPYSTestData"
$sourceFileUrlOnDesSUT = "[TransportType]://[SutComputerName]/sites/[SiteCollectionName]/[SourceDocLibraryName]/$MSCOPYSTestData"
Output "Modify the properties as necessary in the MS-COPYS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MCOPYSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SourceSutComputerName`", and set the value to $sut2ComputerName." "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSCOPYSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"MeetingWorkSpaceSiteName`", and set the value to $MSCOPYSSubSite." "Yellow"
$step++
Output "$step.Find the property `"SourceDocLibraryName`", and set the value to $MSCOPYSSourceDocumentLibrary." "Yellow"
$step++
Output "$step.Find the property `"DestinationDocLibrary`", and set the value to $MSCOPYSDestinationDocumentLibrary." "Yellow"
$step++
Output "$step.Find the property `"SourceFileContents`", and set the value to $MSCOPYSTestContent." "Yellow"
$step++
Output "$step.Find the property `"SourceFileUrlOnSourceSUT`", and set the value to $sourceFileUrlOnSourceSUT." "Yellow"
$step++
Output "$step.Find the property `"SourceFileUrlOnDesSUT`", and set the value to $sourceFileUrlOnDesSUT." "Yellow"
$step++
Output "$step.Find the property `"FieldNameOfTestReadOnly`", and set the value to $MSCOPYSTextFieldName." "Yellow"
$step++
Output "$step.Find the property `"FieldDefaultValueOfTestReadOnlyOnSourceDocLibrary`", and set the value to $MSCOPYSSourceLibraryFieldValue." "Yellow"
$step++
Output "$step.Find the property `"FieldDefaultValueOfTestReadOnlyOnDesDocLibrary`", and set the value to $MSCOPYSDestinationLibraryFieldValue." "Yellow"
$step++
Output "$step.Find the property `"FieldNameOfTestWorkflowEventType`", and set the value to $MSCOPYSWorkFlowEventFieldName." "Yellow"
$step++
Output "$step.Find the property `"MSCOPYSCheckOutUserName`", and set the value to $MSCOPYSEditUser." "Yellow"
$step++
Output "$step.Find the property `"PasswordOfCheckOutUser`", and set the value to $MSCOPYSEditUserPassword." "Yellow"
$step++
Output "$step.Find the property `"MSCOPYSNoPermissionUser`", and set the value to $MSCOPYSNoPermissionUser." "Yellow"
$step++
Output "$step.Find the property `"PasswordOfNoPermissionUser`", and set the value to $MSCOPYSNoPermissionUserPassword." "Yellow"
$step++
Output "$step.Find the property `"MeetingWorkSpaceDocLibrary`", and set the value to $MSCOPYSSubSiteDocumentLibrary." "Yellow"

ModifyConfigFileNode $MCOPYSDeploymentFile "SourceSutComputerName"                                           $sut2ComputerName
ModifyConfigFileNode $MCOPYSDeploymentFile "SiteCollectionName"                                              $MSCOPYSSiteCollectionName
ModifyConfigFileNode $MCOPYSDeploymentFile "MeetingWorkSpaceSiteName"                                        $MSCOPYSSubSite
ModifyConfigFileNode $MCOPYSDeploymentFile "SourceDocLibraryName"                                            $MSCOPYSSourceDocumentLibrary
ModifyConfigFileNode $MCOPYSDeploymentFile "DestinationDocLibrary"                                           $MSCOPYSDestinationDocumentLibrary
ModifyConfigFileNode $MCOPYSDeploymentFile "SourceFileContents"                                              $MSCOPYSTestContent
ModifyConfigFileNode $MCOPYSDeploymentFile "SourceFileUrlOnSourceSUT"                                        $sourceFileUrlOnSourceSUT
ModifyConfigFileNode $MCOPYSDeploymentFile "SourceFileUrlOnDesSUT"                                           $sourceFileUrlOnDesSUT
ModifyConfigFileNode $MCOPYSDeploymentFile "FieldNameOfTestReadOnly"                                         $MSCOPYSTextFieldName
ModifyConfigFileNode $MCOPYSDeploymentFile "FieldDefaultValueOfTestReadOnlyOnSourceDocLibrary"               $MSCOPYSSourceLibraryFieldValue
ModifyConfigFileNode $MCOPYSDeploymentFile "FieldDefaultValueOfTestReadOnlyOnDesDocLibrary"                  $MSCOPYSDestinationLibraryFieldValue
ModifyConfigFileNode $MCOPYSDeploymentFile "FieldNameOfTestWorkflowEventType"                                $MSCOPYSWorkFlowEventFieldName
ModifyConfigFileNode $MCOPYSDeploymentFile "MSCOPYSCheckOutUserName"                                         $MSCOPYSEditUser
ModifyConfigFileNode $MCOPYSDeploymentFile "PasswordOfCheckOutUser"                                          $MSCOPYSEditUserPassword
ModifyConfigFileNode $MCOPYSDeploymentFile "MSCOPYSNoPermissionUser"                                         $MSCOPYSNoPermissionUser
ModifyConfigFileNode $MCOPYSDeploymentFile "PasswordOfNoPermissionUser"                                      $MSCOPYSNoPermissionUserPassword
ModifyConfigFileNode $MCOPYSDeploymentFile "MeetingWorkSpaceDocLibrary"                                      $MSCOPYSSubSiteDocumentLibrary

Output "Configuration for the MS-COPYS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-VIEWSS ptfconfig file.
#-----------------------------------------------------

Output "Configure the MS-VIEWSS_TestSuite.deployment.ptfconfig file ..." "White"

Output "Modify the properties as necessary in the MS-VIEWSS_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSVIEWSSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSVIEWSSSiteCollectionName." "Yellow"
$step++
Output "$step.Find the property `"DisplayListName`", and set the value to $MSVIEWSSViewListName." "Yellow"

ModifyConfigFileNode $MSVIEWSSDeploymentFile "SiteCollectionName"                                   $MSVIEWSSSiteCollectionName
ModifyConfigFileNode $MSVIEWSSDeploymentFile "DisplayListName"                                      $MSVIEWSSViewListName

Output "Configuration for the MS-VIEWSS_TestSuite.deployment.ptfconfig file is complete." "Green"

#----------------------------------------------------------------------------
# End script
#----------------------------------------------------------------------------
Output "[SharePointClientConfiguration.ps1] has run sucessfully." "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0