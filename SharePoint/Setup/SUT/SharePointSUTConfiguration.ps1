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
# Starting script
#----------------------------------------------------------------------------
$script:ErrorActionPreference = "Stop"
[String]$containerPath = Get-Location
[String]$logPath       = $containerPath + "\SetupLogs"
[String]$logFile       = $logPath + "\SharePointSUTConfiguration.ps1.log"
[String]$debugLogFile  = $logPath + "\SharePointSUTConfiguration.ps1.debug.log"
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

$ExecutionTimeout                            = ReadConfigFileNode "$environmentResourceFile" "ExecutionTimeout"

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
# Check whether the unattended client configuration XML is available if run in unattended mode.
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
            Output "The client configuration XML path `"$unattendedXmlName`" is not correct." "Yellow"
            Output "Retry with the correct file path or press `"Enter`" if you want the SUT setup script to run in attended mode." "Cyan"
            $unattendedXmlName = Read-Host
        }
    }
}

#----------------------------------------------------------------------------
# Start to automatic services required by test case
#----------------------------------------------------------------------------
iisreset /restart
StartService "MSSQL*" "Auto"

#----------------------------------------------------------------------------
# Try to get the SharePoint Server Version
#----------------------------------------------------------------------------
Output "Trying to get the SharePoint Server Version ..." "White"
$SharePointVersionInfo = GetSharePointVersion
$SharePointVersion = $SharePointVersionInfo[0]
if($SharePointVersion  -eq "Unknown Version")
{
    Write-Warning "Could not find supported SharePoint Server installation on the system! Install it first and then re-run this SUT configuration script. `r"
    Stop-Transcript
    exit 2
}
else
{
    OutPutSupportVersionInfo
}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
[void][System.Reflection.Assembly]::Loadwithpartialname("Microsoft.Office.Policy")

$product = "14.0" 
if($SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0]) 
{ 
	$product = "15.0" 
}elseif($SharePointVersion -eq $WindowsSharePointServices3[0] -or $SharePointVersion -eq $SharePointServer2007[0]) 
{ 
	$product = "12.0" 
}elseif ($SharePointVersion -eq $SharePointServer2016[0])
{
    $product = "16.0" 
}
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
	$SharePointShellSnapIn = Get-PSSnapin | Where-Object -FilterScript {$_.Name -eq "Microsoft.SharePoint.PowerShell"}
	if($SharePointShellSnapIn -eq $null)
	{
		Add-PSSnapin Microsoft.SharePoint.PowerShell
	}
}

#----------------------------------------------------------------------------
# Remove WebDAV Publishing role service
#----------------------------------------------------------------------------
if($product -eq "12.0")
{
   $os = Get-WmiObject -class Win32_OperatingSystem -computerName $env:COMPUTERNAME
   if([int]$os.BuildNumber -le 7601)
   {
        Import-Module Servermanager
   }
   $roleStatus= Get-WindowsFeature Web-DAV-Publishing 
   if($roleStatus.Installed -eq "true")
   {
      Output "The WebDAV Publishing role service is installed, we need remove it." "White"
      $roleRemovedStatus=Remove-WindowsFeature Web-DAV-Publishing
      if($roleRemovedStatus.RestartNeeded -eq "Yes")
      {   
            $locationPath = (Get-Location).Path
            Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "CMD" -Value "$psHome\powershell.exe  `"& '$locationPath\SharePointSUTConfiguration.cmd' '$unattendedXmlName'`""
            if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
            {    
                Output "The WebDAV Publishing role service is removed, but a system restart will be required, so press Enter whenever you are ready." "Cyan"
                Output "After the restart, log on to the server, and the setup configuration script will continue to run automatically." "Cyan"
                cmd /c pause 
            }
            shutdown -r -f -t 0
       } 
   }
   
}

#-----------------------------------------------------
# Start to configure server.
#-----------------------------------------------------
Output "Start to configure the server ..." "White"

Output "Steps for manual configuration:" "Yellow" 
Output "Enable remoting in Powershell." "Yellow"
Invoke-Command {
    $ErrorActionPreference = "Continue"
    Enable-PSRemoting -Force
}

[int]$recommendedMaxMemory = 1024
Output "Steps for manual configuration:" "Yellow" 
Output "Ensure that the maximum amount of memory allocated per shell for remote shell management is at least $recommendedMaxMemory MB." "Yellow"
[int]$originalMaxMemory = (Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).Value
if($originalMaxMemory -lt $recommendedMaxMemory)
{
    Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB $recommendedMaxMemory
    $actualMaxMemory = (Get-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB).Value
    Output "The maximum amount of memory allocated per shell for remote shell management increased from $originalMaxMemory MB to $actualMaxMemory MB." "White"
}
else
{
    Output "The maximum amount of memory allocated per shell for remote shell management is $originalMaxMemory MB." "White"
}

#-----------------------------------------------------
# Get SharePoint server basic information.
#-----------------------------------------------------
Output "The basic information of the SharePoint server:" "White"

$domain            = $Env:USERDNSDOMAIN
Output "Domain name: $domain" "White"
$sutComputerName   = $ENV:ComputerName
Output "The SharePoint server name: $sutComputerName" "White"
$userName          = $ENV:UserName
Output "The logon name of the current user: $userName " "White"

Output "Enter the password of the current user:" "Cyan"
$password = CheckForEmptyUserInput "Password" "password"
Output "The password you entered: $password" "Yellow"

#----------------------------------------------------------------------------
# Start to configure SharePoint SUT to support HTTPS transport.
#----------------------------------------------------------------------------
Output "Configure the HTTPS service in the SharePoint site." "White"
Output "Steps for manual configuration" "Yellow"
Output "1. Configure the SUT to support HTTPS." "Yellow"
Output "2. Set an alternate access mapping for HTTPS." "Yellow"
$webAppName = GetWebAPPName
AddHTTPSBinding "$sutComputerName" $SharePointVersion $webAppName

#----------------------------------------------------------------------------
# Start to update the executionTimeout attribute of httpRuntime element in web.config file of SharePoint Central Administration
#----------------------------------------------------------------------------
Output "Start to update the executionTimeout attribute of httpRuntime element in web.config file of SharePoint Central Administration." "White"
Output "Steps for manual configuration:" "Yellow"
Output "1. Find the web.config file of SharePoint Central Administration." "Yellow"
Output "2. Update the executionTimeout attribute of httpRuntime element." "Yellow"
# Get the port number of SharePoint central administration
$adminUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.sites[0].url
if($adminUrl -eq $null -or $adminUrl -eq "")
{
    Throw "Cannot get the SPAdministrationWebApplication url."
}
$adminUrlArray = $adminUrl.Split(":")
$adminNumber = $adminUrlArray[2]

# Get the path of SharePoint Central Administration V4 web.config
$webConfigPath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${sutComputerName}:${adminNumber}").IisSettings.item(0).Path.ToString()
if($webConfigPath -eq $null -or $webConfigPath -eq "")
{
    Throw "Cannot get the path of webconfig."
}
$webConfigfile = "$webConfigPath\web.config"
if(!([System.IO.File]::Exists($webConfigfile)))
{
    Throw "The webconfig file $webConfigfile does not exist."
}

# Update the attribute of “httpRuntime” element
$xml=[xml](Get-Content $webConfigfile)
$root=$xml.get_DocumentElement()
$root."system.web".httpRuntime.SetAttribute("executionTimeout", $ExecutionTimeout)
$xml.Save($webConfigfile)

#----------------------------------------------------------------------------
# Add the user policy to enable the client side PowerShell scripts 
# to manage the SharePoint remotely.
#----------------------------------------------------------------------------
$webApplicationUrl = "http://$sutComputerName"
$uri = new-object System.Uri($webApplicationUrl)
$webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($uri)
$useClaims = $webApp.UseClaimsAuthentication
if($useClaims)
{
    Output "Steps for manual configuration:" "Yellow"
    Output "Add the user policy for $userName with ""Full Control"" permission without name prefixed." "Yellow"
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $userName)
}

#-----------------------------------------------------
# Start to configure SUT for MS-SHDACCWS.
#-----------------------------------------------------
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
	Output "Start to configure MS-SHDACCWS." "White"

	Output "Steps for manual configuration:" "Yellow"
	Output "Create a site collection named $MSSHDACCWSSiteCollectionName ..." "Yellow"
	$MSSHDACCWSSiteCollectionObject = CreateSiteCollection $MSSHDACCWSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

	Output "Steps for manual configuration:" "Yellow"
	Output "Create the document library $MSSHDACCWSDocumentLibrary under the root web of the site collection $MSSHDACCWSSiteCollectionName ..." "Yellow"
	CreateListItem $MSSHDACCWSSiteCollectionObject.RootWeb $MSSHDACCWSDocumentLibrary 101

	Output "Steps for manual configuration:" "Yellow"
	Output "Upload test data $MSSHDACCWSTestData to $MSSHDACCWSDocumentLibrary under $MSSHDACCWSSiteCollectionName ..." "Yellow"
	UploadFileToSharePointFolder $MSSHDACCWSSiteCollectionObject.RootWeb $MSSHDACCWSDocumentLibrary $MSSHDACCWSTestData ".\$MSSHDACCWSTestData"  $True

	Output "Steps for manual configuration:" "Yellow"
	Output "Upload test data $MSSHDACCWSCoStatusTestData to $MSSHDACCWSDocumentLibrary under $MSSHDACCWSSiteCollectionName ..." "Yellow"
	UploadFileToSharePointFolder $MSSHDACCWSSiteCollectionObject.RootWeb $MSSHDACCWSDocumentLibrary $MSSHDACCWSCoStatusTestData ".\$MSSHDACCWSCoStatusTestData"  $True

	Output "Steps for manual configuration:" "Yellow"
	Output "Upload test data $MSSHDACCWSLockedTestData to $MSSHDACCWSDocumentLibrary under $MSSHDACCWSSiteCollectionName ..." "Yellow"
	UploadFileToSharePointFolder $MSSHDACCWSSiteCollectionObject.RootWeb $MSSHDACCWSDocumentLibrary $MSSHDACCWSLockedTestData ".\$MSSHDACCWSLockedTestData"  $True

	$MSSHDACCWSSiteCollectionObject.Dispose()
}
#-----------------------------------------------------
# Start to configure SUT for MS-SITESS.
#-----------------------------------------------------
Output "Start to run configurations of MS-SITESS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSSITESSSiteCollectionName ..." "Yellow"
$MSSITESSSiteCollectionObject  = CreateSiteCollection $MSSITESSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

# Activate the workflows feature for SharePoint Server 2010 and SharePoint Server 2013 and SharePoint Server 2016.
if ($SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
    Output "Steps for manual configuration:" "Yellow"
    Output "Active the workflows feature on site collection feature page ..." "Yellow"
    SetWebFeature  "http://$sutComputerName/sites/$MSSITESSSiteCollectionName" "Workflows"
}

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite named $MSSITESSSite under site collection $MSSITESSSiteCollectionName ..." "Yellow"
$MSSITESSSiteObject = CreateWeb $MSSITESSSiteCollectionObject $false $MSSITESSSite $MSSITESSSite "STS#2" $true
$MSSITESSSiteObject.Dispose()

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite of $MSSITESSSite named $MSSITESSNormalSubSite ..." "Yellow"
$MSSITESSNormalSubSiteObject = CreateWeb $MSSITESSSiteCollectionObject $true ($MSSITESSSite + "/" + $MSSITESSNormalSubSite)
$MSSITESSNormalSubSiteObject.Dispose()

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite of $MSSITESSSite named $MSSITESSSpecialSubSite ..." "Yellow"
$MSSITESSSpecialSubSiteObject = CreateWeb $MSSITESSSiteCollectionObject $true ($MSSITESSSite + "/" + $MSSITESSSpecialSubSite)

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSSITESSDocumentLibrary under the root web of the site collection $MSSITESSSiteCollectionName ..." "Yellow"
CreateListItem $MSSITESSSiteCollectionObject.RootWeb $MSSITESSDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Create the docment library $MSSITESSSubSiteDocumentLibrary under $MSSITESSSpecialSubSite ..." "Yellow"
CreateListItem $MSSITESSSpecialSubSiteObject $MSSITESSSubSiteDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Create a 24M size txt file $MSSITESSTestData."
CreateFile $MSSITESSTestData 24mb $containerPath

Output "Steps for manual configuration:" "Yellow"
Output "Upload a 24M size txt file $MSSITESSTestData to $MSSITESSSubSiteDocumentLibrary under $MSSITESSSpecialSubSite ..." "Yellow"
UploadFileToSharePointFolder $MSSITESSSpecialSubSiteObject $MSSITESSSubSiteDocumentLibrary $MSSITESSTestData ".\$MSSITESSTestData"  $True

Output "Steps for manual configuration:" "Yellow"
Output "Upload a user custom page $MSSITESSCustomPage to $MSSITESSSubSiteDocumentLibrary under $MSSITESSSpecialSubSite ..." "Yellow"
UploadFileToSharePointFolder $MSSITESSSpecialSubSiteObject $MSSITESSSubSiteDocumentLibrary $MSSITESSCustomPage ".\$MSSITESSCustomPage"  $True

$MSSITESSSiteCollectionObject.Dispose()
$MSSITESSSpecialSubSiteObject.Dispose()

Output "Enable the user custom pages running on the server side." "White"
Output "Steps for manual configuration:" "Yellow"
Output "Add the following content to PageParserPaths element in Web.Config file ..." "Yellow"
Output "<PageParserPath VirtualPath=""/*"" CompilationMode=""Always"" AllowServerSideScript=""true"" IncludeSubFolders=""true"" />" "Yellow"
EnableServerSideScriptForCustomPages $MSSITESSSiteCollectionObject

#-----------------------------------------------------
# Start to configure SUT for MS-DWSS.
#-----------------------------------------------------
Output "Start to run configurations of MS-DWSS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create three users on the domain controller:" "Yellow"
Output "Name:$MSDWSSNoneRoleUser Password:$MSDWSSNoneRoleUserPassword" "Yellow"
Output "Name:$MSDWSSReaderRoleUser Password:$MSDWSSReaderRoleUserPassword" "Yellow"
Output "Name:$MSDWSSGroupOwner Password:$MSDWSSGroupOwnerPassword" "Yellow"
Output "Steps for manual configuration:" "Yellow"
Output "1. Open Active Directory Users and Computers..." "Yellow"
Output "2. Create three new users with the name of MSDWSS_NoneRoleUser, MSDWSS_ReaderRole and MSDWSS_GroupOwner with the password mentioned above..." "Yellow"
CreateUserOnDC $MSDWSSNoneRoleUser $MSDWSSNoneRoleUserPassword
CreateUserOnDC $MSDWSSReaderRoleUser $MSDWSSReaderRoleUserPassword
CreateUserOnDC $MSDWSSGroupOwner $MSDWSSGroupOwnerPassword

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSDWSSSiteCollectionName ..." "Yellow"
$MSDWSSSiteCollectionObject = CreateSiteCollection $MSDWSSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSDWSSSiteDocumentWorkpace with the document workspace site template, and the site collection has no any subsites ..." "Yellow"
CreateSiteCollection $MSDWSSSiteDocumentWorkpace $sutComputerName "$domain\$userName" "$userName@$domain" "STS#2" 1033

Output "Steps for manual configuration:" "Yellow"
Output "Update user email under site collection $MSDWSSSiteCollectionObject..." "Yellow"
ConfigSPUserEmail ($domain.split(".")[0] + "\" + $userName) "$userName@$domain" $MSDWSSSiteCollectionObject

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite named $MSDWSSSite under site collection $MSDWSSSiteCollectionName ..." "Yellow"
$MSDWSSSiteObject = CreateWeb $MSDWSSSiteCollectionObject $false $MSDWSSSite $MSDWSSSite "STS#2" $true

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite named $MSDWSSInheritPermissionSite under site collection $MSDWSSSiteCollectionName ..." "Yellow"
$MSDWSSInheritPermissionSiteObject = CreateWeb $MSDWSSSiteCollectionObject $false $MSDWSSInheritPermissionSite $MSDWSSInheritPermissionSite "STS#2" $false
$MSDWSSInheritPermissionSiteObject.Dispose()

Output "Steps for manual configuration:" "Yellow"
Output "Create the docment library $MSDWSSDocumentLibrary under $MSDWSSSite ..." "Yellow"
CreateListItem $MSDWSSSiteObject $MSDWSSDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Create a folder under $MSDWSSSite named $MSDWSSDocumentLibrary/$MSDWSSTestFolder ..." "Yellow"
CreateSharePointFolder $MSDWSSSiteObject "$MSDWSSDocumentLibrary/$MSDWSSTestFolder"

Output "Steps for manual configuration:" "Yellow"
Output "Upload $MSDWSSTestData to http://$sutComputerName/sites/$MSDWSSSiteCollectionName/$MSDWSSSite/$MSDWSSDocumentLibrary/$MSDWSSTestFolder ..." "Yellow"
UploadFileToSharePointFolder $MSDWSSSiteObject "$MSDWSSDocumentLibrary/$MSDWSSTestFolder" $MSDWSSTestData ".\$MSDWSSTestData" $true

Output "Steps for manual configuration:" "Yellow"
Output "Grant read permission level to $domain\$MSDWSSReaderRoleUser on site $MSDWSSSite..." "Yellow"
GrantUserPermission $MSDWSSSiteObject "Read" $domain.Split(".")[0] $MSDWSSReaderRoleUser

Output "Steps for manual configuration:" "Yellow"
Output "Create a group on the site collection $MSDWSSSiteCollectionName, the group name is $MSDWSSGroupName, and owner is $MSDWSSGroupOwner ..." "Yellow"
AddGroupForSiteCollection $MSDWSSSiteCollectionObject ($domain.split(".")[0]+ "\$MSDWSSGroupOwner") $MSDWSSGroupName

Output "Steps for manual configuration:" "Yellow"
Output "Grant the group $MSDWSSGroupName with full control permission level to $MSDWSSSite..." "Yellow"
GrantGroupPermission $MSDWSSSiteObject "Full Control" $MSDWSSGroupName

$MSDWSSSiteCollectionObject.Dispose()
$MSDWSSSiteObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-VERSS.
#-----------------------------------------------------
Output "Start to run configurations of MS-VERSS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSVERSSSiteCollectionName ..." "Yellow"
$MSVERSSSiteCollectionObject = CreateSiteCollection $MSVERSSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
$MSVERSSSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-LISTSWS.
#-----------------------------------------------------
Output "Start to run configurations of MS-LISTSWS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSLISTSWSSiteCollectionName ..." "Yellow"
$MSLISTSWSSiteCollectionObject = CreateSiteCollection $MSLISTSWSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

Output "Steps for manual configuration:" "Yellow"
Output "Create the docment library $MSLISTSWSDocumentLibrary under the root web of the site collection $MSLISTSWSSiteCollectionName ..." "Yellow"
CreateListItem $MSLISTSWSSiteCollectionObject.RootWeb $MSLISTSWSDocumentLibrary 101
$MSLISTSWSSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-ADMINS.
#-----------------------------------------------------
Output "Start to run configurations of MS-ADMINS..." "White"

#The default name of the SharePoint Central Administration.
$defaultWebAppName = "SharePoint Central Administration v4"

if($SharePointVersion -eq $WindowsSharePointServices3[0] -or $SharePointVersion -eq $SharePointServer2007[0])
{
    $defaultWebAppName = "SharePoint Central Administration v3"
}

Output "Steps for manual configuration:" "Yellow"
Output "Configure HTTPS service in the SharePoint Central Administration." "Yellow"
Output "1. Configure the SUT to support HTTPS." "Yellow"
Output "2. Set an alternate access mapping for HTTPS." "Yellow"
$WebApplicationName = GetWebAPPName $defaultWebAppName
AddHTTPSBinding "$sutComputerName" $SharePointVersion $WebApplicationName $httpsPortNumberOnAdminSite $true

# Activate the feature DocumentManagement and DocumentSet for SharePoint Server 2013.
if($SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
    $siteUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.sites[0].Url
	
	Output "Steps for manual configuration:" "Yellow"
    Output "Active the feature DocumentManagement and DocumentSet on Manage Web Application Features Page ..." "Yellow"
	SetWebFeature $siteUrl "DocumentManagement"
	SetWebFeature $siteUrl "DocumentSet"
}

# Add a firewall rule to allow local port of admin application to receive TCP data.
$adminLocalPorts = $httpsPortNumberOnAdminSite
if($SharePointVersion -eq $SharePointServer2007[0] -or $SharePointVersion -eq $WindowsSharePointServices3[0])
{
    # Get the central administration port number.
    $adminLocalPorts = "$httpsPortNumberOnAdminSite,$adminNumber"
}
AddFirewallInboundRule  "Enable admin site port number" "TCP" $adminLocalPorts $true

#-----------------------------------------------------
# Start to configure SUT for MS-MEETS.
#-----------------------------------------------------
Output "Start to run configurations of MS-MEETS..." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a user on the domain controller:" "Yellow"
Output "Name:$MSMEETSUser Password:$MSMEETSUserPassword" "Yellow"
Output "1. Open Active Directory Users and Computers..." "Yellow"
Output "2. Create a new user with the name of $MSMEETSUser with the password mentioned above..." "Yellow"
CreateUserOnDC $MSMEETSUser $MSMEETSUserPassword

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSMEETSSiteCollectionName ..." "Yellow"
$MSMEETSSiteCollectionObject = CreateSiteCollection $MSMEETSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
$MSMEETSSiteCollectionObject.Dispose()

if($SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
    $webTempfileName = "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\15\TEMPLATE\1033\XML\WEBTEMP.XML"
	$meetingTemplates = "Basic Meeting Workspace","Blank Meeting Workspace","Decision Meeting Workspace","Social Meeting Workspace","Multipage Meeting Workspace"
	
    if(Test-Path $webTempfileName)
    {
        Output "Change the hidden value to false for the meeting template" "White"
		$step = 1
        Output "Steps for manual configuration:" "Yellow"
        Output "$step.Open $webTempfileName ..." "Yellow"
		$step++
        Output "$step.Find meeting template `"Basic Meeting Workspace`", set the hidden value to false" "Yellow"
		$step++
        Output "$step.Find meeting template `"Blank Meeting Workspace`", set the hidden value to false" "Yellow"
		$step++
        Output "$step.Find meeting template `"Decision Meeting Workspace`", set the hidden value to false" "Yellow"
		$step++
        Output "$step.Find meeting template `"Social Meeting Workspace`", set the hidden value to false" "Yellow"
		$step++
        Output "$step.Find meeting template `"Multipage Meeting Workspace`", set the hidden value to false" "Yellow"
		
		foreach($meetingTemplate in $meetingTemplates)
		{
		    ModifyXMLFileNode $webTempfileName $meetingTemplate "FALSE" "Configuration" "Title" "Hidden"
		}
    }
}

Output "Restart IIS ..." "White"
iisreset /restart

#-----------------------------------------------------
# Start to configure SUT for MS-WWSP.
#-----------------------------------------------------

if($SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointServer2007[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
    Output "Start to run configurations of MS-WWSP." "White"

    Output "Steps for manual configuration:" "Yellow"
	Output "Create a group on the domain controller:" "Yellow"
    Output "GroupName:$MSWWSPUserGroupName." "Yellow"
    Output "1. Open Active Directory Users and Computers..." "Yellow"
    Output "2. Create a group with the name of $MSWWSPUserGroupName." "Yellow"
    Invoke-Command{
	$ErrorActionPreference = "Continue"
	cmd /c net GROUP /domain $MSWWSPUserGroupName /add 2>&1 | Out-Null
	}

	Output "Steps for manual configuration:" "Yellow"
	Output "Create a user on the domain controller:" "Yellow"
    Output "Name:$MSWWSPUser Password:$MSWWSPUserPassword" "Yellow"
    Output "1. Open Active Directory Users and Computers..." "Yellow"
    Output "2. Create a new user with the name of $MSWWSPUser with the password mentioned above..." "Yellow"
    CreateUserOnDC $MSWWSPUser $MSWWSPUserPassword
	
    Output "Steps for manual configuration:" "Yellow"
	Output "Add users to a group on the domain controller:" "Yellow"
    Output "1. Open Active Directory Users and Computers..." "Yellow"
    Output "2. Add users with the name of $userName and $MSWWSPUser to the group $MSWWSPUserGroupName." "Yellow"
	Invoke-Command{
	$ErrorActionPreference = "Continue"
	cmd /c net GROUP /domain $MSWWSPUserGroupName $userName /add 2>&1 | Out-Null
	cmd /c net GROUP /domain $MSWWSPUserGroupName $MSWWSPUser /add 2>&1 | Out-Null
	}

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a site collection named $MSWWSPSiteCollectionName ..." "Yellow"
    $MSWWSPSiteCollectionObject = CreateSiteCollection $MSWWSPSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
        	
	Output "Steps for manual configuration:" "Yellow"
    Output "Create the document library $MSWWSPDocumentLibrary under the root web of the site collection $MSSITESSSiteCollectionName ..." "Yellow"
    CreateListItem $MSWWSPSiteCollectionObject.RootWeb $MSWWSPDocumentLibrary 101

    # Activate the workflows feature for SharePoint Server 2010 and SharePoint Server 2013.
    if ($SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
    {
        Output "Steps for manual configuration:" "Yellow"
        Output "Active the workflows feature on site collection feature page ..." "Yellow"
        SetWebFeature "http://$sutComputerName/sites/$MSWWSPSiteCollectionName" "Workflows"
    }

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a Workflow history list $MSWWSPWorkflowHistoryList under the root web of the site collection $MSWWSPSiteCollectionName ..." "Yellow"
    CreateListItem $MSWWSPSiteCollectionObject.RootWeb $MSWWSPWorkflowHistoryList 140

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a task list $MSWWSPWorkflowTaskList under the root web of the site collection $MSWWSPSiteCollectionName ..." "Yellow"
    CreateListItem $MSWWSPSiteCollectionObject.RootWeb $MSWWSPWorkflowTaskList 107

    # The workflow template name is 'Approval' for WindowsSharePointServices3 and SharePointServer2007. 
    # The workflow template name is 'Approval - SharePoint 2010' for SharePointFoundation2010, SharePointServer2010, SharePointFoundation2013 and SharePointServer2013.
    $WorkFlowTemplatename = 'Approval - SharePoint 2010'
    if($SharePointVersion -eq $SharePointServer2007[0] -or $SharePointVersion -eq $WindowsSharePointServices3[0])
    {
        $WorkFlowTemplatename = 'Approval'
    }

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a workflow association with the name of $MSWWSPWorkflowName under specified list $MSWWSPDocumentLibrary ..." "Yellow"
    AddListWorkFlow $MSWWSPSiteCollectionObject $MSWWSPDocumentLibrary $MSWWSPWorkflowName $WorkFlowTemplatename $MSWWSPWorkflowTaskList $MSWWSPWorkflowHistoryList
   	
	# Get the root web that is located at "http://$sutComputerName/sites/$MSWWSPSiteCollectionName".
    $MSWWSP_web = $MSWWSPSiteCollectionObject.OpenWeb()	
	Output "Steps for manual configuration:" "Yellow"
    Output "Grant full control permission level to $domain\$MSWWSPUser on site $MSWWSPSiteCollectionName..." "Yellow"
    GrantUserPermission $MSWWSP_web "Full Control" $domain.Split(".")[0] $MSWWSPUser
	
    $MSWWSPSiteCollectionObject.Dispose()
}

#-----------------------------------------------------
# Start to configure SUT for MS-WEBSS.
#-----------------------------------------------------
Output "Start to run configurations of MS-WEBSS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSWEBSSSiteCollectionName ..." "Yellow"
$MSWEBSSSiteCollectionObject = CreateSiteCollection $MSWEBSSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite named $MSWEBSSSite under site collection $MSWEBSSSiteCollectionName ..." "Yellow"
$MSWEBSSSiteObject = CreateWeb $MSWEBSSSiteCollectionObject $false $MSWEBSSSite $MSWEBSSSiteTitle "STS#0" $false $MSWEBSSSiteDescription 1033

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSWEBSSDocumentLibrary under the root web of the site $MSWEBSSSite ..." "Yellow"
CreateListItem $MSWEBSSSiteObject $MSWEBSSDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Upload test data $MSWEBSSTestData to http://$sutComputerName/sites/$MSWEBSSSiteCollectionObject/$MSWEBSSSite ..." "Yellow"
UploadFileToSharePointFolder $MSWEBSSSiteObject $MSWEBSSDocumentLibrary $MSWEBSSTestData ".\$MSWEBSSTestData" $true

$MSWEBSSSiteObject.Dispose()
$MSWEBSSSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-OUTSPS.
#-----------------------------------------------------
Output "Start to run configurations of MS-OUTSPS." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSOUTSPSSiteCollectionName ..." "Yellow"
$MSOUTSPSSiteCollectionObject = CreateSiteCollection $MSOUTSPSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
$MSOUTSPSSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-WDVMODUU.
#-----------------------------------------------------
Output "Start to run configurations of MS-WDVMODUU." "White"

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSWDVMODUUSiteCollectionName ..." "Yellow"
$MSWDVMODUUSiteCollectionObject = CreateSiteCollection $MSWDVMODUUSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
 
Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSWDVMODUUDocumentLibrary1 under the root web of the site collection $MSWDVMODUUSiteCollectionName ..." "Yellow"
CreateListItem $MSWDVMODUUSiteCollectionObject.RootWeb $MSWDVMODUUDocumentLibrary1 101

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSWDVMODUUDocumentLibrary2 under the root web of the site collection $MSWDVMODUUSiteCollectionName ..." "Yellow"
CreateListItem $MSWDVMODUUSiteCollectionObject.RootWeb $MSWDVMODUUDocumentLibrary2 101

# Get the root web that is located at "http://$sutComputerName/sites/$MSWDVMODUUSiteCollectionName".
$MSWDVMODUU_web = $MSWDVMODUUSiteCollectionObject.OpenWeb()

Output "Steps for manual configuration:" "Yellow"
Output "Create a folder under $MSWDVMODUUSiteCollectionName named $MSWDVMODUUDocumentLibrary1/$MSWDVMODUUTestFolder ..." "Yellow"
CreateSharePointFolder $MSWDVMODUU_web "$MSWDVMODUUDocumentLibrary1/$MSWDVMODUUTestFolder"

Output "Steps for manual configuration:" "Yellow"
Output "Upload a test data $MSWDVMODUUTestData1 to http://$sutComputerName/sites/$MSWDVMODUUSiteCollectionName ..." "Yellow"
UploadFileToSharePointFolder $MSWDVMODUU_web $MSWDVMODUUDocumentLibrary1 $MSWDVMODUUTestData1 ".\$MSWDVMODUUTestData1" $true

Output "Steps for manual configuration:" "Yellow"
Output "Upload test data $MSWDVMODUUTestData2 to http://$sutComputerName/sites/$MSWDVMODUUSiteCollectionName ..." "Yellow"
UploadFileToSharePointFolder $MSWDVMODUU_web $MSWDVMODUUDocumentLibrary1 $MSWDVMODUUTestData2 ".\$MSWDVMODUUTestData2" $true

Output "Steps for manual configuration:" "Yellow"
Output "Upload $MSWDVMODUUTestData3 to http://$sutComputerName/sites/$MSWDVMODUUSiteCollectionObject/$MSWDVMODUUDocumentLibrary1/$MSWDVMODUUTestFolder ..." "Yellow"
UploadFileToSharePointFolder $MSWDVMODUU_web "$MSWDVMODUUDocumentLibrary1/$MSWDVMODUUTestFolder" $MSWDVMODUUTestData3 ".\$MSWDVMODUUTestData3" $true

$MSWDVMODUUSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-AUTHWS.
#-----------------------------------------------------
Output "Start to run configurations of MS-AUTHWS." "White"
	
#Get the path of SecurityToken.
$regkey = "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\$product"
$propertyName = "Location"
$location = (Get-ItemProperty $regkey $propertyName).$propertyName
$secuTokenFilepath = $location.Trim('\') + "\WebServices\SecurityToken\web.config"
	
#Get the active directory connetion string.
$connectionString= GetFQDN
$groupConnectingString = "LDAP://" + $domain + "/" + $connectionString

$rootWebFilePath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://$sutComputerName").iisSettings.Item(0).Path.ToString()
$webFilePath = $rootWebFilePath.SubString(0,$rootWebFilePath.LastIndexOf("\"))

#Create Web Application.
Output "Steps for manual configuration:" "Yellow"
Output "Create three web applications named $MSAUTHWSFormsWebAPPName,$MSAUTHWSPassportWebAPPName and $MSAUTHWSNoneWebAPPName." "Yellow"

if($product -eq "15.0" -or $product -eq "14.0" -or $product -eq "16.0")
{
	$poolAccount = ($domain.split(".")[0] + "\" + $userName)	
	CreateWebApplication $sutComputerName $poolAccount $MSAUTHWSFormsWebAPPPort $MSAUTHWSFormsWebAPPName $MSAUTHWSFormsWebAPPName $password $SharePointVersion $true		
	CreateWebApplication $sutComputerName $poolAccount $MSAUTHWSPassportWebAPPPort $MSAUTHWSPassportWebAPPName $MSAUTHWSPassportWebAPPName $password $SharePointVersion
	CreateWebApplication $sutComputerName $poolAccount $MSAUTHWSNoneWebAPPPort $MSAUTHWSNoneWebAPPName $MSAUTHWSNoneWebAPPName $password $SharePointVersion	$false $true
	     
}elseif($product -eq "12.0")
{ 
	CreateWebApplicationOn2007 $sutComputerName $MSAUTHWSFormsWebAPPPort $MSAUTHWSFormsWebAPPName $webFilePath
	CreateWebApplicationOn2007 $sutComputerName $MSAUTHWSPassportWebAPPPort $MSAUTHWSPassportWebAPPName $webFilePath
	CreateWebApplicationOn2007 $sutComputerName $MSAUTHWSNoneWebAPPPort $MSAUTHWSNoneWebAPPName $webFilePath		
}
#Get the path of web file.
$webFormsFolderPath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${sutComputerName}:${MSAUTHWSFormsWebAPPPort}").IisSettings.item(0).Path.ToString()
$webPassportFormsFolderPath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${sutComputerName}:${MSAUTHWSPassportWebAPPPort}").IisSettings.item(0).Path.ToString()
$webNoneFolderPath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${sutComputerName}:${MSAUTHWSNoneWebAPPPort}").IisSettings.item(0).Path.ToString()

$formWebFilePath = "$webFormsFolderPath\web.config"
$passPortWebFilePath = "$webPassportFormsFolderPath\web.config"
$noneWebFilePath = "$webNoneFolderPath\web.config"

Output "Steps for manual configuration:" "Yellow"
Output "Add membership in file: $formWebFilePath." "Yellow"
SetServerAuthenticationMode "$formWebFilePath" $SharePointVersion $domain "$domain\$userName" $password $connectionString $groupConnectingString "Forms" $false

#Update server authentication mode.
if([System.IO.File]::Exists($secuTokenFilepath))
{   
    SetServerAuthenticationMode $secuTokenFilepath $SharePointVersion $domain "$domain\$userName" $password $connectionString $groupConnectingString "Forms" $false
}

Output "Steps for manual configuration:" "Yellow"
Output "Update authentication mode to PassPort in file: $passPortWebFilePath." "Yellow"
SetServerAuthenticationMode "$passPortWebFilePath" $SharePointVersion $domain "$domain\$userName" $password $connectionString $groupConnectingString "PassPort"

Output "Steps for manual configuration:" "Yellow"
Output "Update authentication mode to None in file: $noneWebFilePath." "Yellow"
SetServerAuthenticationMode "$noneWebFilePath" $SharePointVersion $domain "$domain\$userName" $password $connectionString $groupConnectingString "None"
	
Output "Steps for manual configuration:" "Yellow"
Output "Configure HTTPS service in SUT web-site $MSAUTHWSFormsWebAPPName." "Yellow"
AddHTTPSBinding "$sutComputerName" $SharePointVersion $MSAUTHWSFormsWebAPPName $MSAUTHWSFormsWebAPPHTTPSPort $false $MSAUTHWSFormsWebAPPPort

Output "Steps for manual configuration:" "Yellow"
Output "Configure HTTPS service in SUT web-site $MSAUTHWSPassportWebAPPName." "Yellow"
AddHTTPSBinding "$sutComputerName" $SharePointVersion $MSAUTHWSPassportWebAPPName $MSAUTHWSPassportWebAPPHTTPSPort $false $MSAUTHWSPassportWebAPPPort

Output "Steps for manual configuration:" "Yellow"
Output "Configure HTTPS service in SUT web-site $MSAUTHWSNoneWebAPPName." "Yellow"
AddHTTPSBinding "$sutComputerName" $SharePointVersion $MSAUTHWSNoneWebAPPName $MSAUTHWSNoneWebAPPHTTPSPort $false $MSAUTHWSNoneWebAPPPort

if($product -eq "15.0" -or $product -eq "16.0")
{   
    Output "Steps for manual configuration:" "Yellow"
	Output "Create a web application named $MSAUTHWSWindowsWebAPPName." "Yellow"
	CreateWebApplication $sutComputerName $poolAccount $MSAUTHWSWindowsWebAPPPort $MSAUTHWSWindowsWebAPPName $MSAUTHWSWindowsWebAPPName $password $SharePointVersion
    
	#Get the path of web file.
	$webWindowsFolderPath = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${sutComputerName}:${MSAUTHWSWindowsWebAPPPort}").IisSettings.item(0).Path.ToString()
	$windowsWebFilePath = "$webWindowsFolderPath\web.config"
	
	Output "Steps for manual configuration:" "Yellow"
    Output "Update authentication mode to Windows in file: $windowsWebFilePath." "Yellow"
	SetServerAuthenticationMode "$windowsWebFilePath" $SharePointVersion $domain "$domain\$userName" $password $connectionString $groupConnectingString "Windows"
    
	Output "Steps for manual configuration:" "Yellow"
	Output "Configure HTTPS service in SUT web-site $MSAUTHWSWindowsWebAPPName." "Yellow"
	AddHTTPSBinding "$sutComputerName" $SharePointVersion $MSAUTHWSWindowsWebAPPName $MSAUTHWSWindowsWebAPPHTTPSPort $false $MSAUTHWSWindowsWebAPPPort

}
Output "Restart IIS ..." "White"
iisreset /restart

# Add a firewall rule to allow local port of applications for MS-AUTHWS to receive TCP data.
if($product -eq "12.0" -or $product -eq "14.0")
{
    $authwsLocalPorts= "$MSAUTHWSFormsWebAPPPort,$MSAUTHWSFormsWebAPPHTTPSPort,$MSAUTHWSNoneWebAPPPort,$MSAUTHWSNoneWebAPPHTTPSPort"
}
elseif($product -eq "15.0" -or $product -eq "16.0")
{ 
    $authwsLocalPorts= "$MSAUTHWSFormsWebAPPPort,$MSAUTHWSFormsWebAPPHTTPSPort,$MSAUTHWSNoneWebAPPPort,$MSAUTHWSNoneWebAPPHTTPSPort,$MSAUTHWSWindowsWebAPPPort,$MSAUTHWSWindowsWebAPPHTTPSPort" 
}
AddFirewallInboundRule  "Enable authws site port number" "TCP" $authwsLocalPorts $true

#-----------------------------------------------------
# Start to configure SUT for MS-CPSWS.
#-----------------------------------------------------
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0] )
{
	Output "Start to run configurations of MS-CPSWS." "White"
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create a new user on the domain controller:" "Yellow"
	Output "Name:$MSCPSWSUser Password:$MSCPSWSUserPassword" "Yellow"
	Output "1. Open Active Directory Users and Computers..." "Yellow"
	Output "2. Create a new user with the name of $MSCPSWSUser with the password mentioned above..." "Yellow"
	CreateUserOnDC $MSCPSWSUser $MSCPSWSUserPassword

	Output "Enable anonymous authentication for spclaimproviderwebservice.https.svc and spclaimproviderwebservice.svc in IIS..." "Yellow"
	cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" set config "$webAPPName/_vti_bin/spclaimproviderwebservice.https.svc" /section:system.webServer/security/authentication/AnonymousAuthentication /enabled:true /commit:apphost
	cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" set config "$webAPPName/_vti_bin/spclaimproviderwebservice.svc" /section:system.webServer/security/authentication/AnonymousAuthentication /enabled:true /commit:apphost
	RestartWebApplication $webAPPName
		
	#Get the path of web config file.
	$regkey = "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\$product"
	$propertyName = "Location"
	$location = (Get-ItemProperty $regkey $propertyName).$propertyName
	$webConfigFilepath = "$location" + "ISAPI\web.config"
	ModifyXMLFileSingleNode $webConfigFilepath "ClaimProviderWebServiceBehavior" "true" "behavior" "name" "includeExceptionDetailInFaults" "serviceDebug"
	ModifyXMLFileSingleNode $webConfigFilepath "HttpsClaimProviderWebServiceBehavior" "true" "behavior" "name" "includeExceptionDetailInFaults" "serviceDebug"

}

#-----------------------------------------------------
# Start to configure SUT for MS-WSSREST.
#-----------------------------------------------------
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{
    Output "Start to run configurations of MS-WSSREST." "White"
	Output "Steps for manual configuration:" "Yellow"
    Output "Create a site collection named $MSWSSRESTSiteCollectionName ..." "Yellow"
    $MSWSSRESTSiteCollectionObject = CreateSiteCollection $MSWSSRESTSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

    $ListNames = @{$MSWSSRESTCalendar = 106;$MSWSSRESTDocumentLibrary = 101;$MSWSSRESTDiscussionBoard = 108;$MSWSSRESTGenericList = 100;$MSWSSRESTSurvey = 102;$MSWSSRESTWorkflowHistoryList = 140;$MSWSSRESTWorkflowTaskList = 107}
	foreach($ListName in $ListNames.keys)
	{
	    Output "Steps for manual configuration:" "Yellow"
        Output "Create List $ListName under the root web of the site collection $MSWSSRESTSiteCollectionName ..." "Yellow"
        CreateListItem $MSWSSRESTSiteCollectionObject.RootWeb $ListName $ListNames[$ListName]
	}
    
	#Add Field in list $MSWSSRESTSurvey	
	AddFieldInList $MSWSSRESTSiteCollectionObject.RootWeb $MSWSSRESTSurvey $MSWSSRESTGridChoiceFieldName 16
	AddFieldInList $MSWSSRESTSiteCollectionObject.RootWeb $MSWSSRESTSurvey $MSWSSRESTPageSeparatorFieldName 26
	
    #Add field in list $MSWSSRESTGenericList
	AddFieldInList $MSWSSRESTSiteCollectionObject.RootWeb $MSWSSRESTGenericList $MSWSSRESTChoiceFieldName 6 $true $MSWSSREST_SingleChoiceOptions
	AddFieldInList $MSWSSRESTSiteCollectionObject.RootWeb $MSWSSRESTGenericList $MSWSSRESTMultiChoiceFieldName 15 $true $MSWSSREST_MultiChoiceOptions

    $fieldNames = @{$MSWSSRESTBooleanFieldName = 8;$MSWSSRESTCurrencyFieldName = 10;$MSWSSRESTIntegerFieldName = 1;$MSWSSRESTNumberFieldName = 9;$MSWSSRESTUrlFieldName = 11;$MSWSSRESTWorkFlowEventTypeFieldName = 30}
    foreach($fieldName in $fieldNames.keys)
	{
	    AddFieldInList $MSWSSRESTSiteCollectionObject.RootWeb $MSWSSRESTGenericList $fieldName $fieldNames[$fieldName]
	}
	
	#Add Lookup field in list $MSWSSRESTGenericList
    $web = $MSWSSRESTSiteCollectionObject.RootWeb
    $listName = $MSWSSRESTGenericList
    $list = $web.Lists[$listName]
    $list.Fields.AddLookup($MSWSSRESTLookupFieldName,$list.ID,$false)
		
	# Create a workflow association under specified list
	Output "Steps for manual configuration:" "Yellow"
    Output "Create a workflow association with the name of $MSWWSPWorkflowName under specified list $MSWSSRESTWorkflowTaskList ..." "Yellow"
    AddListWorkFlow $MSWSSRESTSiteCollectionObject $MSWSSRESTWorkflowTaskList $MSWSSRESTWorkflowName "Three-state" $MSWSSRESTWorkflowTaskList $MSWSSRESTWorkflowHistoryList
	
}
#-----------------------------------------------------
# Start to configure SUT for MS-VIEWSS.
#-----------------------------------------------------
Output "Start to run configurations of MS-VIEWSS." "White"
Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSVIEWSSSiteCollectionName ..." "Yellow"
$MSVIEWSSSiteCollectionObject = CreateSiteCollection $MSVIEWSSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

Output "Steps for manual configuration:" "Yellow"
Output "Create the generic list $MSVIEWSSViewListName under the root web of the site collection $MSVIEWSSSiteCollectionName ..." "Yellow"
CreateListItem $MSVIEWSSSiteCollectionObject.RootWeb $MSVIEWSSViewListName 100
$listItemNames = $MSVIEWSSListItem1,$MSVIEWSSListItem2,$MSVIEWSSListItem3,$MSVIEWSSListItem4,$MSVIEWSSListItem5,$MSVIEWSSListItem6,$MSVIEWSSListItem7,$MSVIEWSSListItem7
foreach($listItemName in $listItemNames)      
{
    Output "Steps for manual configuration:" "Yellow"
    Output "Add list item $listItemName under the list $MSVIEWSSViewListName ..." "Yellow"
	AddListItem $MSVIEWSSSiteCollectionObject.RootWeb $MSVIEWSSViewListName $listItemName
}
$MSVIEWSSSiteCollectionObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-COPYS.
#-----------------------------------------------------
Output "Start to run configurations of MS-COPYS." "White"
Output "Create two users on the domain controller:" "Yellow"
Output "Name:$MSCOPYSNoPermissionUser Password:$MSCOPYSNoPermissionUserPassword" "Yellow"
Output "Name:$MSCOPYSEditUser Password:$MSCOPYSEditUserPassword" "Yellow"
Output "Steps for manual configuration:" "Yellow"
Output "1. Open Active Directory Users and Computers..." "Yellow"
Output "2. Create two new users with the name of $MSCOPYSNoPermissionUser and $MSCOPYSEditUser with the password mentioned above..." "Yellow"
CreateUserOnDC $MSCOPYSNoPermissionUser $MSCOPYSNoPermissionUserPassword
CreateUserOnDC $MSCOPYSEditUser $MSCOPYSEditUserPassword

Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSCOPYSSiteCollectionName ..." "Yellow"
$MSCOPYSSiteCollectionObject = CreateSiteCollection $MSCOPYSSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null

Output "Steps for manual configuration:" "Yellow"
Output "Grant Edit permission level to $domain\$MSCOPYSEditUser on site $MSCOPYSSiteCollectionName..." "Yellow"
GrantUserPermission $MSCOPYSSiteCollectionObject.RootWeb "Edit" $domain.Split(".")[0] $MSCOPYSEditUser

Output "Steps for manual configuration:" "Yellow"
Output "Create a subsite named $MSCOPYSSubSite with meeting workspace template under site collection $MSCOPYSSiteCollectionName ..." "Yellow"
$MSCOPYSSiteObject = CreateWeb $MSCOPYSSiteCollectionObject $false $MSCOPYSSubSite $MSCOPYSSubSite "MPS#0" $true

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSCOPYSSubSiteDocumentLibrary under the subsite $MSCOPYSSubSite ..." "Yellow"
CreateListItem $MSCOPYSSiteObject $MSCOPYSSubSiteDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSCOPYSSourceDocumentLibrary under the site $MSCOPYSSiteCollectionName ..." "Yellow"
CreateListItem $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSSourceDocumentLibrary 101

Output "Steps for manual configuration:" "Yellow"
Output "Create the document library $MSCOPYSDestinationDocumentLibrary under the site $MSCOPYSSiteCollectionName ..." "Yellow"
CreateListItem $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSDestinationDocumentLibrary 101

#Add field in list $MSCOPYSSourceDocumentLibrary
AddFieldInList $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSSourceDocumentLibrary $MSCOPYSTextFieldName 2 $false "" $MSCOPYSSourceLibraryFieldValue "false"
AddFieldInList $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSSourceDocumentLibrary $MSCOPYSWorkFlowEventFieldName 30

#Add field in list $MSCOPYSDestinationDocumentLibrary
AddFieldInList $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSDestinationDocumentLibrary $MSCOPYSTextFieldName 2 $false "" $MSCOPYSDestinationLibraryFieldValue "true"
AddFieldInList $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSDestinationDocumentLibrary $MSCOPYSWorkFlowEventFieldName 30

Output "Steps for manual configuration:" "Yellow"
Output "Upload test data $MSCOPYSTestData to $MSCOPYSSourceDocumentLibrary under $MSSHDACCWSSiteCollectionName ..." "Yellow"
UploadFileToSharePointFolder $MSCOPYSSiteCollectionObject.RootWeb $MSCOPYSSourceDocumentLibrary $MSCOPYSTestData ".\$MSCOPYSTestData"  $True

$MSCOPYSSiteCollectionObject.Dispose()
$MSCOPYSSiteObject.Dispose()

#-----------------------------------------------------
# Start to configure SUT for MS-OFFICIALFILE.
#-----------------------------------------------------
Output "Start to run configurations of MS-OFFICIALFILE." "White"
Output "Steps for manual configuration:" "Yellow"
Output "Create a site collection named $MSOFFICIALFILESiteCollectionName ..." "Yellow"
$MSOFFICIALFILESiteCollectionObject = CreateSiteCollection $MSOFFICIALFILESiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" $null $null
$MSOFFICIALFILEWeb = $MSOFFICIALFILESiteCollectionObject.Openweb()
$MSOFFICIALFILESiteUrl = "http://$SutComputerName/sites/$MSOFFICIALFILESiteCollectionName"
$MSOFFICIALFILERoutingRepositorySiteUrl = "$MSOFFICIALFILESiteUrl/$MSOFFICIALFILERoutingRepositorySite"
    
Output "Steps for manual configuration:" "Yellow"
Output "1. Open Active Directory Users and Computers..." "Yellow"
Output "2. Create user $MSOFFICIALFILEReadUser..." "Yellow"
CreateUserOnDC $MSOFFICIALFILEReadUser $MSOFFICIALFILEReadUserPassword

if($SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
{	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create a subsite named $MSOFFICIALFILERoutingRepositorySite under site collection MSOFFICIALFILESiteCollectionName ..." "Yellow"
	$MSOFFICIALFILERoutingRepositoryWeb = CreateWeb $MSOFFICIALFILESiteCollectionObject $false $MSOFFICIALFILERoutingRepositorySite $MSOFFICIALFILERoutingRepositorySite "OFFILE#1" $true
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create a subsite named $MSOFFICIALFILENoRoutingRepositorySite under site collection MSOFFICIALFILESiteCollectionName ..." "Yellow"
	$MSOFFICIALFILENoRoutingRepositoryWeb = CreateWeb $MSOFFICIALFILESiteCollectionObject $false $MSOFFICIALFILENoRoutingRepositorySite $MSOFFICIALFILENoRoutingRepositorySite "OFFILE#1" $true
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create a subsite named $MSOFFICIALFILEEnabledParsingRepositorySite under site collection MSOFFICIALFILESiteCollectionName ..." "Yellow"
	$MSOFFICIALFILEEnabledParsingRepositoryWeb = CreateWeb $MSOFFICIALFILESiteCollectionObject $false $MSOFFICIALFILEEnabledParsingRepositorySite $MSOFFICIALFILEEnabledParsingRepositorySite "BDR#0" $true
    
	
	Output "Steps for manual configuration:" "Yellow"
    Output "Grant Read permission to user $domain\$MSOFFICIALFILEReadUser on subsite $MSOFFICIALFILERoutingRepositoryWeb..." "Yellow"
    GrantUserPermission $MSOFFICIALFILERoutingRepositoryWeb "Read" $domain.Split(".")[0] $MSOFFICIALFILEReadUser
	
	$subSite_librarys = @{$MSOFFICIALFILEDocumentRuleLocationLibrary = $MSOFFICIALFILERoutingRepositorySite;$MSOFFICIALFILENoEnforceLibrary = $MSOFFICIALFILERoutingRepositorySite;$MSOFFICIALFILEDocumentSetLocationLibrary = $MSOFFICIALFILERoutingRepositorySite;`
	$MSOFFICIALFILEDropOffLibrary = $MSOFFICIALFILENoRoutingRepositorySite}
	foreach($library in $subSite_librarys.keys)
	{  
	    $subSite =$MSOFFICIALFILESiteUrl+"/" +$subSite_librarys[$library]			
		$spSites = new-object Microsoft.SharePoint.SPSite("$subSite")
	    $rootWeb = $spSites.OpenWeb()
		Output "Steps for manual configuration:" "Yellow"
		Output "Create the document library $subSite under the subsite $subSite_librarys[$library] ..." "Yellow"
	    CreateListItem $rootWeb $library 101
	}
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create the document library $MSOFFICIALFILEDocumentRuleLocationLibrary under the subsite $MSOFFICIALFILEEnabledParsingRepositorySite ..." "Yellow"
	$parsingRepositoryRootWeb = $MSOFFICIALFILESiteCollectionObject.OpenWeb("$MSOFFICIALFILEEnabledParsingRepositorySite")
	CreateListItem $parsingRepositoryRootWeb $MSOFFICIALFILEDocumentRuleLocationLibrary 101
    
	Output "Steps for manual configuration:" "Yellow"
	Output "Enable major version on document library $MSOFFICIALFILEDocumentRuleLocationLibrary under the subsite $MSOFFICIALFILEEnabledParsingRepositorySite ..." "Yellow"
	EnableMajorVersion $MSOFFICIALFILEEnabledParsingRepositoryWeb $MSOFFICIALFILEDocumentRuleLocationLibrary	
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Active the content organizer feature on web $MSOFFICIALFILERoutingRepositorySite ..." "Yellow"
	SetWebFeature  "$MSOFFICIALFILERoutingRepositorySiteUrl" "DocumentRouting"
		
    Output "Steps for manual configuration:" "Yellow"
	Output "Enable the document parser on web $MSOFFICIALFILERoutingRepositorySite ..." "Yellow"
    $MSOFFICIALFILERoutingRepositoryWeb.ParserEnabled = $false
	$MSOFFICIALFILERoutingRepositoryWeb.Update()
		
    Output "Steps for manual configuration:" "Yellow"
	Output "Active the content organizer feature on web $MSOFFICIALFILEEnabledParsingRepositorySite ..." "Yellow"
	SetWebFeature  "$MSOFFICIALFILESiteUrl/$MSOFFICIALFILEEnabledParsingRepositorySite" "DocumentRouting"
        
    Output "Steps for manual configuration:" "Yellow"
	Output "Enable the document parser on web $MSOFFICIALFILEEnabledParsingRepositorySite ..." "Yellow"
	$MSOFFICIALFILEEnabledParsingRepositoryWeb.ParserEnabled = $true
	$MSOFFICIALFILEEnabledParsingRepositoryWeb.Update()
		
	Output "Steps for manual configuration:" "Yellow"
	Output "Deactive the content organizer feature on web $MSOFFICIALFILENoRoutingRepositorySite ..." "Yellow"
	SetWebFeature  "$MSOFFICIALFILESiteUrl/$MSOFFICIALFILENoRoutingRepositorySite" "DocumentRouting" $false
				
	Output "Steps for manual configuration:" "Yellow"
	Output "Active the feature DocumentManagement and DocumentSet at site collection $MSOFFICIALFILESiteCollectionName ..." "Yellow"
	SetWebFeature $MSOFFICIALFILESiteUrl "DocumentManagement"
	SetWebFeature $MSOFFICIALFILESiteUrl "DocumentSet"	
	
	Output "Steps for manual configuration:" "Yellow"
    Output "Grant user $userName to group Records Center Web Service Submitters ..." "Yellow"
    AddUserToSharePointGroup $MSOFFICIALFILERoutingRepositoryWeb "Records Center Web Service Submitters for $MSOFFICIALFILERoutingRepositorySite" "$domain" "$userName"
	AddUserToSharePointGroup $MSOFFICIALFILENoRoutingRepositoryWeb "Records Center Web Service Submitters for $MSOFFICIALFILENoRoutingRepositorySite" "$domain" "$userName"
	AddUserToSharePointGroup $MSOFFICIALFILEEnabledParsingRepositoryWeb "Records Center Web Service Submitters for $MSOFFICIALFILEEnabledParsingRepositorySite" "$domain" "$userName"
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create Organizer Rules for $MSOFFICIALFILEDocumentRuleLocationLibrary under $MSOFFICIALFILERoutingRepositorySite ..."
	CreateContentOrganizerRules "$MSOFFICIALFILESiteUrl/$MSOFFICIALFILERoutingRepositorySite" "$MSOFFICIALFILEDocumentRuleLocationLibrary" "DocumentRule" "Document" "Name,IsNotEmpty,$null","Title,IsNotEmpty,$null"
	
	Output "Steps for manual configuration:" "Yellow"
	Output "Create Organizer Rules for $MSOFFICIALFILEDocumentRuleLocationLibrary under $MSOFFICIALFILEEnabledParsingRepositorySite ..."
	CreateContentOrganizerRules "$MSOFFICIALFILESiteUrl/$MSOFFICIALFILEEnabledParsingRepositorySite" "$MSOFFICIALFILEDocumentRuleLocationLibrary" "DocumentRule" "Document" "Name,IsNotEmpty,$null","Title,IsNotEmpty,$null"

	Output "Steps for manual configuration:" "Yellow"
	Output "Add Content Types which is Document Set from Exiting Site Content Types on $MSOFFICIALFILEDocumentLibrary1..." "Yellow"
	AddContentTypeToList $MSOFFICIALFILERoutingRepositorySiteUrl $MSOFFICIALFILEDocumentSetLocationLibrary "Document Set"

	Output "Steps for manual configuration:" "Yellow"
	Output "Create a document set $MSOFFICIALFILEDocumentSetName on $MSOFFICIALFILERoutingRepositorySiteUrl..." "Yellow"
	CreateDocumentSet $MSOFFICIALFILERoutingRepositorySiteUrl $MSOFFICIALFILEDocumentSetLocationLibrary $MSOFFICIALFILEDocumentSetName   
    
	Output "Steps for manual configuration:" "Yellow"
	Output "Create a hold $MSOFFICIALFILEHolds on $MSOFFICIALFILERoutingRepositorySiteUrl..." "Yellow"
	AddHolds $MSOFFICIALFILERoutingRepositorySiteUrl $MSOFFICIALFILEHolds
	
	$MSOFFICIALFILESiteCollectionObject.Dispose()
	$MSOFFICIALFILERoutingRepositoryWeb.Dispose()
	$MSOFFICIALFILEEnabledParsingRepositoryWeb.Dispose()
	$MSOFFICIALFILENoRoutingRepositoryWeb.Dispose()
}

if($SharePointVersion -eq $SharePointServer2007[0])
{
    Output "Steps for manual configuration:" "Yellow"
	Output "Create a subsite named $MSOFFICIALFILERoutingRepositorySite under site collection $MSOFFICIALFILESiteCollectionName ..." "Yellow"
    CreateRecordCenterOn2007 $MSOFFICIALFILESiteUrl $MSOFFICIALFILERoutingRepositorySite	
	$MSOFFICIALFILERoutingRepositoryWeb = $MSOFFICIALFILESiteCollectionObject.OpenWeb($MSOFFICIALFILERoutingRepositorySite)

	Output "Steps for manual configuration:" "Yellow"
    Output "Grant user $userName to group Records Center Web Service Submitters ..." "Yellow"
    AddUserToSharePointGroup $MSOFFICIALFILERoutingRepositoryWeb "Records Center Web Service Submitters for $MSOFFICIALFILERoutingRepositorySite" "$domain" "$userName"

	Output "Steps for manual configuration:" "Yellow"
	Output "Create a hold $MSOFFICIALFILEHolds on $MSOFFICIALFILERoutingRepositorySiteUrl..." "Yellow"
	AddHolds $MSOFFICIALFILERoutingRepositorySiteUrl $MSOFFICIALFILEHolds	
	
	Output "Steps for manual configuration:" "Yellow"
    Output "Grant Read permission to user $domain\$MSOFFICIALFILEReadUser on site $MSOFFICIALFILEWeb..." "Yellow"
    GrantUserPermission $MSOFFICIALFILEWeb "Read" $Domain.Split(".")[0] $MSOFFICIALFILEReadUser

	Output "Steps for manual configuration:" "Yellow"
	Output "Create the document library $MSOFFICIALFILEDocumentRuleLocationLibrary and $MSOFFICIALFILEDropOffLibrary under the subsite $MSOFFICIALFILERoutingRepositorySite ..." "Yellow"
	CreateListItem $MSOFFICIALFILERoutingRepositoryWeb $MSOFFICIALFILEDocumentRuleLocationLibrary 101
    CreateListItem $MSOFFICIALFILERoutingRepositoryWeb $MSOFFICIALFILEDropOffLibrary 101
    
	Output "Steps for manual configuration:" "Yellow"
	Output "New item to record routing for $MSOFFICIALFILEDocumentRuleLocationLibrary under $MSOFFICIALFILERoutingRepositorySite ..." "Yellow"	
	NewItemToRecordRouting $MSOFFICIALFILERoutingRepositoryWeb "DocumentRule" "$MSOFFICIALFILEDocumentRuleLocationLibrary" "Document" $false
	
	Output "Steps for manual configuration:" "Yellow"
	Output "New item to record routing for $MSOFFICIALFILEDropOffLibrary under $MSOFFICIALFILERoutingRepositorySite ..." "Yellow"
	NewItemToRecordRouting $MSOFFICIALFILERoutingRepositoryWeb "DefaultRule" "$MSOFFICIALFILEDropOffLibrary" "Picture"
    
	$MSOFFICIALFILERoutingRepositoryWeb.Dispose()
}

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
Output "The server configuration script was executed successfully" "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0