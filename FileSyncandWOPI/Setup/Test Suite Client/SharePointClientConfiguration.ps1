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
#----------------------------------------------------------------------------
# Default Values of Configuration. 
#----------------------------------------------------------------------------
$environmentResourceFile                     = "$commonScriptDirectory\SharePointTestSuite.config"

$User1                                       = ReadConfigFileNode "$environmentResourceFile" "User1"
$User1Password                               = ReadConfigFileNode "$environmentResourceFile" "User1Password"                 
$User2                                       = ReadConfigFileNode "$environmentResourceFile" "User2"
$User2Password                               = ReadConfigFileNode "$environmentResourceFile" "User2Password"
$User3                                       = ReadConfigFileNode "$environmentResourceFile" "User3"
$User3Password                               = ReadConfigFileNode "$environmentResourceFile" "User3Password"
$ReadOnlyUser                                = ReadConfigFileNode "$environmentResourceFile" "ReadOnlyUser"
$ReadOnlyUserPassword                        = ReadConfigFileNode "$environmentResourceFile" "ReadOnlyUserPassword"
$NoUseRemoteUser                             = ReadConfigFileNode "$environmentResourceFile" "NoUseRemoteUser"
$NoUseRemoteUserUserPs                       = ReadConfigFileNode "$environmentResourceFile" "NoUseRemoteUserUserPs"
$FileSyncWOPIUser                            = ReadConfigFileNode "$environmentResourceFile" "FileSyncWOPIUser"
$FileSyncWOPIUserPassword                    = ReadConfigFileNode "$environmentResourceFile" "FileSyncWOPIUserPassword"
$FileSyncWOPIBigTestData1                    = ReadConfigFileNode "$environmentResourceFile" "FileSyncWOPIBigTestData1"

$MSFSSHTTPFSSHTTPBSiteCollectionName         = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBSiteCollectionName"
$MSFSSHTTPFSSHTTPBDocumentLibrary            = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBDocumentLibrary"
$MSFSSHTTPFSSHTTPBZipTestData2               = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBZipTestData2"
$MSFSSHTTPFSSHTTPBTestData3                  = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBTestData3"
$MSFSSHTTPFSSHTTPBTestData4                  = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBTestData4"
$MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel = ReadConfigFileNode "$environmentResourceFile" "MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel"

$MSWOPISiteCollectionName                    = ReadConfigFileNode "$environmentResourceFile" "MSWOPISiteCollectionName"
$MSWOPISharedDocumentLibrary                 = ReadConfigFileNode "$environmentResourceFile" "MSWOPISharedDocumentLibrary"
$MSWOPISharedZipTestData2                    = ReadConfigFileNode "$environmentResourceFile" "MSWOPISharedZipTestData2"
$MSWOPISharedTestData3                       = ReadConfigFileNode "$environmentResourceFile" "MSWOPISharedTestData3"
$MSWOPISharedTestData4                       = ReadConfigFileNode "$environmentResourceFile" "MSWOPISharedTestData4"
$MSWOPIDocumentLibrary                       = ReadConfigFileNode "$environmentResourceFile" "MSWOPIDocumentLibrary"
$MSWOPITestFolder                            = ReadConfigFileNode "$environmentResourceFile" "MSWOPITestFolder"
$MSWOPITestData1                             = ReadConfigFileNode "$environmentResourceFile" "MSWOPITestData1"
$MSWOPITestData2                             = ReadConfigFileNode "$environmentResourceFile" "MSWOPITestData2"
$MSWOPITargetAppWithNotGroupAndWindows       = ReadConfigFileNode "$environmentResourceFile" "MSWOPITargetAppWithNotGroupAndWindows"
$MSWOPITargetAppWithGroupAndNoWindows        = ReadConfigFileNode "$environmentResourceFile" "MSWOPITargetAppWithGroupAndNoWindows"
$MSWOPIUserCredentialItem                    = ReadConfigFileNode "$environmentResourceFile" "MSWOPIUserCredentialItem"
$MSWOPIPasswordCredentialItem                = ReadConfigFileNode "$environmentResourceFile" "MSWOPIPasswordCredentialItem"
$MSWOPIFolderCreatedByUser1                  = ReadConfigFileNode "$environmentResourceFile" "MSWOPIFolderCreatedByUser1"
$MSWOPINoUseRemotePermissionLevel            = ReadConfigFileNode "$environmentResourceFile" "MSWOPINoUseRemotePermissionLevel"

$MSONESTORESiteCollectionName                = ReadConfigFileNode "$environmentResourceFile" "MSONESTORESiteCollectionName"
$MSONESTORELibraryName                       = ReadConfigFileNode "$environmentResourceFile" "MSONESTORELibraryName"
$MSONESTOREOneFileWithFileData               = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileWithFileData"
$MSONESTOREOneFileWithoutFileData            = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileWithoutFileData"
$MSONESTOREOneFileEncryption                 = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileEncryption"
$MSONESTOREOneWithInvalid                    = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneWithInvalid"
$MSONESTOREOneWithLarge                      = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneWithLarge"
$MSONESTOREOnetocFileLocal                   = ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOnetocFileLocal"
$MSONESTORENoSectionFile                     = ReadConfigFileNode "$environmentResourceFile" "MSONESTORENoSectionFile"

#-----------------------------------------------------
# Paths for all PTF configuration files.
#-----------------------------------------------------
$CommonDeploymentFile = resolve-path "..\..\Source\Common\FssWopiCommonConfiguration.deployment.ptfconfig"
$MSFSSHTTPFSSHTTPBDeploymentFile = resolve-path "..\..\Source\MS-FSSHTTP-FSSHTTPB\TestSuite\MS-FSSHTTP-FSSHTTPB_TestSuite.deployment.ptfconfig"
$MSWOPIDeploymentFile = resolve-path "..\..\Source\MS-WOPI\TestSuite\MS-WOPI_TestSuite.deployment.ptfconfig"
$MSONESTOREDeploymentFile = resolve-path "..\..\Source\MS-ONESTORE\TestSuite\MS-ONESTORE_TestSuite.deployment.ptfconfig"

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
Output "The SUT configuration must be configured before running the client setup script." "Cyan"
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
    Output "1: CONTINUE (Some test cases may fail if the recommended application(s) are not installed)." "Cyan"
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
# Add a firewall rule to allow Test suite client receive the WOPI discovery request from the WOPI server.
#-----------------------------------------------------
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ENV:COMPUTERNAME)
$profileNames = "DomainProfile","StandardProfile","PublicProfile"	
foreach($profileName in $profileNames)
{
    $firewallStatus = $reg.OpenSubKey("System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\$profileName").GetValue("EnableFirewall")
	$totalStatus = $firewallStatus -bor $totalStatus
}	
if($totalStatus)
{
	Output "Add a firewall rule to allow the test suite client to receive the WOPI Discovery request from the WOPI server." "White"
	$firewallPolicy = New-Object -ComObject hnetcfg.fwpolicy2 
	$newRule = New-Object -ComObject HNetCfg.FWRule
	$newRule.Name = "Receive the WOPI discovery request from the WOPI server " #Rule display name
	$newRule.Protocol = 6 #TCP
	$newRule.LocalPorts = 80
	$newRule.Enabled = $true
	$isRuleAdded = $false
	foreach($rule in $firewallPolicy.Rules)
	{
	    if(($rule.Name -eq $newRule.Name) -and ($rule.Protocol -eq $newRule.Protocol) -and ($rule.LocalPorts -eq $newRule.LocalPorts) -and ($rule.Enabled -eq $newRule.Enabled))
	    {
	        Output "Firewall rule is already added." "Yellow"
	        $isRuleAdded = $true
	        break
	    }
	}
	if(!$isRuleAdded)
	{
	    $firewallPolicy.Rules.Add($newRule)
		Output "Add the port number 80 to the firewall rule successfully. " "Green"
	}
}

#-----------------------------------------------------
# Configuration for common ptfconfig file.
#-----------------------------------------------------
Output "Enter the computer name of the SUT:" "Cyan"
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported." "Cyan"
$sutComputerName = ReadComputerName $false "sutComputerName"
Output "The computer name of SUT that you entered: $sutComputerName" "White"

Output "Check the status of the Windows Remote Management (WinRM) service to make sure that the service is running." "White"
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
	    if ( $error[0].Exception -match "Microsoft.PowerShell.Commands.ServiceCommandException")
		{
		    Output "Failed to start the service $service. Start it manually and then press any key to continue ..." "Red"
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

Output "Enter the user who will call protocol methods in the test suite and remotely configure the SUT if the SUT control adapter is set to Powershell mode:" "Cyan"
Output "The user should be able to create users in Active Directory directory service, be a part of the local admin group on the server, and also be the SUT administrator." "Cyan"
if (($Env:USERDNSDOMAIN -ne $null) -and ($Env:USERDNSDOMAIN -ine $ENV:COMPUTERNAME) -and ($ENV:USERNAME -ne $null))
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

if (!$useCurrentUser)
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
Output "The password you entered: $password" "White"

$endPointComputerName   = $ENV:ComputerName
Output "The computer name of the test suite client: $endPointComputerName" "White"

Output "Try to get the SharePoint version on the selected server ..." "White"
$sutVersionInfo = GetSharePointServerVersion $sutComputerName ($dnsDomain.split(".")[0]+ "\" + $userName) $password

if($sutVersionInfo -ne $null -and $sutVersionInfo -ne "" -and $sutVersionInfo -ne "Unknown Version")
{
    $sutVersion = $sutVersionInfo[0]
    if($sutVersion -eq $script:WindowsSharePointServices3OnSUT[0] -or $sutVersion -eq $script:SharePointServer2007OnSUT[0])
    {
        Write-Warning "Could not find the supported version of SharePoint server on the server! Install one of the recommended versions ($($script:SharePointFoundation2010OnSUT[1]) $($script:SharePointFoundation2010OnSUT[2]), $($script:SharePointServer2010OnSUT[1]) $($script:SharePointServer2010OnSUT[2]), $($script:SharePointFoundation2013OnSUT[1]) $($script:SharePointFoundation2013OnSUT[2]), $($script:SharePointServer2013OnSUT[1])$($script:SharePointServer2013OnSUT[2]), $($script:SharePointServer2016OnSUT[1])) first and run the SharePointClientConfiguration.ps1 again.`r`n"
        Stop-Transcript        
        exit 2
    }
    else
    {
        Output ("The SharePoint version installed on the server is " + $sutVersionInfo[1] +" " + $sutVersionInfo[2]+ ".") "Green"
    }
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

#The URL of the first file.
$firstFileUrl = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/[MSFSSHTTPFSSHTTPBLibraryName]/$MSFSSHTTPFSSHTTPBTestData3"

#The URL of the file which is larger than 1MB.
$bigFile_Url = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/[MSFSSHTTPFSSHTTPBLibraryName]/$FileSyncWOPIBigTestData1"

#The URL of the file should be zip file format.
$zipFile_Url = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/[MSFSSHTTPFSSHTTPBLibraryName]/$MSFSSHTTPFSSHTTPBZipTestData2"

#The URL of the file should be OneNote file format.
$OneNoteFile_Url = "[TransportType]://[SUTComputerName]/sites/[SiteCollectionName]/[MSFSSHTTPFSSHTTPBLibraryName]/$MSFSSHTTPFSSHTTPBTestData4"

Output "Configure the SharePointCommonConfiguration.deployment.ptfconfig file ..." "White"
Output "Modify the properties as necessary in the the SharePointCommonConfiguration.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $CommonDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SutComputerName`", and set the value to $sutComputerName." "Yellow"
$step++
Output "$step.Find the property `"Domain`", and set the value to $dnsDomain." "Yellow"
$step++
Output "$step.Find the property `"SutVersion`", and set the value to $sutVersion." "Yellow"
$step++
Output "$step.Find the property `"TransportType`", and set the value to $transportType." "Yellow"
$step++
Output "$step.Find the property `"NormalFile`", and set the value to $firstFileUrl" "Yellow"
$step++
Output "$step.Find the property `"BigFile`", and set the value to $bigFile_Url" "Yellow"
$step++
Output "$step.Find the property `"ZipFile`", and set the value to $zipFile_Url" "Yellow"
$step++
Output "$step.Find the property `"OneNoteFile`", and set the value to $OneNoteFile_Url" "Yellow"
$step++
Output "$step.Find the property `"UserName1`", and set the value to $User1" "Yellow"
$step++
Output "$step.Find the property `"Password1`", and set the value to $User1Password" "Yellow"
$step++
Output "$step.Find the property `"UserName2`", and set the value to $User2" "Yellow"
$step++
Output "$step.Find the property `"Password2`", and set the value to $User2Password" "Yellow"
$step++
Output "$step.Find the property `"UserName3`", and set the value to $User3" "Yellow"
$step++
Output "$step.Find the property `"Password3`", and set the value to $User3Password" "Yellow"
$step++
Output "$step.Find the property `"ReadOnlyUser`", and set the value to $ReadOnlyUser" "Yellow"
$step++
Output "$step.Find the property `"ReadOnlyUserPwd`", and set the value to $ReadOnlyUserPassword" "Yellow"
$step++
Output "$step.Find the property `"NoPermisionToUseRemoteInterfaceUser`", and set the value to $NoUseRemoteUser" "Yellow"
$step++
Output "$step.Find the property `"NoPermisionToUseRemoteInterfaceUserPwd`", and set the value to $NoUseRemoteUserUserPs" "Yellow"

ModifyConfigFileNode $CommonDeploymentFile "SutComputerName"                            $sutComputerName
ModifyConfigFileNode $CommonDeploymentFile "Domain"                                     $dnsDomain
ModifyConfigFileNode $CommonDeploymentFile "SutVersion"                                 $sutVersion
ModifyConfigFileNode $CommonDeploymentFile "TransportType"                              $transportType
ModifyConfigFileNode $CommonDeploymentFile "NormalFile"                                 $firstFileUrl
ModifyConfigFileNode $CommonDeploymentFile "BigFile"                                    $bigFile_Url
ModifyConfigFileNode $CommonDeploymentFile "ZipFile"                                    $zipFile_Url
ModifyConfigFileNode $CommonDeploymentFile "OneNoteFile"                                $OneNoteFile_Url
ModifyConfigFileNode $CommonDeploymentFile "UserName1"                                  $User1
ModifyConfigFileNode $CommonDeploymentFile "Password1"                                  $User1Password
ModifyConfigFileNode $CommonDeploymentFile "UserName2"                                  $User2
ModifyConfigFileNode $CommonDeploymentFile "Password2"                                  $User2Password
ModifyConfigFileNode $CommonDeploymentFile "UserName3"                                  $User3
ModifyConfigFileNode $CommonDeploymentFile "Password3"                                  $User3Password
ModifyConfigFileNode $CommonDeploymentFile "ReadOnlyUser"                               $ReadOnlyUser
ModifyConfigFileNode $CommonDeploymentFile "ReadOnlyUserPwd"                            $ReadOnlyUserPassword
ModifyConfigFileNode $CommonDeploymentFile "NoPermisionToUseRemoteInterfaceUser"        $NoUseRemoteUser
ModifyConfigFileNode $CommonDeploymentFile "NoPermisionToUseRemoteInterfaceUserPwd"     $NoUseRemoteUserUserPs

Output "Configuration for the SharePointCommonConfiguration.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-FSSHTTP-FSSHTTPB ptfconfig file.
#-----------------------------------------------------
Output "Configure the MS-FSSHTTP-FSSHTTPB_TestSuite.deployment.ptfconfig file ..." "White"

Output "Modify the properties as necessary in the MS-FSSHTTP-FSSHTTPB_TestSuite.deployment.ptfconfig file..." "White"
$step = 1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSFSSHTTPFSSHTTPBDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSFSSHTTPFSSHTTPBSiteCollectionName" "Yellow"
$step++
Output "$step.Find the property `"MSFSSHTTPFSSHTTPBLibraryName`", and set the value to $MSFSSHTTPFSSHTTPBDocumentLibrary" "Yellow"

ModifyConfigFileNode $MSFSSHTTPFSSHTTPBDeploymentFile "SiteCollectionName"                    $MSFSSHTTPFSSHTTPBSiteCollectionName
ModifyConfigFileNode $MSFSSHTTPFSSHTTPBDeploymentFile "MSFSSHTTPFSSHTTPBLibraryName"          $MSFSSHTTPFSSHTTPBDocumentLibrary

Output "Configuration for the MS-FSSHTTP-FSSHTTPB_TestSuite.deployment.ptfconfig file is complete." "Green"

#-----------------------------------------------------
# Configuration for MS-WOPI ptfconfig file.
#-----------------------------------------------------
if($sutVersion -ge $script:SharePointFoundation2013OnSUT[0] -or $sutVersion -ge $script:SharePointServer2013OnSUT[0])
{
    Output "Configure the MS-WOPI_TestSuite.deployment.ptfconfig file ..." "White"

    Output "Get the value of the DisplayName attribute of $FileSyncWOPIUser..." "White"
    $fileSyncWOPIUserDisplayName = GetUserDisplayName $sutComputerName $MSWOPISiteCollectionName $FileSyncWOPIUser "$dnsDomain\$userName" $password

    Output "Modify the properties as necessary in the MS-WOPI_TestSuite.deployment.ptfconfig file..." "White"
    $step = 1
    Output "Steps for manual configuration:" "Yellow"
    Output "$step.Open $MSWOPIDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSWOPISiteCollectionName" "Yellow"
    $step++
    Output "$step.Find the property `"MSFSSHTTPFSSHTTPBLibraryName`", and set the value to $MSWOPISharedDocumentLibrary" "Yellow"
    $step++
    Output "$step.Find the property `"MSWOPIDocLibraryName`", and set the value to $MSWOPIDocumentLibrary" "Yellow"
    $step++
    Output "$step.Find the property `"TestClientName`", and set the value to $endPointComputerName" "Yellow"
    $step++
    Output "$step.Find the property `"IdOfAppWithIndividualAndWindows`", and set the value to $MSWOPITargetAppWithNotGroupAndWindows" "Yellow"
    $step++
    Output "$step.Find the property `"IdOfAppWithGroupAndNotWindows`", and set the value to $MSWOPITargetAppWithGroupAndNoWindows" "Yellow"
    $step++
    Output "$step.Find the property `"ValueOfUserCredentialItem`", and set the value to $MSWOPIUserCredentialItem" "Yellow"
    $step++
    Output "$step.Find the property `"ValueOfPasswordCredentialItem`", and set the value to $MSWOPIPasswordCredentialItem" "Yellow"
    $step++
    Output "$step.Find the property `"UserName`", and set the value to $FileSyncWOPIUser." "Yellow"
    $step++
    Output "$step.Find the property `"Password`", and set the value to $FileSyncWOPIUserPassword." "Yellow"
    $step++
    Output "$step.Find the property `"UserFriendlyName`", and set the value to $fileSyncWOPIUserDisplayName." "Yellow"

    ModifyConfigFileNode $MSWOPIDeploymentFile "SiteCollectionName"                    $MSWOPISiteCollectionName
    ModifyConfigFileNode $MSWOPIDeploymentFile "MSFSSHTTPFSSHTTPBLibraryName"          $MSWOPISharedDocumentLibrary 
    ModifyConfigFileNode $MSWOPIDeploymentFile "MSWOPIDocLibraryName"                  $MSWOPIDocumentLibrary 
    ModifyConfigFileNode $MSWOPIDeploymentFile "TestClientName"                        $endPointComputerName
    ModifyConfigFileNode $MSWOPIDeploymentFile "IdOfAppWithIndividualAndWindows"       $MSWOPITargetAppWithNotGroupAndWindows
    ModifyConfigFileNode $MSWOPIDeploymentFile "IdOfAppWithGroupAndNotWindows"         $MSWOPITargetAppWithGroupAndNoWindows
    ModifyConfigFileNode $MSWOPIDeploymentFile "ValueOfUserCredentialItem"             $MSWOPIUserCredentialItem
    ModifyConfigFileNode $MSWOPIDeploymentFile "ValueOfPasswordCredentialItem"         $MSWOPIPasswordCredentialItem
    ModifyConfigFileNode $MSWOPIDeploymentFile "UserName"                              $FileSyncWOPIUser
    ModifyConfigFileNode $MSWOPIDeploymentFile "Password"                              $FileSyncWOPIUserPassword
    ModifyConfigFileNode $MSWOPIDeploymentFile "UserFriendlyName"                      $fileSyncWOPIUserDisplayName

    Output "Configuration for the MS-WOPI_TestSuite.deployment.ptfconfig file is complete." "Green"
}

#-----------------------------------------------------
# Configuration for MS-ONESTORE ptfconfig file.
#-----------------------------------------------------
if($sutVersion -ge $script:SharePointFoundation2010OnSUT[0] -or $sutVersion -ge $script:SharePointServer2010OnSUT[0])
{
    Output "Modify the properties as necessary in the MS-ONESTORE_TestSuite.deployment.ptfconfig..." "White"
    $step = 1
    Output "Steps for manual configuration:" "Yellow"
    Output "$step.Open $MSONESTOREDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"SiteCollectionName`", and set the value to $MSONESTORESiteCollectionName" "Yellow"
    $step++
    Output "$step.Find the property `"MSONESTORELibraryName`", and set the value to $MSONESTORELibraryName" "Yellow"
    $step++
    Output "$step.Find the property `"OneFileWithFileData`", and set the value to $MSONESTOREOneFileWithFileData" "Yellow"
    $step++
    Output "$step.Find the property `"OneFileWithoutFileData`", and set the value to $MSONESTOREOneFileWithoutFileData" "Yellow"
    $step++
    Output "$step.Find the property `"OneFileEncryption`", and set the value to $MSONESTOREOneFileEncryption" "Yellow"
    $step++
    Output "$step.Find the property `"OneWithInvalid`", and set the value to $MSONESTOREOneWithInvalid" "Yellow"
    $step++
    Output "$step.Find the property `"OneWithLarge`", and set the value to $MSONESTOREOneWithLarge" "Yellow"
    $step++
    Output "$step.Find the property `"OnetocFileLocal`", and set the value to $MSONESTOREOnetocFileLocal" "Yellow"
    $step++
    Output "$step.Find the property `"NoSectionFile`", and set the value to $MSONESTORENoSectionFile." "Yellow"

    ModifyConfigFileNode $MSONESTOREDeploymentFile "SiteCollectionName"                    $MSONESTORESiteCollectionName
    ModifyConfigFileNode $MSONESTOREDeploymentFile "MSONESTORELibraryName"                 $MSONESTORELibraryName 
    ModifyConfigFileNode $MSONESTOREDeploymentFile "OneFileWithFileData"                   $MSONESTOREOneFileWithFileData 
    ModifyConfigFileNode $MSONESTOREDeploymentFile "OneFileWithoutFileData"                $MSONESTOREOneFileWithoutFileData
    ModifyConfigFileNode $MSONESTOREDeploymentFile "OneFileEncryption"                     $MSONESTOREOneFileEncryption
   #ModifyConfigFileNode $MSONESTOREDeploymentFile "OneWithInvalid"                        $MSONESTOREOneWithInvalid
    ModifyConfigFileNode $MSONESTOREDeploymentFile "OneWithLarge"                          $MSONESTOREOneWithLarge
    ModifyConfigFileNode $MSONESTOREDeploymentFile "OnetocFileLocal"                       $MSONESTOREOnetocFileLocal
    ModifyConfigFileNode $MSONESTOREDeploymentFile "NoSectionFile"                         $MSONESTORENoSectionFile


    Output "Configuration for the MS-ONESTORE_TestSuite.deployment.ptfconfig file is complete." "Green"
}

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
Output "[SharePointClientConfiguration.ps1] has run successfully." "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0