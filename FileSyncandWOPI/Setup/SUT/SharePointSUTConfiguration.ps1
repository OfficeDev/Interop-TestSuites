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

$MSONESTORESiteCollectionName                =ReadConfigFileNode "$environmentResourceFile" "MSONESTORESiteCollectionName"
$MSONESTORELibraryName                       =ReadConfigFileNode "$environmentResourceFile" "MSONESTORELibraryName"
$MSONESTOREOneFileWithFileData               =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileWithFileData"
$MSONESTOREOneFileWithoutFileData            =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileWithoutFileData"
$MSONESTOREOneFileEncryption                 =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneFileEncryption"
$MSONESTOREOneWithInvalid                    =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneWithInvalid"
$MSONESTOREOneWithLarge                      =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOneWithLarge"
$MSONESTOREOnetocFileLocal                   =ReadConfigFileNode "$environmentResourceFile" "MSONESTOREOnetocFileLocal"
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
            Output "The SUT setup script will run in unattended mode with the information provided by the SUT configuration XML `"$unattendedXmlName`"." "White"
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
# Start automatic services required by test case
#-----------------------------------------------------
StartService "SP*" "Auto" "*SharePoint*"
StartService "MSSQL*" "Auto" "*SQL Server*"

#-----------------------------------------------------
# Start to configure server.
#-----------------------------------------------------
Output "Start to configure the server ..." "White"

Output "Steps for manual configuration:" "Yellow" 
Output "Enable PowerShell remoting." "Yellow"
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
    Output "The maximum amount of memory allocated per shell for remote shell management is increased from $originalMaxMemory MB to $actualMaxMemory MB." "White"
}
else
{
    Output "The maximum amount of memory allocated per shell for remote shell management is $originalMaxMemory MB." "White"
}

#-----------------------------------------------------
# Get SharePoint server basic information.
#-----------------------------------------------------
Output "SharePoint server basic information:" "White"

$domain            = $Env:USERDNSDOMAIN
Output "Domain name: $domain" "White"
$sutComputerName   = $ENV:ComputerName
Output "SharePoint server name: $sutComputerName" "White"
$userName          = $ENV:UserName
Output "Current user logon name: $userName " "White"

Output "Try to get the SharePoint server version ..." "White"
$SharePointVersionInfo = GetSharePointVersion
$SharePointVersion = $SharePointVersionInfo[0]
if($SharePointVersion -eq $WindowsSharePointServices3[0] -or $SharePointVersion -eq $SharePointServer2007[0] -or $SharePointVersion -eq "Unknown Version")
{
    Write-Warning "Could not find the supported version of SharePoint server on the system! Install one of the recommended versions ($($script:SharePointFoundation2010[0]) $($script:SharePointFoundation2010[2]), $($script:SharePointServer2010[0]) $($script:SharePointServer2010[2]), $($script:SharePointFoundation2013[0]) $($script:SharePointFoundation2013[2]), $($script:SharePointServer2013[0]) $($script:SharePointServer2013[2]), $($script:SharePointServer2016[0])) and run the SUT configuration script again.`r`n"
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

$SharePointShellSnapIn = Get-PSSnapin | Where-Object -FilterScript {$_.Name -eq "Microsoft.SharePoint.PowerShell"}
if($SharePointShellSnapIn -eq $null)
{
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

#----------------------------------------------------------------------------
# Start to configure SharePoint SUT to support HTTPS transport.
#----------------------------------------------------------------------------
Output "Configure the HTTPS service in the SUT." "White"
Output "Steps for manual configuration:" "Yellow"
Output "1. Configure the SUT to support HTTPS." "Yellow"
Output "2. Set an alternate access mapping for HTTPS." "Yellow"
$webAppName = GetWebAPPName
AddHTTPSBinding "$sutComputerName" $SharePointVersion $webAppName

#----------------------------------------------------------------------------
# Change the authentication mode to claim based.
#----------------------------------------------------------------------------
$webApplicationUrl = "http://$sutComputerName"
if($SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -ge $SharePointServer2013[0] )
{
    ChangeAuthenticationModeToClaimBased $webApplicationUrl
}

#----------------------------------------------------------------------------
# Add the user policy to enable the client side PowerShell scripts 
# to manage the SharePoint remotely.
#----------------------------------------------------------------------------
$uri = new-object System.Uri($webApplicationUrl)
$webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($uri)
$useClaims = $webApp.UseClaimsAuthentication
if($useClaims)
{
    Output "Steps for manual configuration:" "Yellow"
    Output "Add a user policy for $userName with ""Full Control"" permissions without prefixed names (such as i:0#.w)." "Yellow"
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $userName)
}

#----------------------------------------------------------------------------
# Create and configure users on the domain controller
#----------------------------------------------------------------------------
Output "Steps for manual configuration:" "Yellow"
Output "Create six users on the domain controller:" "Yellow"
Output "Name:$User1 Password:$User1Password" "Yellow"
Output "Name:$User2 Password:$User2Password" "Yellow"
Output "Name:$User3 Password:$User3Password" "Yellow"
Output "Name:$NoUseRemoteUser Password:$NoUseRemoteUserUserPs" "Yellow"
Output "Name:$ReadOnlyUser Password:$ReadOnlyUserPassword" "Yellow"
Output "Name:$FileSyncWOPIUser Password:$FileSyncWOPIUserPassword" "Yellow"

Output "Steps for manual configuration:" "Yellow"
Output "1. Open Active Directory Users and Computers..." "Yellow"
Output "2. Create six new users with the names $User1, $User2, $User3, $NoUseRemoteUser, $ReadOnlyUser, and$FileSyncWOPIUser; and the passwords $User1Password, $User2Password, $User3Password, $NoUseRemoteUserUserPs, $ReadOnlyUserPassword, and $FileSyncWOPIUserPassword." "Yellow"
CreateUserOnDC $User1 $User1Password
CreateUserOnDC $User2 $User2Password
CreateUserOnDC $User3 $User3Password
CreateUserOnDC $NoUseRemoteUser  $NoUseRemoteUserUserPs
CreateUserOnDC $ReadOnlyUser $ReadOnlyUserPassword
CreateUserOnDC $FileSyncWOPIUser $FileSyncWOPIUserPassword

Output "Steps for manual configuration:" "Yellow"
Output "Add three domain users as administrators." "Yellow"
Output "1. Open Manage User Accounts." "Yellow"
Output "2. Add users $User1, $User2 and $FileSyncWOPIUser as administrators ..." "Yellow"
AddUserToGroup $domain.Split(".")[0] $User1 "Administrators"
AddUserToGroup $domain.Split(".")[0] $User2 "Administrators"
AddUserToGroup $domain.Split(".")[0] $FileSyncWOPIUser "Administrators"

if($useClaims)
{
    Output "Steps for manual configuration:" "Yellow"
    Output "Add user policies for $User1, $User2, $User3, and $FileSyncWOPIUser with ""Full Control"" permissions without prefixed names (such as i:0#.w)." "Yellow"
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $User1)
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $User2)
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $User3)
    AddUserPolicyWithoutNamePrefix $webApplicationUrl ($domain.split(".")[0] + "\" + $FileSyncWOPIUser)
}

$isFarmMode = CheckServerInstallationMode $SharePointVersion
if($isFarmMode)
{   
    #Create login user in SQL Server.
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO")
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
    
    CreateLoginUserOnSQL ($domain.split(".")[0] + "\" + $User1)
    CreateLoginUserOnSQL ($domain.split(".")[0] + "\" + $User2)
    CreateLoginUserOnSQL ($domain.split(".")[0] + "\" + $User3)
    CreateLoginUserOnSQL ($domain.split(".")[0] + "\" + $ReadOnlyUser)
    CreateLoginUserOnSQL ($domain.split(".")[0] + "\" + $FileSyncWOPIUser)
}

#----------------------------------------------------------------------------
# Create a test data
#----------------------------------------------------------------------------
Output "Steps for manual configuration:" "Yellow"
Output "Create a 1.1MB-sized text file named $FileSyncWOPIBigTestData1." "Yellow"
CreateFile $FileSyncWOPIBigTestData1 1.1mb $containerPath

#-----------------------------------------------------
# Start to configure SUT for MS-FSSHTTP-FSSHTTPB
#-----------------------------------------------------
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -ge $SharePointServer2013[0] )
{
    Output "Start to run configuration for MS-FSSHTTP-FSSHTTPB ..." "White"
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a site collection with the name $MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    $MSFSSHTTPFSSHTTPBSiteCollectionObject = CreateSiteCollection $MSFSSHTTPFSSHTTPBSiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" "STS#0" 1033
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a permission with the name $MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel with permissions: ViewListItems, EditListItems, DeleteListItems, OpenItems, ViewVersions, Open, and ViewPages." "Yellow"
    CreatePermissionLevel $MSFSSHTTPFSSHTTPBSiteCollectionObject $MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel "ViewListItems","EditListItems","DeleteListItems","OpenItems","ViewVersions","Open","ViewPages" $false

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a document library $MSFSSHTTPFSSHTTPBDocumentLibrary in the root site $MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    CreateListItem $MSFSSHTTPFSSHTTPBSiteCollectionObject.RootWeb $MSFSSHTTPFSSHTTPBDocumentLibrary 101

    # Get the root web that is located at "http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName"
    $MSFSSHTTPFSSHTTPB_web = $MSFSSHTTPFSSHTTPBSiteCollectionObject.OpenWeb()
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $FileSyncWOPIBigTestData1 to http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSFSSHTTPFSSHTTPB_web $MSFSSHTTPFSSHTTPBDocumentLibrary $FileSyncWOPIBigTestData1 ".\$FileSyncWOPIBigTestData1" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSFSSHTTPFSSHTTPBZipTestData2 to http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSFSSHTTPFSSHTTPB_web $MSFSSHTTPFSSHTTPBDocumentLibrary $MSFSSHTTPFSSHTTPBZipTestData2 ".\$MSFSSHTTPFSSHTTPBZipTestData2" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSFSSHTTPFSSHTTPBTestData3 to http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSFSSHTTPFSSHTTPB_web $MSFSSHTTPFSSHTTPBDocumentLibrary $MSFSSHTTPFSSHTTPBTestData3 ".\$MSFSSHTTPFSSHTTPBTestData3" $true
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSFSSHTTPFSSHTTPBTestData4 to http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSFSSHTTPFSSHTTPB_web $MSFSSHTTPFSSHTTPBDocumentLibrary $MSFSSHTTPFSSHTTPBTestData4 ".\$MSFSSHTTPFSSHTTPBTestData4" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Grant full control permissions to the users $User1,$User2 and $User3 on the site $MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    GrantUserPermission $MSFSSHTTPFSSHTTPB_web "Full Control" $domain.split(".")[0] $User1
    GrantUserPermission $MSFSSHTTPFSSHTTPB_web "Full Control" $domain.split(".")[0] $User2
    GrantUserPermission $MSFSSHTTPFSSHTTPB_web "Full Control" $domain.split(".")[0] $User3

    Output "Steps for manual configuration:" "Yellow"
    Output "Grant the permission $MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel to the user $domain\$NoUseRemoteUser on the site $MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"
    GrantUserPermission $MSFSSHTTPFSSHTTPB_web $MSFSSHTTPFSSHTTPBNoUseRemotePermissionLevel $domain.Split(".")[0] $NoUseRemoteUser
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Grant read permissions to the user $domain\$ReadOnlyUser on the site $MSFSSHTTPFSSHTTPBSiteCollectionName ..." "Yellow"    
    GrantUserPermission $MSFSSHTTPFSSHTTPB_web "Read" $Domain.Split(".")[0] $ReadOnlyUser

    $MSFSSHTTPFSSHTTPBSiteCollectionObject.Dispose()
    
    #start url in IE
    StartUrlInIE $sutComputerName "http://$sutComputerName/sites/$MSFSSHTTPFSSHTTPBSiteCollectionName"
}
#-----------------------------------------------------
# Start to configure SUT for MS-WOPI
#-----------------------------------------------------
if($SharePointVersion -ge $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointFoundation2013[0])
{        
    Output "Start to run the configuration for MS-WOPI ..." "White"
    if($SharePointVersion -ge $SharePointServer2013[0])
    {
        #Creates Secure Store application item MSWOPI_TargetAppWithNotGroupAndWindows
        CreateSecureStoreServiceApplication $MSWOPITargetAppWithNotGroupAndWindows "WindowsUserName" "WindowsPassword" "Individual" $FileSyncWOPIUser $FileSyncWOPIUserPassword $domain "http://$sutComputerName" $MSWOPIUserCredentialItem $MSWOPIPasswordCredentialItem
    }
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a site collection with the name $MSWOPISiteCollectionName ..." "Yellow"
    $MSWOPISiteCollectionObject = CreateSiteCollection $MSWOPISiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" "STS#0" 1033
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a permission with the name $MSWOPINoUseRemotePermissionLevel with permissions: ViewListItems, EditListItems, DeleteListItems, OpenItems, ViewVersions, Open, and ViewPages." "Yellow"
    CreatePermissionLevel $MSWOPISiteCollectionObject $MSWOPINoUseRemotePermissionLevel "ViewListItems","EditListItems","DeleteListItems","OpenItems","ViewVersions","Open","ViewPages" $false

    Output "Steps for manual configuration:" "Yellow"
    Output "Create a document library $MSWOPISharedDocumentLibrary in the root site $MSWOPISiteCollectionName ..." "Yellow"
    CreateListItem $MSWOPISiteCollectionObject.RootWeb $MSWOPISharedDocumentLibrary 101

    # Get the root web that is located at "http://$sutComputerName/sites/$MSWOPISiteCollectionName"
    $MSWOPI_web = $MSWOPISiteCollectionObject.OpenWeb()
        
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $FileSyncWOPIBigTestData1 to http://$sutComputerName/sites/$MSWOPISiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web $MSWOPISharedDocumentLibrary $FileSyncWOPIBigTestData1 ".\$FileSyncWOPIBigTestData1" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSWOPISharedZipTestData2 to http://$sutComputerName/sites/$MSWOPISiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web $MSWOPISharedDocumentLibrary $MSWOPISharedZipTestData2 ".\$MSWOPISharedZipTestData2" $true
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSWOPISharedTestData3 to http://$sutComputerName/sites/$MSWOPISiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web $MSWOPISharedDocumentLibrary $MSWOPISharedTestData3 ".\$MSWOPISharedTestData3" $true
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSWOPISharedTestData4 to http://$sutComputerName/sites/$MSWOPISiteCollectionName ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web $MSWOPISharedDocumentLibrary $MSWOPISharedTestData4 ".\$MSWOPISharedTestData4" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Grant full control permission to the users $User1,$User2,$User3, and $FileSyncWOPIUser on the site $MSWOPISiteCollectionName..." "Yellow"
    GrantUserPermission $MSWOPI_web "Full Control" $domain.split(".")[0] $User1
    GrantUserPermission $MSWOPI_web "Full Control" $domain.split(".")[0] $User2
    GrantUserPermission $MSWOPI_web "Full Control" $domain.split(".")[0] $User3
    GrantUserPermission $MSWOPI_web "Full Control" $domain.split(".")[0] $FileSyncWOPIUser
        
    Output "Steps for manual configuration:" "Yellow"
    Output "Grant the permission $MSWOPINoUseRemotePermissionLevel to the user $domain\$NoUseRemoteUser on the site $MSWOPISiteCollectionName..." "Yellow"
    GrantUserPermission $MSWOPI_web $MSWOPINoUseRemotePermissionLevel $domain.Split(".")[0] $NoUseRemoteUser   
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Grant read permissions to the user $domain\$ReadOnlyUser on the site $MSWOPISiteCollectionName..." "Yellow"
    GrantUserPermission $MSWOPI_web "Read" $Domain.Split(".")[0] $ReadOnlyUser
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a document library $MSWOPIDocumentLibrary in the root site $MSWOPISiteCollectionName ..." "Yellow"
    CreateListItem $MSWOPISiteCollectionObject.RootWeb $MSWOPIDocumentLibrary 101
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSWOPITestData1 to http://$sutComputerName/sites/$MSWOPISiteCollectionName/$MSWOPIDocumentLibrary ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web $MSWOPIDocumentLibrary $MSWOPITestData1 ".\$MSWOPITestData1" $true
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a folder under $MSWOPISiteCollectionName with the name $MSWOPIDocumentLibrary/$MSWOPITestFolder ..." "Yellow"
    CreateSharePointFolder $MSWOPI_web "$MSWOPIDocumentLibrary/$MSWOPITestFolder"
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the test data $MSWOPITestData2 to http://$sutComputerName/sites/$MSWOPISiteCollectionName/$MSWOPIDocumentLibrary/$MSWOPITestFolder ..." "Yellow"
    UploadFileToSharePointFolder $MSWOPI_web "$MSWOPIDocumentLibrary/$MSWOPITestFolder" $MSWOPITestData2 ".\$MSWOPITestData2" $true
    
    $MSWOPISiteCollectionObject.Dispose()
            
    if($SharePointVersion -ge $SharePointServer2013[0])
    {
        #Creates Secure Store application item MSWOPI_TargetAppWithGroupAndNoWindows
        CreateSecureStoreServiceApplication $MSWOPITargetAppWithGroupAndNoWindows "UserName" "Password" "Group" $FileSyncWOPIUser $FileSyncWOPIUserPassword $domain "http://$sutComputerName" $MSWOPIUserCredentialItem $MSWOPIPasswordCredentialItem 
    }
    
    #Change the AllowOAuthOverHttp setting to true    
    $SecurityTokenConfig = Get-SPSecurityTokenServiceConfig
    $SecurityTokenConfig.AllowOAuthOverHttp = $true
    $SecurityTokenConfig.Update()
    if(!($SecurityTokenConfig.AllowOAuthOverHttp))
    {
        Throw "Failed to set the value of AllowOAuthOverHttp."
    }
    
    #Create folder by user testuser1
    $securePassword = ConvertTo-Securestring "$User1Password" -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential(($domain.Split()[0]+"\"+"$User1"),$securePassword)
    $web = "http://$sutComputerName/sites/MSWOPI_SiteCollection"
    $folderUrl = "$MSWOPIDocumentLibrary/$MSWOPIFolderCreatedByUser1"
    invoke-command -computer $sutComputerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    param
    ($web,$folderUrl)
    $SharePointShellSnapIn = Get-PSSnapin | Where-Object -FilterScript {$_.Name -eq "Microsoft.SharePoint.PowerShell"}
    if($SharePointShellSnapIn -eq $null)
    {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
    $web = Get-SPWeb $web      
    $originalFolder = $web.GetFolder($folderUrl)
    if ($originalFolder.Exists)
    {
         $originalFolder.Delete() | Out-Null
    }
    $web.Folders.Add($folderUrl) | Out-Null
    }-ArgumentList $web,$folderUrl
    
    #start url in IE
    StartUrlInIE $sutComputerName $web  
}
#-----------------------------------------------------
# Start to configure SUT for MS-ONESTORE
#-----------------------------------------------------
if($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0] -or $SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -ge $SharePointServer2013[0] )
{
    Output "Start to run configuration for MS-ONESTORE ..." "White"
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a site collection with the name $MSONESTORESiteCollectionName ..." "Yellow"
    $MSONESTORESiteCollectionNameObject = CreateSiteCollection $MSONESTORESiteCollectionName $sutComputerName "$domain\$userName" "$userName@$domain" "STS#0" 1033
      
    Output "Steps for manual configuration:" "Yellow"
    Output "Create a document library $MSONESTORELibraryName in the root site $MSONESTORESiteCollectionName ..." "Yellow"
    CreateListItem $MSONESTORESiteCollectionNameObject.RootWeb $MSONESTORELibraryName 101

    $MSONESTORE_web = $MSONESTORESiteCollectionNameObject.OpenWeb()

    Output "Steps for manual configuration:" "Yellow"
    Output "Grant full control permissions to the users $User1 on the site $MSONESTORESiteCollectionName ..." "Yellow"
    GrantUserPermission $MSONESTORE_web "Full Control" $domain.split(".")[0] $User1
    
    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOneFileWithFileData ".\$MSONESTOREOneFileWithFileData" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOneFileWithoutFileData  ".\$MSONESTOREOneFileWithoutFileData" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOneFileEncryption ".\$MSONESTOREOneFileEncryption" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOneWithInvalid ".\$MSONESTOREOneWithInvalid" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOneWithLarge ".\$MSONESTOREOneWithLarge" $true

    Output "Steps for manual configuration:" "Yellow"
    Output "Upload the file $MSONESTOREOneFileWithFileData to http://$sutComputerName/sites/$MSONESTORESiteCollectionName/$MSONESTORELibraryName ..." "Yellow"
    UploadFileToSharePointFolder $MSONESTORE_web $MSONESTORELibraryName $MSONESTOREOnetocFileLocal ".\$MSONESTOREOnetocFileLocal" $true


    $MSONESTORESiteCollectionNameObject.Dispose()
    

 }
#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
Output "The server configuration script ran successfully" "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0