#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

#------------------------------------------------------------------------------
# <param name="unattendedXmlName">The unattended SUT configuration XML.</param>
#------------------------------------------------------------------------------
param(
[string]$unattendedXmlName
)

#-----------------------------------------------------------------------
# Starting script
#-----------------------------------------------------------------------
$ErrorActionPreference = "Stop"
[String]$containerPath = & {Split-Path $MyInvocation.scriptName}
[String]$logPath       = $containerPath + "\SetupLogs"
[String]$logFile       = $logPath + "\ExchangeSUTConfiguration.ps1.log"
[String]$debugLogFile  = $logPath + "\ExchangeSUTConfiguration.ps1.debug.log"
if(!(Test-Path $logPath))
{
    Write-Host "Create a directory for storing log files." -ForegroundColor "White"
    New-Item $logPath -ItemType directory |Out-null
}
Start-Transcript $debugLogFile -Force -Append

#---------------------------------------------------------
# Configuration Variables
#---------------------------------------------------------
$userPassword                              = "Password01!" 

$MSASAIRSUser01                            = "MSASAIRS_User01"
$MSASAIRSUser02                            = "MSASAIRS_User02" 

$MSASCALUser01                             = "MSASCAL_User01"
$MSASCALUser02                             = "MSASCAL_User02" 

$MSASCMDUser01                             = "MSASCMD_User01"
$MSASCMDUser02                             = "MSASCMD_User02"
$MSASCMDUser03                             = "MSASCMD_User03"
$MSASCMDUser04                             = "MSASCMD_User04"
$MSASCMDUser05                             = "MSASCMD_User05"
$MSASCMDUser06                             = "MSASCMD_User06"
$MSASCMDUser07                             = "MSASCMD_User07"
$MSASCMDUser08                             = "MSASCMD_User08"
$MSASCMDUser09                             = "MSASCMD_User09"
$MSASCMDUser10                             = "MSASCMD_User10"
$MSASCMDUser11                             = "MSASCMD_User11"
$MSASCMDUser12                             = "MSASCMD_User12"
$MSASCMDUser13                             = "MSASCMD_User13"
$MSASCMDUser14                             = "MSASCMD_User14"
$MSASCMDUser15                             = "MSASCMD_User15"
$MSASCMDUser16                             = "MSASCMD_User16"
$MSASCMDUser17                             = "MSASCMD_User17"
$MSASCMDUser18                             = "MSASCMD_User18"
$MSASCMDUser19                             = "MSASCMD_User19"
$MSASCMDSearchUser01                       = "MSASCMD_SearchUser01"
$MSASCMDSearchUser02                       = "MSASCMD_SearchUser02"
$MSASCMDTestGroup                          = "MSASCMD_TestGroup"
$MSASCMDLargeGroup                         = "MSASCMD_LargeGroup"
$MSASCMDSharedFolder                       = "MSASCMD_SharedFolder"
$MSASCMDNonEmptyDocument                   = "MSASCMD_Non-emptyDocument.txt"
$MSASCMDEmptyDocument                      = "MSASCMD_EmptyDocument.txt"
$MSASCMDUser01Photo                        = "MSASCMD_User01Photo.jpg" 
$MSASCMDUser02Photo                        = "MSASCMD_User02Photo.jpg"
$MSASCMDPfxFileName                        = "MSASCMD_PfxFile.pfx"
$MSASCMDEmailSubjectName                   = "MSASCMD_SecureEmailForTest" 

$MSASCNTCUser01                            = "MSASCNTC_User01"
$MSASCNTCUser02                            = "MSASCNTC_User02"

$MSASCONUser01                             = "MSASCON_User01"
$MSASCONUser02                             = "MSASCON_User02"
$MSASCONUser03                             = "MSASCON_User03"

$MSASDOCUser01                             = "MSASDOC_User01"
$MSASDOCSharedFolder                       = "MSASDOC_SharedFolder"
$MSASDOCVisibleFolder                      = "MSASDOC_VisibleFolder"
$MSASDOCHiddenFolder                       = "MSASDOC_HiddenFolder"
$MSASDOCVisibleDocument                    = "MSASDOC_VisibleDocument.txt"
$MSASDOCHiddenDocument                     = "MSASDOC_HiddenDocument.txt"

$MSASEMAILUser01                           = "MSASEMAIL_User01"
$MSASEMAILUser02                           = "MSASEMAIL_User02"
$MSASEMAILUser03                           = "MSASEMAIL_User03"
$MSASEMAILUser04                           = "MSASEMAIL_User04"
$MSASEMAILUser05                           = "MSASEMAIL_User05"

$MSASHTTPUser01                            = "MSASHTTP_User01"
$MSASHTTPUser02                            = "MSASHTTP_User02"
$MSASHTTPUser03                            = "MSASHTTP_User03"
$MSASHTTPUser04                            = "MSASHTTP_User04"

$MSASNOTEUser01                            = "MSASNOTE_User01"

$MSASPROVUser01                            = "MSASPROV_User01"
$MSASPROVUser02                            = "MSASPROV_User02"
$MSASPROVUser03                            = "MSASPROV_User03"
$MSASPROVUserPolicy01                      = "MSASPROV_UserPolicy01"
$MSASPROVUserPolicy02                      = "MSASPROV_UserPolicy02"

$MSASRMUser01                              = "MSASRM_User01"
$MSASRMUser02                              = "MSASRM_User02"
$MSASRMUser03                              = "MSASRM_User03"
$MSASRMUser04                              = "MSASRM_User04"
$MSASRMADUser                              = "MSASRM_ADUser"
$MSASRMSuperUserGroup                      = "MSASRM_SuperUserGroup"
$MSASRMAllRights_AllowedTemplate           = "MSASRM_AllRights_AllowedTemplate"
$MSASRMView_AllowedTemplate                = "MSASRM_View_AllowedTemplate"
$MSASRMView_ReplyAll_AllowedTemplate       = "MSASRM_ViewReplyAll_AllowedTemplate"
$MSASRMView_Reply_AllowedTemplate          = "MSASRM_ViewReply_AllowedTemplate"
$MSASRMView_Reply_ReplyAll_AllowedTemplate = "MSASRM_ViedwReplyReplyAll_AllowedTemplate"
$MSASRMEdit_Export_NotAllowedTemplate      = "MSASRM_EditExport_NotAllowedTemplate"
$MSASRMExport_NotAllowedTemplate           = "MSASRM_Export_NotAllowedTemplate"
$MSASRMReplyAll_NotAllowedTemplate         = "MSASRMReplyAll_NotAllowedTemplate"

$MSASTASKUser01                            = "MSASTASK_User01"

$Exchange2013                              = "Microsoft Exchange Server 2013"
$Exchange2010                              = "Microsoft Exchange Server 2010"
$Exchange2007                              = "Microsoft Exchange Server 2007"

#-----------------------------------------------------------------------------------
# <summary>
# Print the content with the specified color and add the content to the log file. 
# </summary>
# <param name="content">The content to be printed.</param>
# <param name="color">The color of the content.</param>
#-----------------------------------------------------------------------------------
function Output
{
    param([string]$content, [string]$color)
    $timeString = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $timeContent = "[$timeString] $content"
    $content = $content + "`r`n"
    if (($color -eq $null) -or ($color -eq ""))
    {
        Write-Host $content -NoNewline
        Add-Content -Path $logFile -Force -Value $timeContent
    }
    else
    {
        Write-Host $content -NoNewline -ForegroundColor $color
        Add-Content -Path $logFile -Force -Value $timeContent
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Start specified services.
# </summary>
# <param name="serviceName">Service name needed to be started. Wildcards are allowed.</param>
# <param name="startMode">Start mode of the specified services needed to be started.It could be left empty and will only be used when param serviceName contains wildcards.</param>
#-----------------------------------------------------------------------------------
function StartService
{
    param(
    [string]$serviceName,
    [string]$startMode 
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    if ($serviceName -eq $null -or $serviceName -eq "")
    {
        Throw "Parameter serviceName cannot be empty."
    }
    
    $startModes = @("Auto","Manual","Disabled")
    if($startMode -ne $null -and $startMode -ne "" -and ($startModes -notcontains $startMode))
    {
        Throw "Parameter startMode should be empty or be one of the following enumerator names: $startModes."
    }

    if($serviceName.Contains('*') -or $serviceName.Contains('?'))
    {
        if($startMode -ne $null -and $startMode -ne "")
        {
            $services =  Get-WmiObject win32_service | Where-Object {($_.Name -like $serviceName) -and ($_.StartMode -eq $startMode) -and ($_.State -ne 'running')}
            if($services -ne $null)
            {
                if($startMode -eq 'Disabled')
                {
                    $services | Set-Service -StartupType Automatic
                }
            }
        }
        else
        {
            $disabledServices = Get-WmiObject win32_service | Where-Object {($_.Name -like $serviceName) -and ($_.StartMode -eq 'Disabled') -and ($_.State -ne 'running')}
            if($disabledServices -ne $null)
            {
                $disabledServices | Set-Service -StartupType Automatic
            }
            $services = Get-WmiObject win32_service | Where-Object {($_.Name -like $serviceName) -and ($_.State -ne 'running')}
        }
    }
    else
    {
        $disabledServices = Get-WmiObject win32_service | Where-Object {($_.Name -eq $serviceName) -and ($_.StartMode -eq 'Disabled') -and ($_.State -ne 'running')}
        if($disabledServices -ne $null)
        {
            $disabledServices | Set-Service -StartupType Automatic
        }
        $services = Get-WmiObject win32_service | Where-Object {($_.Name -eq $serviceName) -and ($_.State -ne 'running')}
    }
    if($services -ne $null)
    {
        $services | Start-Service
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Get user input by manually input or by reading unattended configuration XML.
# </summary>
# <param name="nodeName">Property name in unattended configuration XML.</param>
# <returns>
# user input or value read from XML.
# </returns>
#-----------------------------------------------------------------------------------
function GetUserInput
{
    param(
    [string]$nodeName
    )
    [string]$userInput = ""
    if($unattendedXmlName -eq "")
    {
        $userInput = Read-Host
    }
    else
    {
        $isNodeFound = $false
        [xml]$xmlContent = New-Object XML
        $xmlContent.Load($unattendedXmlName)
        $propertyNodes = $xmlContent.GetElementsByTagName("Property")
        foreach($node in $propertyNodes)
        {
            if($node.name -eq $nodeName)
            {
                $userInput = $node."value"
                $isNodeFound = $true
                Output "$userInput (Received from the ExchangeSUTConfigurationAnswers.xml file for property : $nodeName)." "White"
                break
            }
        }        
        if(!$isNodeFound)
        {
            Output "Cannot find node with the name attribute $nodeName in $unattendedXmlName. An empty value will be used instead." "Yellow"
        }
    }
    return $userInput
}

#-----------------------------------------------------------------------------------
# <summary>
# Check user's input until it is a valid one. 
# </summary>
# <param name="userChoices">The available number list user can select from.</param>
# <param name="nodeName">Property name in unattended configuration XML.</param> 
# <returns>
# The valid number.
# </returns>
#-----------------------------------------------------------------------------------
function ReadUserChoice
{
    param(
    [Array]$userChoices,
    [string]$nodeName
    )
    While(1)
    {
        [String]$userChoice = GetUserInput $nodeName
        if($userChoices -contains $userChoice)
        {
            return $userChoice
        }
        else
        {
            Output """$userChoice"" is not a correct input." "Yellow"
            if($unattendedXmlName -eq "")
            {
                Output "Retry with a correct number from the values listed." "Yellow"
            }
            else
            {
                Write-Warning "Change the value of $nodeName in the ExchangeSUTConfigurationAnswers.xml file with the values listed and run the script again."
                Stop-Transcript
                exit 2
            }
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Compare the recommended Exchange minor version with the installed Exchange minor version.
# </summary>
# <param name="actualVersion">The display version of the Exchange installed currently.</param>
# <param name="recommendedVersion">The recommended Exchange display version.</param> 
# <returns>
# A Boolean value, true if the server installed the recommended service pack, otherwise false.
# </returns>
#-----------------------------------------------------------------------------------           
function CompareExchangeMinorVersion
{
    param(
    [string]$actualVersion,
    [string]$recommendedVersion
    )

    $actualMinorVersion = $actualVersion.split(".")[1]
    $recommendedMinorVersion = $recommendedVersion.split(".")[1]
    $actualVersionBuildNumber = $actualVersion.split(".")[2]
    $recommendedVersionBuildNumber = $recommendedVersion.split(".")[2]
    
    if(($actualMinorVersion -eq $recommendedMinorVersion) -and ($actualVersionBuildNumber -ge $recommendedVersionBuildNumber))
    {
        return  $true
    }
    else
    {
        return  $false
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the Exchange server Version. 
# </summary>
# <returns>
# The version name of the Exchange server.
# Return "Microsoft Exchange Server 2007" if Exchange version is "Microsoft Exchange Server 2007".
# Return "Microsoft Exchange Server 2010" if Exchange version is "Microsoft Exchange Server 2010".
# Return "Microsoft Exchange Server 2013" if Exchange version is "Microsoft Exchange Server 2013".
# Others return warning on "Unknown Version" and exit script.
# </returns>
#-----------------------------------------------------------------------------------
function GetExchangeServerVersion
{
    $ExchangeServer2007             = $Exchange2007,   "8.3.83.0",      "SP3"
    $ExchangeServer2010             = $Exchange2010,   "14.3.123.4",    "SP3"
    $ExchangeServer2013             = $Exchange2013,   "15.0.847.32",   "SP1"
    $ExchangeVersion                = "Unknown Version"
 
    Output "Trying to get the Exchange server version; please wait ..." "White"
    $keys = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
    $items = $keys | foreach-object {Get-ItemProperty $_.PsPath}    
    foreach ($item in $items)
    {
        if (($item.DisplayName -eq $null) -or ($item.DisplayName -eq ""))
        {
            continue
        }
        if($item.DisplayName -eq $ExchangeServer2007[0])
        {
            $version = $item.DisplayVersion
            $ExchangeVersion = $ExchangeServer2007[0]
            $recommendVersion = $ExchangeServer2007[1]
            $recommendMinorVersion = $ExchangeServer2007[2]
            $isRecommendMinorVersion = CompareExchangeMinorVersion $version $recommendVersion
            break
        }
        if($item.DisplayName -eq $ExchangeServer2010[0])
        {
            $version = $item.DisplayVersion
            $ExchangeVersion = $ExchangeServer2010[0]
            $recommendVersion = $ExchangeServer2010[1]
            $recommendMinorVersion = $ExchangeServer2010[2]
            $isRecommendMinorVersion = CompareExchangeMinorVersion $version $recommendVersion
            break
        }        
        if($item.DisplayName.StartsWith($ExchangeServer2013[0]))
        {
            $version = $item.DisplayVersion
            $ExchangeVersion = $ExchangeServer2013[0]
            $recommendVersion = $ExchangeServer2013[1]
            $recommendMinorVersion = $ExchangeServer2013[2]
            $isRecommendMinorVersion = CompareExchangeMinorVersion $version $recommendVersion
            break
        }        
    }
    if ($ExchangeVersion -eq "Unknown Version")
    {
        Write-Warning "Could not find the supported version of Exchange server on the system! Install one of the recommended versions ($($ExchangeServer2007[0]) $($ExchangeServer2007[2]), $($ExchangeServer2010[0]) $($ExchangeServer2010[2]), $($ExchangeServer2013[0]) $($ExchangeServer2013[2])) and run the SUT configuration script again.`r`n"
        Stop-Transcript
        exit 2
    }
    else
    {
        if($isRecommendMinorVersion)
        {
            Output ("Exchange server version: $ExchangeVersion " + $recommendMinorVersion) "White"
        }
        else
        {
            Output "$ExchangeVersion $version is not the recommended version." "Yellow"
            Output ("Please install the recommended $ExchangeVersion " + "$recommendVersion, otherwise some cases might fail.") "Yellow"
            Output "Would you like to continue configuring the server or exit?" "Cyan"
            Output "1: CONTINUE." "Cyan"
            Output "2: EXIT." "Cyan"
            $runOnNonRecommendedSUTChoices = @('1','2')
            $runOnNonRecommendedSUT = ReadUserChoice $runOnNonRecommendedSUTChoices "runOnNonRecommendedSUT"
            if ($runOnNonRecommendedSUT -eq "2")
            {
                Stop-Transcript
                exit 0
            }           
        }
    }
    return $ExchangeVersion
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether Exchange server is installed on a domain controller.
# </summary>
#-----------------------------------------------------------------------------------
function CheckExchangeInstalledOnDCOrNot
{
    Import-module ActiveDirectory
    $domainControllerInfo = Get-ADDomainController | where {($_.Name -eq $env:COMPUTERNAME ) -and ($_.Domain -eq $env:USERDNSDOMAIN)}
    if(($domainControllerInfo -ne $null) -and ($domainControllerInfo -ne ""))
    {
        Output "An Exchange server installed on a domain controller is not recommended." "Yellow"
        Output "Would you like to continue to run the SUT setup script on this machine or exit?" "Cyan"    
        Output "1: CONTINUE." "Cyan"    
        Output "2: EXIT." "Cyan"    
        $runOnDomainControllerChoices = @('1','2')
        $runOnDomainController = ReadUserChoice $runOnDomainControllerChoices "runOnDomainController"
        if($runOnDomainController -eq "2")
        {
            Stop-Transcript
            exit 0
        }
        
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Set user's password never expires. 
# </summary>
# <param name="userName">The user name that will be set password never expires.</param>
#-----------------------------------------------------------------------------------
function SetPasswordNeverExpires
{
    param(
    [string]$userName
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    if ($userName -eq $null -or $userName -eq "")
    {
        Throw "Parameter userName cannot be empty."
    }
        
    try
    {
        $context = new-object System.DirectoryServices.AccountManagement.PrincipalContext("Domain") 
        $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context,$userName)
        $userPrincipal.PasswordNeverExpires = $true
        $userPrincipal.save()
         Output "Enable the setting such that Password for $userName never expires." "Green"
    }
    catch [Exception]
    {
        throw $_.Exception.Message
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Configure the SSL settings of the specified page in IIS.
# </summary>
# <param name="pageName">The page name that will be configured for SSL settings.</param>
# <param name="labelName">The label of the object that will be configured for SSL settings.</param>
# <param name="SSLStatus">The SSL status that will be configured to for the specified page.</param>
#-----------------------------------------------------------------------------------
function ConfigureSSLSettings
{
    param(
    [string]$pageName,
    [string]$labelName,
    [string]$SSLStatus
    )

    switch ($SSLStatus)
    {
        "None"{ $expectSSLStatus = $false }
        "Ssl" { $expectSSLStatus = $true }
    }
		
    $retryCount = 20
    do
    {
        cmd /c $env:windir\system32\inetsrv\appcmd.exe set config $pageName /commit:APPHOST /section:system.webServer/security/access /sslFlags:$SSLStatus
        Start-Sleep -s 3
        $EASWebSettingsObj = get-wmiobject -namespace "root/MicrosoftIISv2" -query "select * from IIsWebVirtualDirSetting where Name='$labelName'" -computer $Env:ComputerName
        $currentSSLStatus = $EASWebSettingsObj.AccessSSL
        $retryCount = $retryCount-1
    }
    while($currentSSLStatus -ne $expectSSLStatus -and $retryCount -gt 0)

    if($currentSSLStatus -ne $expectSSLStatus)
    {
        if($SSLStatus -eq "None")
        {
            Throw "Failed to clear the `"Require SSL`" and set `"Ignore`" for Client certificates in the `"SSL Settings`" page of `"$pageName`" in IIS."
        }
        else
        {
            Throw "Failed to enable the `"Require SSL`" and set `"Ignore`" for Client certificates in the `"SSL Settings`" page of `"$pageName`" in IIS."
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check if the specified mailbox user already exists in the server. 
# </summary>
# <param name="mailboxUserName">The username of the mailbox.</param>
# <returns>
# Return true if the mailbox user already exists.
# Return false if the mailbox user does not exist.
# </returns>
#-----------------------------------------------------------------------------------
function CheckMailboxUserExistOrNot
{
    param(
    [string]$mailboxUserName
    )
	
    $mailboxInfo = Get-Mailbox -filter {Name -eq $mailboxUserName}
    if(($mailboxInfo -ne $null) -and ($mailboxInfo -ne ""))
    {
        return $true
    }
    else
    {
        return $false
    }
}

#--------------------------------------------------------------------------------------
# <summary>
# Create a new mailbox user. 
# </summary>
# <param name="mailboxUserName">The name of the mailbox user.</param>
# <param name="mailboxUserPassword">The password of the mailbox user.</param>
# <param name="mailboxUserDatabase">The database of the mailbox user.</param>
# <param name="mailboxUserDomain">The domain that the mailbox user in.</param>
#-----------------------------------------------------------------------------------
function CreateMailboxUser
{
    param(
    [string]$mailboxUserName,
    [string]$mailboxUserPassword,
    [string]$mailboxUserDatabase,
    [string]$mailboxUserDomain
    )
    if(($mailboxUserName -eq $null) -or ($mailboxUserName -eq ""))
    {
        throw "Parameter mailboxUserName cannot be empty."
    }
    if(($mailboxUserPassword -eq $null) -or ($mailboxUserPassword -eq ""))
    {
        throw "Parameter mailboxUserPassword cannot be empty."
    }
    if(($mailboxUserDatabase -eq $null) -or ($mailboxUserDatabase -eq ""))
    {
        throw "Parameter mailboxUserDatabase cannot be empty."
    }
    if(($mailboxUserDomain -eq $null) -or ($mailboxUserDomain -eq ""))
    {
        throw "Parameter mailboxUserDomain cannot be empty."
    }
	
    Output "Create a mailbox user $mailboxUserName." "White"
    $exist = CheckMailboxUserExistOrNot $mailboxUserName
    if($exist -eq $true)
    {
        Output "Mailbox for $mailboxUserName already exists." "Yellow"
    }
    else
    {
       if($mailboxUserName.Length -ge 21)
        {
            Output "The mailbox username should be below 20 characters, the length of ""$mailboxUserName"" exceeds 20 characters." "Yellow"
            Output "Warning: creating the mailbox user $mailboxUserName exceeds 20 characters. This will cause an RPC connection failure." "Yellow"
            Output "Would you like to continue creating the mailbox for ""$mailboxUserName"" or exit?" "Cyan"
            Output "1: CONTINUE." "Cyan"
			Output "2: EXIT." "Cyan"
            $runWithLongerMailboxNameChoices = @('1','2')
            $runWithLongerMailboxNameChoice = ReadUserChoice $runWithLongerMailboxNameChoices "runWithLongerMailboxNameChoice"
            if ($runWithLongerMailboxNameChoice -eq "2")
            {
                Stop-Transcript
                exit 0
            }
        }
       $securePassword = ConvertTo-SecureString -String $mailboxUserPassword -AsPlainText -Force

       New-Mailbox -UserPrincipalName "$mailboxUserName@$mailboxUserDomain" -Database $mailboxUserDatabase -Name $mailboxUserName -Password $securePassword  |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        
       $check = CheckMailboxUserExistOrNot $mailboxUserName
       if($check)
       {
           Output "Mailbox for $mailboxUserName was created successfully." "Green"
       }
       else
       {
           Throw("Failed to create the mailbox for $mailboxUserName!")
       }
    }
    SetPasswordNeverExpires $mailboxUserName
}

#-----------------------------------------------------------------------------------------------------
# <summary>
# Check if the user already exists in the specified group.
# </summary>
# <param name="userName">The name of the user.</param>
# <param name="groupName">The name of the group that will be checked whether containing the user.</param>
# <returns>
# Return true if the user already exists in the specified group.
# Return false if the user does not exist in the specified group.
# </returns>
#------------------------------------------------------------------------------------------------------
function CheckGroupMember
{
    param(
    [string]$userName,
    [string]$groupName
    )
	
    $groupMembers = Get-WmiObject -Class Win32_GroupUser |where{$_.GroupComponent -like "*$groupName*"}
    if(($groupMembers -ne $null) -and ($groupMembers -ne ""))
    {
        foreach($groupMemberInfo in $groupMembers)
        {
            #An example value of $groupMember is "\\sutComputerName\root\cimv2:Win32_UserAccount.Domain="contoso",Name=$userName"
            $groupMember= $groupMemberInfo.PartComponent
            if((($groupMember.Split(",")[1]).remove(0,5)).replace('"',"") -eq $userName)
            {
                return $true
            } 
        }    
    }
	
    return $false
}

#----------------------------------------------------------------------------------
# <summary>
# Add user to the specified group.
# </summary>
# <param name="ADUser">The name of the user.</param>
# <param name="groupName">The name of group that the user will be added to.</param>
#-----------------------------------------------------------------------------------
function AddUserToGroup
{
    param(
    [string]$ADUser,
    [string]$groupName
    )
	
    if(($ADUser -eq $null) -or ($ADUser -eq ""))
    {
        throw "Parameter ADUser cannot be empty"
    }
    if(($groupName -eq $null) -or ($groupName -eq ""))
    {
        throw "Parameter groupName cannot be empty"
    }
	
    $userName=($ADUser.split("\"))[1]
    $exist =CheckGroupMember $userName $groupName
    
    if($exist)
    {
        Output "User $userName is already a member of Group $groupName." "Yellow"
    }
    else
    {
        cmd /c net.exe localgroup $groupName $ADUser /add
		
        $check =CheckGroupMember $userName $groupName
        if($check)
        {
            Output "$userName was added to $groupName group successfully." "Green"
        }
        else
        {
            Throw("Failed to add user $userName to Group $groupName!")
        }
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Check whether the specified Distribution Group exists or not.
# </summary>
# <param name="distributionGroupName">The name of the Distribution Group that will be checked.</param>
# <returns>
# Return true if the Distribution Group already exists.
# Return false if the Distribution Group does not exist.
# </returns>
#-------------------------------------------------------------------------------------------- 
function CheckDistributionGroup
{
    param(
    [String]$distributionGroupName
    ) 
	
    $distributionGroupInfo = Get-DistributionGroup | where {$_.Name -eq $distributionGroupName}
    if(($distributionGroupInfo -ne $null) -and ($distributionGroupInfo -ne ""))
    {    
        return $true        
    }
    else
    {
        return $false
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Create a distributionGroup.
# </summary>
# <param name="distributionGroupName">The name of the Distribution Group that will be created.</param>
# <param name="groupDomain">The name of the domain that the group belongs to.</param>
#-------------------------------------------------------------------------------------------- 
function NewDistributionGroup
{
    param(
    [string]$distributionGroupName,
    [string]$groupDomain
    )
	
    #------------------------------------------------------
    # Parameter validation
    #------------------------------------------------------
    if(($distributionGroupName -eq $null) -or ($distributionGroupName -eq ""))
    {
    	Throw "Parameter distributionGroupName cannot be empty."
    }
    if(($groupDomain -eq $null) -or ($groupDomain -eq ""))
    {
    	Throw "Parameter groupDomain cannot be empty."
    }
	
    $exist = CheckDistributionGroup $distributionGroupName
	
    if($exist -eq $true)
    {
        Output "Distribution Group $distributionGroupName already exists." "Yellow"
    }
    else
    {
        New-DistributionGroup -Name $distributionGroupName -Type 'Distribution' -OrganizationalUnit "$groupDomain/Users" -SamAccountName $distributionGroupName -Alias $distributionGroupName |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        $check = CheckDistributionGroup $distributionGroupName
        if($check)
        {
            Output "Distribution Group $distributionGroupName is created successfully." "Green"
        }
        else
        {
            throw "Failed to create Distribution Group $distributionGroupName."
        }
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Check whether the user is a member of the distributionGroup.
# </summary>
# <param name="distributionGroup">The name of the distributionGroup to be checked whether containing the specified user.</param>
# <param name="groupMember">The name of the user to be checked.</param>
# <returns>
# Return true if the user is already a member of the distributionGroup.
# Return false if the user is not a member of the distributionGroup.
# </returns>
#-------------------------------------------------------------------------------------------- 
function CheckDistributionGroupMember
{
    param(
    [string]$distributionGroup,
    [string]$groupMember
    )
	
    $groupMembers= Get-DistributionGroupMember -Identity $distributionGroup
    if(($groupMembers -ne $null)-and ($groupMembers -ne ""))
    {
        foreach($member in $groupMembers )
        {
            if($member.name -eq $groupMember)
            {
                return $true
            }
        }
    }
    return $false
}
	
#---------------------------------------------------------------------------------------------------------
# <summary>
# Add a user into the specified distributionGroup.
# </summary>
# <param name="distributionGroup">The name of the distributionGroup that the user will be added to.</param>
# <param name="groupMember">The name of the user to be added to the group.</param>
#--------------------------------------------------------------------------------------------------------- 
function AddDistributionGroupMember
{
    param(
    [string]$distributionGroup,
    [string]$groupMember
    )
	
    #------------------------------------------------------
    # Parameter validation
    #------------------------------------------------------
    if(($distributionGroup -eq $null) -or ($distributionGroup -eq ""))
    {
    	Throw "Parameter distributionGroup cannot be empty."
    }
    if(($groupMember -eq $null) -or ($groupMember -eq ""))
    {
    	Throw "Parameter groupMember cannot be empty."
    }
	
    $exist = CheckDistributionGroupMember $distributionGroup $groupMember
		
    if($exist -eq $true)
    {
        Output "The user $groupMember is already a member of the group $distributionGroup" "Yellow"
    }
    else
    {
        Add-DistributionGroupMember  -Identity $distributionGroup -Member $groupMember 
        $check = CheckDistributionGroupMember $distributionGroup $groupMember
	    if($check)
        {
           Output "$groupMember was added to the group $distributionGroup successfully." "Green"
        }
        else
        {
            throw "Failed to add the user $groupMember to the group $distributionGroup."
        }
        
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Check whether the Mobile Device mailbox policy exists or not.
# </summary>
# <param name="mailboxPolicyName">The name of the mobile device mailbox policy.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <returns>
# Return true if the policy already exists.
# Return false if the policy does not exist.
# </returns>
#-------------------------------------------------------------------------------------------- 
function CheckActiveSyncMailboxPolicy
{
    param(
    [string]$mailboxPolicyName,
    [string]$ExchangeVersion
    )
	
    if($ExchangeVersion -le $Exchange2010)
    {
        Get-ActiveSyncMailboxPolicy $mailboxPolicyName -ErrorAction silentlyContinue
    }
    elseif($ExchangeVersion -eq $Exchange2013)
    {
        Get-MobileDeviceMailboxPolicy $mailboxPolicyName -ErrorAction silentlyContinue
    }
    if(!$?)
    {
        if($error[0].CategoryInfo.Reason -eq "ManagementObjectNotFoundException")
        {
            return $false
        }
        else
        {
            throw $error[0]
        }
    }
    else
    {
        return $true
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Create a Mobile Device mailbox policy.
# </summary>
# <param name="mailboxPolicyName">The name of the mobile device mailbox policy.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#-------------------------------------------------------------------------------------------- 
function CreateActiveSyncMailboxPolicy
{
    param(
    [string]$mailboxPolicyName,
    [string]$ExchangeVersion
    )
	
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    if(($mailboxPolicyName -eq $null) -or ($mailboxPolicyName -eq ""))
    {
    	Throw "Parameter mailboxPolicyName cannot be empty."
    }
    if(($ExchangeVersion -eq $null) -or ($ExchangeVersion -eq ""))
    {
    	Throw "Parameter ExchangeVersion cannot be empty."
    }

    $exist = CheckActiveSyncMailboxPolicy $mailboxPolicyName
    if($exist -eq $true)
    {
        Output "The ActiveSync mailbox policy $mailboxPolicyName already exists." "Yellow"
    }
    else
    {
        if($ExchangeVersion -le $Exchange2010)
        {
            New-ActiveSyncMailboxPolicy -Name $mailboxPolicyName -AllowNonProvisionableDevices $false -DevicePasswordEnabled $false -AlphanumericDevicePasswordRequired $false -MaxInactivityTimeDeviceLock 'unlimited' -MinDevicePasswordLength $null -PasswordRecoveryEnabled $false -RequireDeviceEncryption $false -AttachmentsEnabled $true -AllowSimpleDevicePassword $true -DevicePasswordExpiration 'unlimited' -DevicePasswordHistory '0' -confirm:$false  |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        elseif($ExchangeVersion -eq $Exchange2013)
        {
            New-MobileDeviceMailboxPolicy -Name $mailboxPolicyName -AllowNonProvisionableDevices $false -PasswordEnabled $false -AlphanumericPasswordRequired $false -MaxInactivityTimeLock 'unlimited' -MinPasswordLength $null -PasswordRecoveryEnabled $false -RequireDeviceEncryption $false -AttachmentsEnabled $true -AllowSimplePassword $true -PasswordExpiration 'unlimited' -PasswordHistory '0' -confirm:$false |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        $check = CheckActiveSyncMailboxPolicy $mailboxPolicyName $ExchangeVersion
        if($check)
        {
            Output "Created ActiveSync mailbox policy $mailboxPolicyName successfully." "Green"
        }
        else
        {
            throw "Failed to create the ActiveSync mailbox policy $mailboxPolicyName."
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether this routine applies ActiveSync mailbox policy to the specified user.
# </summary>
# <param name="mailboxPolicyName">The name of the mobile device mailbox policy.</param>
# <param name="userName">The name of the user.</param>
#-----------------------------------------------------------------------------------
function CheckMailboxUserPolicy
{
    param(
    [string]$mailboxPolicyName,
    [string]$userName
    )
	
    $mailboxInfo = Get-CasMailbox -Filter {Name -eq $userName}
    if(($mailboxInfo -ne $null) -and ($mailboxInfo -ne ""))
    {
        if($mailboxInfo.ActiveSyncMailboxPolicy.Name -eq $mailboxPolicyName)
        {
            return $true
        }
    }
    return $false
}

#-----------------------------------------------------------------------------------
# <summary>
# This routine applies ActiveSync mailbox policy to the specified user.
# </summary>
# <param name="mailboxPolicyName">The name of the mobile device mailbox policy.</param>
# <param name="userName">The name of the user.</param>
#-----------------------------------------------------------------------------------
function SetMailboxUserPolicy
{
    param(
    [string]$mailboxPolicyName,
    [string]$userName
    )

    #-----------------------------------------------------
    # Parameter validation
    #-----------------------------------------------------
    if(($mailboxPolicyName -eq $null) -or ($mailboxPolicyName -eq ""))
    {
    	Throw "Parameter mailboxPolicyName cannot be empty."
    }
    if(($userName -eq $null) -or ($userName -eq ""))
    {
    	Throw "Parameter userName cannot be empty."
    }
    
    $exist = CheckMailboxUserPolicy $mailboxPolicyName $userName
    if($exist -eq $true)
    {
        Output "ActiveSync mailbox policy $mailboxPolicyName is already applied to $userName."  "Yellow"
    }
    else
    {
        Set-CasMailbox -ActiveSyncMailboxPolicy $mailboxPolicyName -Identity "$Env:UserDNSDomain/Users/$userName" 
        $check = CheckMailboxUserPolicy $mailboxPolicyName $userName
        if($check)
        {
            Output "ActiveSync mailbox policy $mailboxPolicyName is applied to $userName successfully." "Green"
        }
        else
        {
            throw "Failed to apply the ActiveSync mailbox policy $mailboxPolicyName to $userName."
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the smtp address is added to the specified user. 
# </summary>
# <param name="mailboxUser">The name of the mailbox user.</param>
# <param name="userDomain">The name of the domain that the user belongs to.</param>
# <returns>
# Return true if the smtp address is already added to the mailbox user.
# Return false if the smtp address is not added to the mailbox user.
# </returns>
#-----------------------------------------------------------------------------------
function CheckSmtpAddress
{
    param(
    [string]$mailboxUser,
    [string]$userDomain
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($mailboxUser -eq $null) -or ($mailboxUser -eq ""))
    {
    	Throw "Parameter mailboxUser cannot be empty."
    }	
    if(($userDomain -eq $null) -or ($userDomain -eq ""))
    {
    	Throw "Parameter userDomain cannot be empty."
    }
	
    $mailboxUserInfo = Get-Mailbox -Identity $mailboxUser
    $mailboxUserAddress= $mailboxUserInfo.EmailAddresses.ToArray()
    for($i=0; $i -lt $mailboxUserAddress.length; $i++)
    {
        if($mailboxUserAddress[$i].smtpAddress -eq $mailboxUser+"SMTP@"+$userDomain)
        {
            return $true
        }
    }
    return $false
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the specified folder is shared. 
# </summary>
# <param name="sharedFolderName">The name of the folder to be checked.</param>
# <returns>
# Return true if the folder is already shared.
# Return false if the folder is not shared.
# </returns>
#-----------------------------------------------------------------------------------
function CheckSharedFolder
{
    param(
    [string]$sharedFolderName
    )
	
    $shareSec = Get-WmiObject -Class Win32_LogicalShareSecuritySetting -ComputerName $Env:ComputerName
    
    foreach($shareSecFolder in $shareSec)
    {
        if($shareSecFolder.name -eq $sharedFolderName)
        {
            return $true
        }
    }
    return $false
}

#-------------------------------------------------------------------------------------------------------
# <summary>
# Grant rights for a user on the specified folder. 
# </summary>
# <param name="folderPath">The path of the folder.</param>
# <param name="grantedUser">The name of the user that the rights on the folder will be granted to.</param>
# <param name="grantedRights">The rights of the folder that will be granted to the specified user.</param>
# <param name="accessControlType">The control type of the rights which can be Allow or Deny.</param>
#-------------------------------------------------------------------------------------------------------
function GrantUserRightsOnFolder
{
    param(
    [string]$folderPath,
    [string]$grantedUser,
    [string]$grantedRights,
    [string]$accessControlType
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($folderPath -eq $null) -or ($folderPath -eq ""))
    {
    	Throw "Parameter folderPath cannot be empty."
    }
    if(($grantedUser -eq $null) -or ($grantedUser -eq ""))
    {
    	Throw "Parameter grantedUser cannot be empty."
    }
    if(($grantedRights -eq $null) -or ($grantedRights -eq ""))
    {
    	Throw "Parameter grantedRights cannot be empty."
    }
    if(($accessControlType -eq $null) -or ($accessControlType -eq ""))
    {
    	Throw "Parameter accessControlType cannot be empty."
    }
	
    $acl = Get-Acl $folderPath
    $permission = $grantedUser,$grantedRights,$accessControlType
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $folderPath
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the distributed template is added or not. 
# </summary>
# <param name="templateName">The name of the distributed template.</param>
# <returns>
# Return true if the template is already added.
# Return false if the template is not added.
# </returns>
#-----------------------------------------------------------------------------------
function CheckDistributedTemplate
{
    param(
    [string]$templateName
    )
		
    $policyTemplates= Get-ChildItem "AdrmsCluster:\RightsPolicyTemplate"
    if(($policyTemplates -ne $null) -and ($policyTemplates -ne ""))
    {
        foreach($policyTemplate in $policyTemplates)
        {
            if($policyTemplate.DefaultDisplayName -eq  $templateName)
            {
                return $true
            }
        }
    }
	
    return $false
}

#---------------------------------------------------------------------------------------------------
# <summary>
# Add a distributed template with specified rights. 
# </summary>
# <param name="templateName">The name of the distributed template.</param>
# <param name="rightsInfo">The information about the rights used in description of template.</param>
# <param name="rights">The rights the distributed template will be granted.</param>
#----------------------------------------------------------------------------------------------------
function AddDistributedTemplate
{
    param(
    [string]$templateName,
    [string]$rightsInfo,
    [string]$rights
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------	
    if(($templateName -eq $null) -or ($templateName -eq ""))
    {
    	Throw "Parameter templateName cannot be empty."
    }
    if(($rightsInfo -eq $null) -or ($rightsInfo -eq ""))
    {
    	Throw "Parameter rightsInfo cannot be empty."
    }
    if(($rights -eq $null) -or ($rights -eq ""))
    {
    	Throw "Parameter rights cannot be empty."
    }
    $exist = CheckDistributedTemplate $templateName
    if($exist -eq $true)
    {
        Output "The distributed template $templateName already exists." "Yellow"
    }
    else
    {
        New-Item AdrmsCluster:\RightsPolicyTemplate -LocaleName en-us -DisplayName $templateName -Description "$rightsInfo" -UserGroup ANYONE -Right $rights
        Output "The distributed template $templateName is added successfully." "Green"
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Check if the user already exists in the Organization Management group.
# </summary>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of the user.</param>
# <returns>
# Return true if the user already exists in the Organization Management group.
# Return false if the user does not exist in the Organization Management group.
# </returns>
#----------------------------------------------------------------------------------
function CheckOrgAdminMember
{
    param(
    [string]$ExchangeVersion,
    [string]$userName
    )
	
    if($ExchangeVersion -eq $Exchange2007)
    {
        $orgAdminGroup = "OrgAdmin"
        $orgAdminRoleInfo = Get-ExchangeAdministrator -Identity $userName | where {$_.Role -eq $orgAdminGroup}
        if($orgAdminRoleInfo -ne $null)
        {
            return $true
        }
        return $false
    }
    elseif($ExchangeVersion -ge $Exchange2010)
    {
        $orgAdminGroup= "Organization Management"
        $orgAdminRoleInfo = Get-RoleGroupMember -Identity $orgAdminGroup | where {$_.Name -eq $userName}
        if($orgAdminRoleInfo -ne $null)
        {
            return $true
        }
        return $false
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Add user to the Organization Management group.
# </summary>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of the user.</param>
#-----------------------------------------------------------------------------------
function  AddUserToOrgMgmtGroup
{
    param(
    [string]$ExchangeVersion,
    [string]$userName
    )

    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($ExchangeVersion -eq $null) -or ($ExchangeVersion -eq ""))
    {
        throw "Parameter ExchangeVersion cannot be empty"
    }
    if(($userName -eq $null) -or ($userName -eq ""))
    {
        throw "Parameter userName cannot be empty"
    }
	
    $exist = CheckOrgAdminMember $ExchangeVersion $userName
	
    if($exist)
    {
        Output "The user $userName already exists in the Organization Management group." "Yellow"
    }
    else
    {
        if($ExchangeVersion -eq $Exchange2007)
        {
            $orgAdminGroup = "OrgAdmin"
            Add-ExchangeAdministrator -Role $orgAdminGroup -Identity $userName
        }
		
        elseif($ExchangeVersion -ge $Exchange2010)
        {
            $orgAdminGroup = "Organization Management"
            Add-RoleGroupMember -Identity $orgAdminGroup -member $userName -BypassSecurityGroupManagerCheck 
        }

        $check = CheckOrgAdminMember $ExchangeVersion $userName
        if($check)
        {
            Output "Added the user $userName to Organization Management Group successfully." "Green"
        }
        else
        {
            Throw("Failed to add user $useName to Organization Management Group!")
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a domain user.
# </summary>
# <param name="userName">The name of the user to be created.</param>
# <param name="password">The password of the user to be created.</param>
#-----------------------------------------------------------------------------------
function CreateADUser
{
    param(
    [string]$userName,
    [string]$password
    )

    #----------------------------------------------------------------------------
    # Check if the specific user exists. If not, create the user.
    #----------------------------------------------------------------------------
    Invoke-Command{
	
        $ErrorActionPreference = "Continue"
        cmd /c net user /domain $userName 2>&1 | Out-Null
        if (!$?)
        {
            cmd /c net user /domain $userName $password /add 2>&1 | Out-Null
            if (!$?)
            {
                Throw "Failed to create user $userName."
            }
            else
            {
                Output "User $username is created successfully." "Green"
            }
        }
        else
        {
            Output "User $userName already exists." "Yellow"
        } 
    }
    #Set the password of $userName into never expired.
    SetPasswordNeverExpires $userName
}

#-------------------------------------------------------------------------------------------------------------
# <summary>
# Add a record in DNS manager. 
# </summary>
# <param name="sutComputerName">The name of the server that the Microsoft Exchange Server installed on.</param>
#--------------------------------------------------------------------------------------------------------------
function AddDNSResourceRecord
{
    param(
    [string]$sutComputerName
    )
    
    $domainControllerInfo = Get-ADDomainController |where {$_.Enabled -eq $true}
    $domainControllerHostName = $domainControllerInfo.HostName
    $domainName = $domainControllerInfo.Domain
    $rec = [WmiClass]"\\$domainControllerHostName\root\MicrosoftDNS:MicrosoftDNS_ResourceRecord"
    $serverIPs = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -Computer $sutComputerName | Select-Object -Property IPAddress
    $newUrl = "rms."+$domainName
    foreach($serverIP in $serverIPs)
    {
        $address= ($serverIP.IPAddress)[0]
        $text = "$newUrl IN A $address"
        $rec.CreateInstanceFromTextRepresentation($sutComputerName,$domainName, $text)|Out-Null
        $resourceRecordInfo = Get-WmiObject -namespace "root\MicrosoftDNS" -class MicrosoftDNS_ResourceRecord -ComputerName $domainControllerHostName| Where-Object {$_.TextRepresentation -eq $text}
        if($resourceRecordInfo -ne $null -and $resourceRecordInfo -ne "")
        {
            Output "The host record rms.$domainName is created successfully." "Green"
        }
        else
        {
            throw "Failed to add host record rms.$domainName."
        }
            
    }  
}

#----------------------------------------------------------------------------------------------------------------------------------------
# <summary>
# Add Read and Execute permission for the specified group on the specified file. 
# </summary>
# <param name="securityGroup">The name of the group to be granted Read and Execute permission on the specified file.</param>
# <param name="filePath">The path of the file.</param>
#------------------------------------------------------------------------------------------------------------------------------------------
function AddAcl
{
    param(
    [string]$securityGroup,
    [string]$filePath
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($securityGroup -eq $null) -or ($securityGroup -eq ""))
    {
    	Throw "Parameter securityGroup cannot be empty."
    }
    if(($filePath -eq $null) -or ($filePath -eq ""))
    {
    	Throw "Parameter filePath cannot be empty."
    }
	
   $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($securityGroup,"Read,ReadAndExecute","None","None","Allow")
   $DirInfo = New-Object System.IO.DirectoryInfo($filePath)
   $Acl = $DirInfo.GetAccessControl()
   $Acl.AddAccessRule($AccessRule)
   $DirInfo.SetAccessControl($Acl)
   Output "Added Read and ReadAndExecute permission for the group $securityGroup on the file successfully." "Green"
}

#-------------------------------------------------------------------------------------------------------------
# <summary>
# Trust SUT computer for delegation. 
# </summary>
# <param name="sutComputerName">The name of the server that the Microsoft Exchange Server installed on.</param>
# <param name="domainName">The name of the domain.</param>
#--------------------------------------------------------------------------------------------------------------
function TrustComputerForDelegation
{
    param(
    [string]$sutComputerName,
    [string]$domainName
    )
	
	$domainControllerInfo = Get-ADDomainController | where {($_.Name -eq $sutComputerName) -and ($_.Domain -eq $domainName)}
    if(($domainControllerInfo -eq $null) -or ($domainControllerInfo -eq ""))
    {
        Output "The computer is not a domain controller"
        $TRUSTED_FOR_DELEGATION = 524288;
        $domainContext="GC://" + $([adsi] "LDAP://RootDSE").Get("RootDomainNamingContext")
        $contextInfo = New-Object System.DirectoryServices.DirectoryEntry($domainContext)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = $contextInfo
        $searcher.Filter = "(cn=$sutComputerName)"
        $results = $searcher.FindAll()
        if($results.count -eq 0)
        { 
            Output "Computer $sutComputerName is not found in Active Directory." "Red"
        }
        else
        {
            foreach ($result in $results)
            {
                if(($result.path).contains("CN=Computers"))
                {
                    $dn=[string]$($result.properties["adspath"]).replace("GC://","LDAP://")
                    $computerInfo=New-Object System.DirectoryServices.DirectoryEntry($dn)
                    Output "Enable the setting such that $($computerInfo.cn) is trusted for delegation..." "Green"
                    $uac=$computerInfo.userAccountControl[0] -bor $TRUSTED_FOR_DELEGATION
                    $computerInfo.userAccountControl[0]=$uac
                    $result=$computerInfo.CommitChanges()
                }
            }
        }
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Check whether the mailbox user's photo exists or not.
# </summary>
# <param name="mailboxUser">The name of the mailbox user.</param>
# <returns>
# Return true if the mailbox user's photo already exists.
# Return false if the mailbox user's photo does not exist.
# </returns>
#-------------------------------------------------------------------------------------------- 
function CheckMailboxUserPhoto
{
    param(
    [string]$mailboxUser
    )
	
    $userInfo = Get-ADUser $mailboxUser -Properties thumbnailphoto
    if($userInfo.thumbnailphoto -is [array])
    {
        return $true
    }
    else
    {
        return $false
    }
}

#-------------------------------------------------------------------------------------------
# <summary>
# Add photo to mailbox user.
# </summary>
# <param name="mailboxUser">The name of the mailbox user for which the photo is to be added.</param>
# <param name="userPhotoName">The name of photo that will be added to the mailbox user.</param>
#--------------------------------------------------------------------------------------------
function AddPhotoToMailboxUser
{
    param(
    [string]$mailboxUser,
    [string]$userPhotoName
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($mailboxUser -eq $null) -or ($mailboxUser -eq ""))
    {
        throw "Parameter mailboxUser cannot be empty"
    }
    if(($userPhotoName -eq $null) -or ($userPhotoName -eq ""))
    {
        throw "Parameter userPhotoName cannot be empty"
    }

    $exist = CheckMailboxUserPhoto $mailboxUser
    if($exist -eq $true)
    {
        Output "The photo is already added into the mailbox user $mailboxUser." "Yellow"
    }
    else
    {
        Import-RecipientDataProperty -Identity $mailboxUser -Picture -FileData ([Byte[]]$(Get-Content -Path ".\$userPhotoName" -Encoding Byte -ReadCount 0))
        $check = CheckMailboxUserPhoto $mailboxUser
        if($check -eq $true)
        {
            Output "Added the photo $userPhotoName to mailbox user $mailboxUser successfully." "Green"
        }
        else
        {
            Throw "Failed to add the photo $userPhotoName to mailbox user $mailboxUser."
        }
    }
}

#----------------------------------------------------------------------------------------------------------------------------------------
# <summary>
# Add delegate of mailbox user to another mailbox user. 
# </summary>
# <param name="mainMailboxUser">The name of mailbox user that will grant the delegate permission.</param>
# <param name="mainMailboxUserPassword">The password of the mailbox user.</param>
# <param name="delegateMailboxUser">The name of the mailbox user that will be assigned the delegate permission.</param>
# <param name="sutComputerName">The name of the server that the Microsoft Exchange Server installed on.</param>
# <param name="domainName">The name of the domain.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#--------------------------------------------------------------------------------------------------------------------------------
function AddDelegateForMaiboxUser
{
    param(
    [string]$mainMailboxUser,
    [string]$mainMailboxUserPassword,
    [string]$delegateMailboxUser,
    [string]$sutComputerName,
    [string]$domainName,
    [string]$ExchangeVersion
    )
	
	$dllPath= & {Split-Path $MyInvocation.scriptName}
    if(!(Test-Path "$dllPath\MS_OXWSDLGM_ServerAdapter.dll"))
    {
        Output "The file MS_OXWSDLGM_ServerAdapter.dll is not found, the case related with delegate can not be tested." "Red"
    }
    else
    {
        $asm=[Reflection.Assembly]::LoadFrom("$dllpath\MS_OXWSDLGM_ServerAdapter.dll")
        $delegateInstance = New-Object Microsoft.Protocols.TestSuites.OXWSDLGM.OXWSDLGMAdapter
        if($ExchangeVersion -eq $Exchange2007)
        {
            $version = "Exchange2007_SP3"
        }   
        elseif($ExchangeVersion -ge $Exchange2010)
        {
            $version = "Exchange2010_SP3"
        }
        $delegateInfo= $delegateInstance.AddDelegate($mainMailboxUser, $mainMailboxUserPassword, $delegateMailboxUser, "Https", $sutComputerName, "/ews/exchange.asmx", $domainName, $version)
        if($delegateInfo -eq "NoError")
        {
            Output "Added the delegate of mailbox user $mainMailboxUser to mailbox user $delegateMailboxUser successfully." "Green"
        }
        elseif($delegateInfo.contains("DelegateAlreadyExists"))
        {
            Output "The delegate of mailbox user $mainMailboxUser has already been set to mailbox user $delegateMailboxUser." "Yellow"
        }
	    else
        {
            throw "Failed to add the delegate of mailbox user $mainMailboxUser to $delegateMailboxUser."
        }  
    }
}

#-----------------------------------------------------------------------
# <summary>
# Install Active Directory Certificate Services role.
# </summary>
#-----------------------------------------------------------------------
function InstallADCSRole 
{
    Output "Install and configure the Active Directory Certificate Services role." "White"
    $os = Get-WmiObject -class Win32_OperatingSystem -computerName $env:COMPUTERNAME	
    if([int]$os.BuildNumber -le 7601)
    {
        Import-Module ServerManager
        if ((Get-WindowsFeature ADCS-Cert-Authority | Where-Object {$_.Installed -match "False"}) -and (Get-WindowsFeature ADCS-Web-Enrollment | Where-Object {$_.Installed -match "False"}))
        {
            Add-WindowsFeature ADCS-Cert-Authority |Out-Null
            Add-WindowsFeature ADCS-Web-Enrollment |Out-Null
            Output "Configuring settings for ADCS..." "White"
            #Setting CA type. ENTERPRISE_ROOTCA=0, ENTERPRISE_SUBCA=1, STANDALONE_ROOTCA=3, STANDALONE_SUBCA=4
            [int32]$catype="0"
            #Setting CA years
            [int32]$cayears="5"
            #Setting CA common name
            $cacommonname="Enterprise CA"
            $certocm = New-Object -ComObject "certocm.certsrvsetup"
            $certocm.InitializeDefaults($true,$false)
            $certocm.SetCASetupProperty(0,$catype)
            $certocm.SetCASetupProperty(6,$cayears)
            $certocm.SetCADistinguishedName("CN=$cacommonname",1,1,1)
            $certocm.Install()	
        }
        if (Get-WindowsFeature SMTP-Server | Where-Object {$_.Installed -match "False"})
        {
            Add-WindowsFeature SMTP-Server -IncludeAllSubFeature |Out-Null
        }
    }
    elseif([int]$os.BuildNumber -ge 9200)
    {
        if ((Get-WindowsFeature ADCS-Cert-Authority | Where-Object {$_.Installed -match "False"}) -and (Get-WindowsFeature ADCS-Web-Enrollment | Where-Object {$_.Installed -match "False"}))
        {
            Install-WindowsFeature -Name ADCS-Cert-Authority |Out-Null
            Install-WindowsFeature -Name ADCS-Web-Enrollment |Out-Null
            Output "Configuring settings for ADCS..." "White"
            Install-AdcsCertificationAuthority -confirm:$false |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
            Install-AdcsWebEnrollment -confirm:$false |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        if (Get-WindowsFeature SMTP-Server | Where-Object {$_.Installed -match "False"})
        {
           Install-WindowsFeature SMTP-Server -IncludeAllSubFeature |Out-Null
        }
    }
    $currentPath= & {Split-Path $MyInvocation.scriptName}
    $PolicyFilePath="$currentPath\cert.inf"
    $lines = [System.IO.File]::ReadAllLines("$PolicyFilePath");
    for ($i = 0; $i -lt $lines.Length; $i++)
    {
        $line = $lines[$i];
        if ($line -imatch "Subject")
        {
            $lines[$i] = "Subject = `"CN=$ENV:COMPUTERNAME.$ENV:USERDNSDOMAIN`""
            break;
        }
    }
    [System.IO.File]::WriteAllLines("$PolicyFilePath",$lines)
}

#-------------------------------------------------------------------------------------------
# <summary>
# Create a certificate for a mailbox user.
# </summary>
# <param name="mailboxUserName">The name of mailbox user that the certificate to be used.</param>
# <param name="userPassword">The password of mailbox user that the certificate to be used.</param>
# <param name="certFolderPath">The path of the cert file.</param>
# <param name="pfxFileName">The name of pfx file.</param>
#--------------------------------------------------------------------------------------------
[ScriptBlock] $createCert={
    param(
    [string]$mailboxUserName,
    [string]$userPassword,
    [string]$certFolderPath,
    [string]$pfxFileName
    ) 

    $policyFile="$certFolderPath\cert.inf";
    $requestFile="$certFolderPath\requestFile.req";
    $certFile="$certFolderPath\certFile.cer";
    $pfxFile ="$certFolderPath\$pfxFileName";
    certreq -new -f -q $policyFile $requestFile;
    certreq -submit -f -q $requestFile $certFile;
    certreq -accept $certFile;
    Import-Module ActiveDirectory;
    $userInfo = Get-ADUser $mailboxUserName -Properties "Certificates";
    $userCertificates = $userInfo.Certificates | foreach {New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $_};
    if(($userCertificates -eq $null) -and ($userCertificates -eq ""))
    {
        throw "Failed to create the personal certificate for mailbox user $mailboxUserName."
    }
    else
    {
	    if($userCertificates -is [array])
        {
             certutil -user -f -p $userPassword -exportPFX my $userCertificates[0].Thumbprint $pfxFile;	 
        }
        else
        {
             certutil -user -f -p $userPassword -exportPFX my $userCertificates.Thumbprint $pfxFile;	
        }
    }
}

#---------------------------------------------------------------------------------------------------------
# <summary>
# Send a secure email to mailbox user.
# </summary>
# <param name="serverName">The name of the server that the Microsoft Exchange Server installed on.</param>
# <param name="fromUserName">The name of mailbox user who sends the email.</param>
# <param name="userPassword">The password of mailbox user who sends the email.</param>
# <param name="sendToUserName">The name of mailbox user that the email sent to.</param>
# <param name="pfxPath">The path of the encryption certificate.</param>
# <param name="pfxFileName">The file name of the encryption certificate.</param>
# <param name="emailSubject">The subject name of email.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#---------------------------------------------------------------------------------------------------------
[ScriptBlock] $SendSecureEmail={

    param(
    [string]$serverName,
    [string]$fromUserName,
    [string]$userPassword,
    [string]$sendToUserName,
    [string]$pfxPath,
    [string]$pfxFileName,
    [string]$emailSubject,
    [string]$ExchangeVersion
    )
    [void][reflection.assembly]::LoadWithPartialName("System.Security");
    $x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("$pfxPath\$pfxFileName", $userPassword);
    $body = new-object System.Text.StringBuilder;
    $body.AppendLine("Content-Type: text/plain;charset=iso-8859-1");
    $body.AppendLine("Content-Transfer-Encoding: quoted-printable");
    $body.AppendLine();
    $body.AppendLine("This is a test s/mime message");
    $enc = [System.Text.Encoding]::ASCII;
    $data = $enc.GetBytes($body);
    $contentinfo = new-object security.Cryptography.Pkcs.ContentInfo -argumentList (,$data);
    $cms = new-object system.security.cryptography.pkcs.signedcms($contentinfo, $false);
    $cmssigner = new-object System.Security.Cryptography.Pkcs.CmsSigner([System.Security.Cryptography.Pkcs.SubjectIdentifierType]::IssuerAndSerialNumber, $x509);
    $cmssigner.IncludeOption = [System.Security.Cryptography.X509Certificates.X509IncludeOption]::WholeChain;
    $signtime = New-Object System.Security.Cryptography.Pkcs.Pkcs9SigningTime;
    $cmssigner.SignedAttributes.Add($signtime);
    $cms.ComputeSignature($cmssigner, $false);
    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = New-Object System.Net.Mail.MailAddress($fromUserName);
    $msg.To.Add($sendToUserName);
    $msg.Subject = $emailSubject;
    [byte[]] $bytes = $cms.Encode();
    $ms = New-Object System.IO.MemoryStream(,$bytes);
    $av = New-Object System.Net.Mail.AlternateView($ms, "application/pkcs7-mime; smime-type=signed-data;name=smime.p7s");
    $msg.AlternateViews.Add($av);
    $ServerAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -Computer $env:COMPUTERNAME
    if($ServerAdapters -is [array] -and $ExchangeVersion.contains("2013"))
    {
        $smtp = New-Object System.Net.Mail.SmtpClient($serverName, 2525);
    }
    else
    {
        $smtp = New-Object System.Net.Mail.SmtpClient($serverName, 25);
    }
    $smtp.UseDefaultCredentials = $true;
    $smtp.Send($msg)
}

#----------------------------------------------------------------------------------------------------------------------------------------------------------
# <summary>
# Move meeting forward notification email to Deleted Items for specified mailbox user. 
# </summary>
# <param name="mailboxUser">The name of mailbox user that will enable the setting to move meeting forward notification email to Deleted Items.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#----------------------------------------------------------------------------------------------------------------------------------------------------------
function MoveNotificationEmailToDeleteFolder
{
    param(
    [string]$mailboxUser,
    [string]$ExchangeVersion
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($mailboxUser -eq $null) -or ($mailboxUser -eq ""))
    {
    	Throw "Parameter mailboxUser cannot be empty."
    }
    if(($ExchangeVersion -eq $null) -or ($ExchangeVersion -eq ""))
    {
    	Throw "Parameter ExchangeVersion cannot be empty."
    }
	
    if($ExchangeVersion -eq $Exchange2007)
    {
        $calendarAttendantConfig = Get-MailboxCalendarSettings -Identity $mailboxUser   
        if($calendarAttendantConfig.RemoveForwardedMeetingNotifications -eq $false)
        {
            Set-MailboxCalendarSettings -Identity $mailboxUser -RemoveForwardedMeetingNotifications:$true  
            Output "Enabled the setting such that moving meeting forward notification email to the Deleted Items folder for mailbox user $mailboxUser successfully." "Green"  
        }
        else
        {
            Output "Setting to move meeting forward notification email to the Deleted Items folder for mailbox user $mailboxUser has already been set." "Yellow"
        }
    }
    elseif($ExchangeVersion -ge $Exchange2010)
    {
        if ($ExchangeVersion -eq $Exchange2010) 
        {
            $connectUri = 'http://'+ $sutComputerName + '/PowerShell'
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectUri -Authentication Kerberos  
            Import-PSSession $session -AllowClobber -DisableNameChecking |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
            Set-AdminAuditLogConfig -AdminAuditLogEnabled $false
        }
        $calendarAttendantConfig = Get-CalendarProcessing -Identity $mailboxUser
        if($calendarAttendantConfig.RemoveForwardedMeetingNotifications -eq $false)
        {
            Set-CalendarProcessing -Identity $mailboxUser -RemoveForwardedMeetingNotifications:$true 
            Output "Enabled the setting such that moving meeting forward notification email to the Deleted Items folder for mailbox user $mailboxUser successfully." "Green"  
        }
        else
        {
            Output "Setting to move meeting forward notification email to the Deleted Items folder for mailbox user $mailboxUser has already been set." "Yellow"
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Uninstall the Active Directory Right Management Services role. 
# </summary>
# <param name="domain">The name of domain.</param>
#-----------------------------------------------------------------------------------
function UninstallRoleADRMS
{
    $adrmsModuleInfo = Get-Module -Name ADRMS
    if($adrmsModuleInfo -eq $null)
    {
        Import-Module ADRMS
    }
    Uninstall-ADRMS -Force	
    $featureStatus = Remove-WindowsFeature ADRMS
    $configurationContext = ([ADSI]"LDAP://RootDSE").Get("ConfigurationNamingContext")
    $rmsPath=[ADSI]"LDAP://CN=RightsManagementServices,CN=Services,$configurationContext"
    if($rmsPath.Path -ne $null)
    {
        $rmsPath.DeleteTree()
    }
    if($featureStatus.RestartNeeded -eq "Yes")
    {
        $locationPath = (Get-Location).Path
        Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "CMD" -Value "$psHome\powershell.exe  `"& '$locationPath\ExchangeSUTConfiguration.cmd' '$unattendedXmlName'`""
        if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
        {    
            Output "A system restart will be required, please press enter when you are ready" "Cyan"
            cmd /c pause 
        }
        shutdown -r -f -t 3
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Install the Active Directory Right Management Services role. 
# </summary>
# <param name="ADUser">The name of the domain user account.</param>
# <param name="ADUserPassword">The password of the domain user account.</param>
#-----------------------------------------------------------------------------------
function InstallRoleADRMS
{
    param(
    [string]$ADUser,
    [string]$ADUserPassword
    )    
	
    $newUrl = "rms."+$Env:UserDNSDomain
    $os = Get-WmiObject -class Win32_OperatingSystem -computerName $env:COMPUTERNAME
    if([int]$os.BuildNumber -ge 9200)
    {
        Output "Add AD RMS role services and tools..." "White"
        Install-WindowsFeature ADRMS -IncludeManagementTools   |Out-Null
        Install-WindowsFeature NET-Framework-Core |Out-Null
    }
    Import-Module ADRMS
    New-PSDrive -PSProvider ADRMSInstall -Name RC -Root RootCluster -Scope Global |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    Set-ItemProperty -Path RC:\ClusterDatabase -Name UseWindowsInternalDatabase -Value $true
    $securepw=convertto-securestring $ADUserPassword -asplaintext -force 
    $Credential_adrmsadmin = New-Object system.management.automation.pscredential($ADUser,$securepw)    
    Set-ItemProperty -Path RC:\ -Name ServiceAccount -Value $Credential_adrmsadmin
    Set-ItemProperty -Path RC:\ClusterKey -Name UseCentrallyManaged -Value $true
    Set-ItemProperty -Path RC:\ClusterKey -Name CentrallyManagedPassword -Value $securepw
    Set-ItemProperty -Path RC:\ClusterWebSite -Name WebSiteName -Value "Default Web Site"
    Set-ItemProperty -Path RC:\ -Name ClusterURL -Value "http://$newUrl`:80"
    Set-ItemProperty -Path RC:\ -Name SLCName -Value $Env:ComputerName
    Set-ItemProperty -Path RC:\ -Name RegisterSCP -Value $true
    Install-ADRMS -Path RC:\ -Force -ErrorAction silentlyContinue    
    if(!$?)
    {
        $errorInfo = $Error[0]
        Output $errorInfo "Red"		
        Output "The installation of Active Directory Right Management Service role encountered an error, now removing the incomplete installation and re-installing it later." "Yellow"
        UninstallRoleADRMS $domain
		
        Output "Re-installing Active Directory Right Management Service role."  "Yellow"
        Install-ADRMS -Path RC:\ -Force
    }
		
    cmd /c net.exe localgroup "AD RMS Enterprise Administrators" $ADUser  /add
    ConfigureSSLSettings "Default Web Site/_wmcs/admin" "W3SVC/1/ROOT/_wmcs/admin" "Ssl"
    $locationPath = (Get-Location).Path
    Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "CMD" -Value "$psHome\powershell.exe  `"& '$locationPath\ExchangeSUTConfiguration.cmd' '$unattendedXmlName'`""
    if($unattendedXmlName -eq "" -or $unattendedXmlName -eq $null)
    {    
       Output "A system restart will be required, please press enter when you are ready" "Cyan"
       cmd /c pause 
    }
    shutdown -r -f -t 3
    Start-Sleep -s 6
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the Active Directory Right Management Services role is installed or not.
# </summary>
# <param name="ADUser">The name of the domain user account.</param>
# <param name="ADUserPassword">The password of the domain user account.</param>
#-----------------------------------------------------------------------------------
function CheckRoleADRMS
{
    param(
    [string]$ADUser,
    [string]$ADUserPassword
    )
	
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------
    if(($ADUser -eq $null) -or ($ADUser -eq ""))
    {
    	Throw "Parameter ADUser cannot be empty."
    }
    if(($ADUserPassword -eq $null) -or ($ADUserPassword -eq ""))
    {
    	Throw "Parameter ADUserPassword cannot be empty."
    }
    $os = Get-WmiObject -class Win32_OperatingSystem -computerName $env:COMPUTERNAME
    if([int]$os.BuildNumber -le 7601)
    {
        Import-Module ServerManager
    }
    $global:ADRMSInstalledFlag = $false
    $adRMSinfo = Get-WindowsFeature ADRMS
    $adRMS= $adRMSinfo.installed
    if($adRMS -eq $true)
    {
        Import-Module WebAdministration
        if(Test-Path "IIS:\Sites\default web site\_wmcs")
        {
            Output "In the `"SSL Settings`" page of `"Default Web Site/Default Web Site/_wmcs/admin`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow" 
            ConfigureSSLSettings "Default Web Site/_wmcs/admin" "W3SVC/1/ROOT/_wmcs/admin" "None"
        }
		
        # Start Windows Internal Database service
        if([int]$os.BuildNumber -le 7601)
        {
            Get-Service | where {$_.Name -eq 'MSSQL$MICROSOFT##SSEE'} | Start-Service
        }
        else
        {
            Get-Service | where {$_.Name -eq 'MSSQL$MICROSOFT##WID'} | Start-Service
        }
        Import-Module AdRmsAdmin
        New-PSDrive -Name AdrmsCluster -PsProvider AdRmsAdmin -Root http://localhost -Scope Global -ErrorAction silentlyContinue -force |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        if(!$?)
        {
            $errorInfo = $Error[0]
            Output $errorInfo "Red"			
            Output "The installation of Active Directory Right Management Service role encountered an error, now removing it and re-installing it later." "Yellow"
            UninstallRoleADRMS 
		
            Output "Re-installing Active Directory Right Management Service role." "Yellow"
            InstallRoleADRMS $ADUser $ADUserPassword
        }
        else
        {
            Output "The Active Directory Right Management Services role has already been installed." "Yellow"
            $exist = Get-RmsSvcAccount -path "AdrmsCluster:\" | foreach-Object{$_.userName -eq $ADUser}
            if($exist -eq $false)
            {
                #Update the service account of an Active Directory Rights Management Services (AD RMS) cluster
                $securepw=convertto-securestring $ADUserPassword -asplaintext -force 
                $Credential_adrmsadmin = New-Object system.management.automation.pscredential($ADUser,$securepw) 
                Set-RmsSvcAccount -Path "AdrmsCluster:\" -Credential $Credential_adrmsadmin -force
            }
            $global:ADRMSInstalledFlag = $true 
        }
    }    
    else
    {
        Output "Installing the Active Directory Right Management Services role."  "Yellow" 
        InstallRoleADRMS $ADUser $ADUserPassword
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Configure the Active Directory Right Management Services role. 
# </summary>
# <param name="distributionGroup">The name of the distribution group.</param>
# <param name="ExchangeVersion">The name of the Exchange server version.</param>
#-----------------------------------------------------------------------------------
function ConfigureRoleADRMS
{
    param(
    [string]$distributionGroup,
    [string]$ExchangeVersion
    ) 
   
    #--------------------------------------------------
    # Parameter validation
    #--------------------------------------------------	
    if(($distributionGroup -eq $null) -or ($distributionGroup -eq ""))
    {
    	Throw "Parameter distributionGroup cannot be empty."
    }
    if(($ExchangeVersion -eq $null) -or ($ExchangeVersion -eq ""))
    {
    	Throw "Parameter ExchangeVersion cannot be empty."
    }
	
    if($ExchangeVersion -ge $Exchange2010)
    {
        Output "Add RMS shared identity user into distribution group $distributionGroup." "White"
        AddDistributionGroupMember $distributionGroup "FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042"	
    }
    
    Import-Module AdRmsAdmin
    New-PSDrive -Name AdrmsCluster -PsProvider AdRmsAdmin -Root http://localhost -Scope Global |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    Set-ItemProperty -Path AdrmsCluster:\ -Name IntranetLicensingURL -Value http://localhost/_wmcs/licensing -Force
    Set-ItemProperty -Path AdrmsCluster:\ -Name ScpUrl -Value http://localhost/_wmcs/certification -Force
    Set-ItemProperty -Path AdrmsCluster:\SecurityPolicy\SuperUser -Name IsEnabled -Value $true
    Set-ItemProperty -Path AdrmsCluster:\SecurityPolicy\SuperUser -Name SuperUserGroup -Value "$distributionGroup@$Env:UserDNSDomain"
}

#-------------------------------------------------------------------------------------------
# Check whether the unattended SUT configuration XML is available if run in unattended mode.
#-------------------------------------------------------------------------------------------
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
# Get Exchange server basic information.
#-----------------------------------------------------
$domain          = $Env:UserDNSDomain
Output "Domain name: $domain" "White"
$sutComputerName = $Env:ComputerName
Output "The name of the Exchange server: $sutComputerName" "White"
$userName        = $Env:UserName
Output "The logon name of the current user: $userName " "White"
$ExchangeVersion = GetExchangeServerVersion

#-----------------------------------------------------
# Add Exchange PowerShell snapin.
#-----------------------------------------------------
if($ExchangeVersion -ge $Exchange2010)
{
    $ExchangeShellSnapIn = "Microsoft.Exchange.Management.PowerShell.E2010"	
}
if($ExchangeVersion -eq $Exchange2007)
{
    $ExchangeShellSnapIn = "Microsoft.Exchange.Management.PowerShell.Admin"	
}
if (@(Get-PSSnapin -Registered|Where-Object {$_.Name -eq $ExchangeShellSnapIn}).Count -eq 1)
{
    if (@(Get-PSSnapin|Where-Object {$_.Name -eq $ExchangeShellSnapIn}).Count -eq 0)
    {
        Add-PSSnapin $ExchangeShellSnapIn
    }
}

if(($ExchangeVersion -eq $Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    Set-AdminAuditLogConfig -AdminAuditLogEnabled $False -WarningAction SilentlyContinue
}

#-------------------------------------------------------------------
# Check whether Exchange server is installed on a domain controller.
#-------------------------------------------------------------------
CheckExchangeInstalledOnDCOrNot
	
#-----------------------------------------------------
# Begin to configure server
#-----------------------------------------------------
Output "Begin to configure server ..." "White"
Output "Steps for manual configuration:" "Yellow" 
Output "Enable remoting in PowerShell." "Yellow"
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

StartService "msexchange*" "auto"
[System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") |Out-File -FilePath $logFile -Append -encoding ASCII -width 100

#-----------------------------------------------------------------------------------------------------------------------
# If the SUT is Exchange Server 2010 or Exchange Server 2013, check the Active Directory Right Management Services role. 
#-----------------------------------------------------------------------------------------------------------------------
if($ExchangeVersion -ge $Exchange2010)
{
    #Create an AD user which will be used as the service account of Active Directory Rights Management Services cluster.
    Output "Create AD user $MSASRMADUser..." "White"
    CreateADUser $MSASRMADUser $userPassword
    Output "Add user $MSASRMADUser into Administrators group." "White"
    AddUserToGroup "$env:USERDOMAIN\$MSASRMADUser" "Administrators"
    #Install ADRMS role automatically if it is not installed. The installation may cause the SUT computer to reboot.
    CheckRoleADRMS  "$env:USERDOMAIN\$MSASRMADUser" $userPassword
}

#----------------------------------------------------------------------------
# Start to create mailbox users
#----------------------------------------------------------------------------
Output "Creating mailbox users on Exchange server $sutComputerName, be patient..." "White"
$mailboxDatabases = Get-MailboxDatabase -Server $sutComputerName 
if(@($mailboxDatabases).count -gt 1)
{
    $mailboxDatabaseName = $mailboxDatabases[0].Name
}
else
{
    $mailboxDatabaseName = $mailboxDatabases.Name
}

CreateMailboxUser $MSASAIRSUser01          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASAIRSUser02          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCALUser01           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCALUser02           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCNTCUser01          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCNTCUser02          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCONUser01           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCONUser02           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCONUser03           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASDOCUser01           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASEMAILUser01         $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASEMAILUser02         $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASEMAILUser03         $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASEMAILUser04         $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASEMAILUser05         $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASHTTPUser01          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASHTTPUser02          $userPassword         $mailboxDatabaseName  $domain
CreateMailboxUser $MSASHTTPUser03          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASHTTPUser04          $userPassword         $mailboxDatabaseName  $domain
CreateMailboxUser $MSASNOTEUser01          $userPassword         $mailboxDatabaseName  $domain
CreateMailboxUser $MSASPROVUser01          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASPROVUser02          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASPROVUser03          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASRMUser01            $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASRMUser02            $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASRMUser03            $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASRMUser04            $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASTASKUser01          $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCMDUser01           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser02           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser03           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCMDUser04           $userPassword         $MailboxDatabaseName  $domain
CreateMailboxUser $MSASCMDUser05           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser06           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser07           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser08           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser09           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser10           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser11           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser12           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser13           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser14           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser15           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser16           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser17           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser18           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDUser19           $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDSearchUser01     $userPassword         $MailboxDatabaseName  $domain 
CreateMailboxUser $MSASCMDSearchUser02     $userPassword         $MailboxDatabaseName  $domain

#-------------------------------------------------------------
# Add delegate for specified mailbox user
#--------------------------------------------------------------
Output "Add delegate of mailbox user $MSASCMDUser07 to mailbox user $MSASCMDUser08." "White"
AddDelegateForMaiboxUser $MSASCMDUser07 $userPassword $MSASCMDUser08 $sutComputerName $domain $ExchangeVersion
Output "Add delegate of mailbox user $MSASEMAILUser04 to mailbox user $MSASEMAILUser05." "White"
AddDelegateForMaiboxUser $MSASEMAILUser04 $userPassword $MSASEMAILUser05 $sutComputerName $domain $ExchangeVersion

Output "Disable the Exchange ActiveSync for mailbox user $MSASCMDUser04."  "White"
$mailboxInfo = Get-CasMailbox -Filter {Name -eq $MSASCMDUser04}
if($mailboxInfo.ActiveSyncEnabled -eq $true)
{
    Set-CasMailbox $MSASCMDUser04 -ActiveSyncEnabled $false
    Output "Disabled Exchange ActiveSync for mailbox user $MSASCMDUser04 successfully." "Green"
}
else
{
    Output "Setting to not enable the Exchange ActiveSync for mailbox user $MSASCMDUser04 has already been done" "Yellow"
}

#----------------------------------------------------------------------------
# Set the properties for mailbox user
#----------------------------------------------------------------------------
Output "Set the properties for mailbox user $MSASCMDUser01." "White"
Import-Module ActiveDirectory
Set-ADUser -Identity $MSASCMDUser01 -SamAccountName $MSASCMDUser01 -GivenName "MSASCMD_FirstName" -Surname "MSASCMD_LastName" -Office "D1042" -Company "MS" -Title  "Manager" -homePhone "22222286" -OfficePhone "55555501" -MobilePhone "8612345678910"
Output "Set the properties for mailbox user $MSASCMDUser01 successfully." "Green"

#----------------------------------------------------------------------------------------------------
# If the SUT is Exchange Server 2010 or Exchange Server 2013, add photo to the specified mailbox user
#----------------------------------------------------------------------------------------------------
if($ExchangeVersion -ge $Exchange2010)
{
    Output "Add photo $MSASCMDUser01Photo to the mailbox user $MSASCMDUser01" "White"
    AddPhotoToMailboxUser $MSASCMDUser01 $MSASCMDUser01Photo
    Output "Add photo $MSASCMDUser02Photo to the mailbox user $MSASCMDUser02" "White"
    AddPhotoToMailboxUser $MSASCMDUser02 $MSASCMDUser02Photo
}

#-------------------------------------------------------------
# Add smtpEmailAddress to the specified mailbox user
#-------------------------------------------------------------- 
Output "Add smtpEmailAddress to the mailbox user $MSASCMDUser01." "White"
$exist = CheckSmtpAddress $MSASCMDUser01 $domain
if($exist -eq $true)
{
    Output "The smtpEmailAddress has already been added for $MSASCMDUser01." "Yellow"
}
else
{
    If($ExchangeVersion -ge $Exchange2010)
    {
        Set-Mailbox $MSASCMDUser01 -EmailAddresses @{add=$MSASCMDUser01+"SMTP@"+$domain} 
    }
    elseif($ExchangeVersion -eq $Exchange2007)
    {	
        $mailboxUserInfo = Get-Mailbox -Identity $MSASCMDUser01
        $mailboxUserInfo.EmailAddresses.Add("smtp:"+$MSASCMDUser01+"SMTP@" + $domain)     
        Set-Mailbox -Instance $mailboxUserInfo 
    }
 
    $check = CheckSmtpAddress $MSASCMDUser01 $domain
    if($check)
    {
        Output "Added smtpEmailAddress to the mailbox user $MSASCMDUser01 successfully." "Green"
    }
    else
    {
         throw "Failed to add smtpEmailAddress to the mailbox user $MSASCMDUser01."
    }
}

#-------------------------------------------------------------
# Configure External URL
#-------------------------------------------------------------
Output "Configure the external URL for the site Microsoft-Server-ActiveSync." "White"
cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/Microsoft-Server-ActiveSync" /commit:APPHOST /section:system.webServer/security/access /sslFlags:"Ssl" | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
$MSASCMDUrl= 'https://' + $sutComputerName + '.'+ $domain +'/Microsoft-Server-ActiveSync'
$MSASCMDSite= $sutComputerName +'\Microsoft-Server-ActiveSync (Default Web Site)'
Set-ActiveSyncVirtualDirectory -ExternalUrl $MSASCMDUrl -Identity $MSASCMDSite 
Output "External URL for the site Microsoft-Server-ActiveSync is configured successfully." "Green"

#-----------------------------------------------------
# New DistributionGroup 
#-----------------------------------------------------
Output "Create two distribution groups $MSASCMDTestGroup and $MSASCMDLargeGroup."  "White"
NewDistributionGroup $MSASCMDTestGroup $domain
NewDistributionGroup $MSASCMDLargeGroup $domain

Output "Create an ActiveSync mailbox policy $MSASPROVUserPolicy01."  "White"
CreateActiveSyncMailboxPolicy $MSASPROVUserPolicy01 $ExchangeVersion
Output "Create an ActiveSync mailbox policy $MSASPROVUserPolicy02."  "White"
CreateActiveSyncMailboxPolicy $MSASPROVUserPolicy02 $ExchangeVersion

if($ExchangeVersion -eq $Exchange2007)
{
    Output "Setting the Exchange search not to index this mailbox database." "White"
    Set-MailboxDatabase -Identity $MailboxDatabaseName -IndexEnabled $false
}

#-----------------------------------------------------
# Create a shared folder 
#-----------------------------------------------------
$sharedFolderPath = & {Split-Path $MyInvocation.scriptName}
$MSASDOCSharedFolderPath = Join-Path $sharedFolderPath $MSASDOCSharedFolder
if(Test-path $MSASDOCSharedFolderPath)
{
    $exist= CheckSharedFolder $MSASDOCSharedFolder
    if($exist -eq $true)
    {
        net.exe share $MSASDOCSharedFolder /delete |Out-Null
    }
    Remove-Item $MSASDOCSharedFolderPath -Recurse -Force
}
Output "Create a shared folder : $MSASDOCSharedFolderPath. Allow Full Control on this shared folder to the user $MSASDOCUser01."  "White"
New-Item $MSASDOCSharedFolderPath -type directory |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
net.exe share $MSASDOCSharedFolder=$MSASDOCSharedFolderPath
GrantUserRightsOnFolder $MSASDOCSharedFolderPath "$domain\$MSASDOCUser01" "FullControl" "Allow"

Output "Create a folder (which is not hidden) $MSASDOCVisibleFolder under the shared folder." "White"
New-Item "$MSASDOCSharedFolderPath\$MSASDOCVisibleFolder" -type directory |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
Output "Create a hidden folder $MSASDOCHiddenFolder under the shared folder." "White"
New-Item "$MSASDOCSharedFolderPath\$MSASDOCHiddenFolder" -ItemType Directory | %{$_.Attributes = "hidden"} |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
Output "Create a document (which is not hidden) $MSASDOCVisibleDocument under the shared folder." "White"
New-Item "$MSASDOCSharedFolderPath\$MSASDOCVisibleDocument" -type file -value "This is a visible text file under a shared folder" |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
Output "Create a hidden document $MSASDOCHiddenDocument under the shared folder." "White"
New-Item "$MSASDOCSharedFolderPath\$MSASDOCHiddenDocument" -type file  -value "This is a hidden text file under a shared folder" | %{$_.Attributes = "hidden"} |Out-File -FilePath $logFile -Append -encoding ASCII -width 100

$MSASCMDSharedFolderPath = Join-Path $sharedFolderPath $MSASCMDSharedFolder 
if(Test-path $MSASCMDSharedFolderPath)
{
    $exist= CheckSharedFolder $MSASCMDSharedFolder
    if($exist -eq $true)
    {
        net.exe share $MSASCMDSharedFolder /delete |Out-Null
    }
    Remove-Item $MSASCMDSharedFolderPath -Recurse -Force
}
Output "Create a shared folder : $MSASCMDSharedFolderPath. Deny the Read permission to the user $MSASCMDUser02 on this shared folder."  "White"
New-Item $MSASCMDSharedFolderPath -type directory |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
net.exe share $MSASCMDSharedFolder=$MSASCMDSharedFolderPath
GrantUserRightsOnFolder $MSASCMDSharedFolderPath "$domain\$MSASCMDUser02" "Read" "Deny"

Output "Create a document $MSASCMDNonEmptyDocument that size should keep at least 4 bytes under the shared folder." "White"
New-Item "$MSASCMDSharedFolderPath\$MSASCMDNonEmptyDocument" -type file -value "The size of this text file is at least 4 bytes under a shared folder" |Out-File -FilePath $logFile -Append -encoding ASCII -width 100
Output "Create an empty document $MSASCMDEmptyDocument under the shared folder."  "White"
New-Item "$MSASCMDSharedFolderPath\$MSASCMDEmptyDocument" -type file |Out-File -FilePath $logFile -Append -encoding ASCII -width 100

#-----------------------------------------------------
# Configure SSL Settings    
#-----------------------------------------------------
Output "Configure SSL settings" "white"
Output "If 'Require SSL' is enabled, server will only accept HTTPS traffic, otherwise server will accept both HTTP and HTTPS traffic." "Yellow"
Output "Disable SSL settings of the following web sites which support both HTTP and HTTPS you want to test on Exchange Server." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1 
Output "$step. In the `"SSL Settings`" page of `"Default Web Site`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow"
$step++
Output "$step. In the `"SSL Settings`" page of `"Default Web Site/Microsoft-Server-ActiveSync`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow" 
$step++
Output "$step. In the `"SSL Settings`" page of `"Default Web Site/Autodiscover`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow" 
ConfigureSSLSettings "Default Web Site" "W3SVC/1/ROOT" "None"
ConfigureSSLSettings "Default Web Site/Microsoft-Server-ActiveSync" "W3SVC/1/ROOT/Microsoft-Server-ActiveSync" "None"
ConfigureSSLSettings "Default Web Site/Autodiscover" "W3SVC/1/ROOT/Autodiscover" "None"

#-----------------------------------------------------
# Add user into the specified group
#-----------------------------------------------------
Output "Add user $MSASCMDUser03 to Administrators group."  "White"
AddUserToGroup "$domain\$MSASCMDUser03" "Administrators"

Output "Add user $MSASHTTPUser04 to Administrators group."  "White"
AddUserToGroup "$domain\$MSASHTTPUser04" "Administrators"

Output "Add user $MSASPROVUser01 to Administrators group."  "White"
AddUserToGroup "$domain\$MSASPROVUser01"  "Administrators"

Output "Add user $MSASRMUser04 to Administrators group."  "White"
AddUserToGroup "$domain\$MSASRMUser04"  "Administrators"


Output "Add user $MSASCMDUser03 to Organization Management group." "White"
AddUserToOrgMgmtGroup $ExchangeVersion $MSASCMDUser03

Output "Add user $MSASPROVUser01 to Organization Management group." "White"
AddUserToOrgMgmtGroup $ExchangeVersion $MSASPROVUser01
	
#-----------------------------------------------------
# Add users to Distribution Group   
#-----------------------------------------------------
Output "Add user $MSASCMDUser01 and $MSASCMDUser02 to distribution group $MSASCMDTestGroup." "White"
AddDistributionGroupMember $MSASCMDTestGroup   $MSASCMDUser01
AddDistributionGroupMember $MSASCMDTestGroup   $MSASCMDUser02

Output "Add 21 mailbox users of MS-ASCMD to distribution group $MSASCMDLargeGroup."  "White"
$MSASCMDUsers = @($MSASCMDUser01,$MSASCMDUser02,$MSASCMDUser03,$MSASCMDUser04,$MSASCMDUser05,$MSASCMDUser06,$MSASCMDUser07,$MSASCMDUser08,$MSASCMDUser09,$MSASCMDUser10,$MSASCMDUser11,$MSASCMDUser12,$MSASCMDUser13,$MSASCMDUser14,$MSASCMDUser15,$MSASCMDUser16,$MSASCMDUser17,$MSASCMDUser18,$MSASCMDUser19,$MSASCMDSearchUser01,$MSASCMDSearchUser02)
foreach($MSASCMDUser in $MSASCMDUsers)
{
    AddDistributionGroupMember $MSASCMDLargeGroup  $MSASCMDUser 
}

#-------------------------------------------------------------
# Add IIS 6 WMI Compatibility role service
#-------------------------------------------------------------- 
if($ExchangeVersion -eq $Exchange2007)
{
    Output "Add IIS 6 WMI Compatibility Role Service." "White"
    Import-Module ServerManager
    Add-WindowsFeature Web-wmi |Out-Null
    Output "Add IIS 6 WMI Compatibility Role Service successfully." "Green"
}

#---------------------------------------------------------------------------------------------------------------------
# If the SUT is Exchange Server 2007, trust the computer for delegation and restart Windows Remote Management Service
#---------------------------------------------------------------------------------------------------------------------
if($ExchangeVersion -eq $Exchange2007)
{
    TrustComputerForDelegation $sutComputerName $domain
    $service = "WinRM"
    $serviceStatus = (Get-Service $service).Status
    if($serviceStatus -ne "Running")
    {
        Start-Service $service
    }
    else
    {
        Restart-Service $service
    }
    $serviceObject = Get-Service $service
    if($serviceObject.status -eq "Running")
    {
        Output "Service $service is started successfully." "Green"
    }
    else
    {
        Throw "Failed to start service $service."
    }
}

#--------------------------------------------------------------------------------
# Set policy with the specified user
#--------------------------------------------------------------------------------
Output "Apply several policy settings for the Mobile Device mailbox policy $MSASPROVUserPolicy02, Please refer to Deployment Guide section 5.1.2 for more information on what settings are applied." "White"
if($ExchangeVersion -le $Exchange2010)
{
    Set-ActiveSyncMailboxPolicy -Identity $MSASPROVUserPolicy02 -AllowNonProvisionableDevices $false -DevicePasswordExpiration 730  -MaxAttachmentSize 2097151 -UnapprovedInROMApplicationList MultiValuedProperty -ApprovedApplicationList d5a090797891fb3ac44749551c87c0a68f46dd82:bthci.dll -confirm:$false
}
elseif($ExchangeVersion -eq $Exchange2013)
{
    Set-MobileDeviceMailboxPolicy -Identity $MSASPROVUserPolicy02 -AllowNonProvisionableDevices $false -PasswordExpiration 730  -MaxAttachmentSize 2097151 -UnapprovedInROMApplicationList MultiValuedProperty -ApprovedApplicationList d5a090797891fb3ac44749551c87c0a68f46dd82:bthci.dll -confirm:$false
}
Output "Mobile device mailbox policy settings for $MSASPROVUserPolicy02 applied successfully." "Green"

Output "Set the mailbox policy $MSASPROVUserPolicy01 to user $MSASPROVUser01" "White"
SetMailboxUserPolicy $MSASPROVUserPolicy01            $MSASPROVUser01
Output "Set the mailbox policy $MSASPROVUserPolicy02 to user $MSASPROVUser02" "White"
SetMailboxUserPolicy $MSASPROVUserPolicy02            $MSASPROVUser02

#----------------------------------------------------------------------------------------
# Move the meeting forward notification email to Delete Items folder.
#----------------------------------------------------------------------------------------
Output "Enable the setting to move meeting forward notification email to the Deleted Items folder for mailbox user $MSASCMDUser01" "White"
MoveNotificationEmailToDeleteFolder $MSASCMDUser01 $ExchangeVersion
Output "Enable the setting to move meeting forward notification email to the Deleted Items folder for mailbox user $MSASHTTPUser03" "White"
MoveNotificationEmailToDeleteFolder $MSASHTTPUser03	$ExchangeVersion
    
$ServerAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -Computer $sutComputerName
if($ServerAdapters -is [array] -and $ExchangeVersion -eq $Exchange2013)
{
    Output "Configure the internal DNS lookups." "White"
    $domainIPAddress= (Get-ADDomainController).IPv4Address
    Set-TransportService $sutComputerName -InternalDNSAdapterEnabled $false -InternalDNSServers $domainIPAddress     
}
	
#----------------------------------------------------------------------------------------
# Configure Active Directory Right Management Services role.
#----------------------------------------------------------------------------------------
if($global:ADRMSInstalledFlag -eq $true)
{
    Output "Configure the following SSL settings only for MS-ASRM" "White"
    $step++
    Output "$step. In the `"SSL Settings`" page of `"Default Web Site/Default Web Site/_wmcs`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow" 
    $step++
    Output "$step. In the `"SSL Settings`" page of `"Default Web Site/Default Web Site/_wmcs/certification`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow" 
    $step++
    Output "$step. In the `"SSL Settings`" page of `"Default Web Site/_wmcs/licensing`" in IIS, clear `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow"
    ConfigureSSLSettings "Default Web Site/_wmcs" "W3SVC/1/ROOT/_wmcs" "None"
    ConfigureSSLSettings "Default Web Site/_wmcs/certification" "W3SVC/1/ROOT/_wmcs/certification" "None"
    ConfigureSSLSettings "Default Web Site/_wmcs/licensing" "W3SVC/1/ROOT/_wmcs/licensing" "None"

    Output "Create a distribution group $MSASRMSuperUserGroup." "White"
    NewDistributionGroup $MSASRMSuperUserGroup $domain

    Output "Add a record rms.$domain in DNS manager" "White"
    AddDNSResourceRecord $sutComputerName
	
    Output "Add the Read and Execute permission for the group on the file license.asmx." "White"
    AddAcl "Exchange Servers" "$Env:SystemDrive\inetpub\wwwroot\_wmcs\licensing\license.asmx"
	
    Output "Add the Read and Execute permission for the group on the file ServerCertification.asmx." "White"
    AddAcl "Users" "$Env:SystemDrive\inetpub\wwwroot\_wmcs\certification\ServerCertification.asmx"
    AddAcl "Exchange Servers" "$Env:SystemDrive\inetpub\wwwroot\_wmcs\certification\ServerCertification.asmx"
    AddAcl "AD RMS Service Group" "$Env:SystemDrive\inetpub\wwwroot\_wmcs\certification\ServerCertification.asmx"
	
    ConfigureRoleADRMS $MSASRMSuperUserGroup $ExchangeVersion
	
    AddDistributedTemplate $MSASRMView_AllowedTemplate                 "Denied all rights except View"              "View,ViewRightsData"
    AddDistributedTemplate $MSASRMView_ReplyAll_AllowedTemplate        "Allowed View and ReplyAll"                  "View,ReplyAll,ViewRightsData"
    AddDistributedTemplate $MSASRMView_Reply_AllowedTemplate	       "Allowed View and reply"                     "View,Reply,ViewRightsData"
    AddDistributedTemplate $MSASRMView_Reply_ReplyAll_AllowedTemplate  "Allowed View,Reply and ReplyAll"            "View,Reply,ReplyAll,ViewRightsData"
    AddDistributedTemplate $MSASRMEdit_Export_NotAllowedTemplate       "Allowed all rigths except Edit and Export"  "View,Print,Forward,Reply,ReplyAll,Extract,AllowMacros,ViewRightsData"
    AddDistributedTemplate $MSASRMExport_NotAllowedTemplate            "Allowed all rights except Export"           "View,Edit,Save,Print,Forward,Reply,ReplyAll,Extract,AllowMacros,ViewRightsData,EditRightsData"
    AddDistributedTemplate $MSASRMReplyAll_NotAllowedTemplate          "Allowed all rights except ReplyAll"         "View,Edit,Save,Export,Print,Forward,Reply,Extract,AllowMacros,ViewRightsData,EditRightsData"
    AddDistributedTemplate $MSASRMAllRights_AllowedTemplate            "Allowed all rights"                         "View,Edit,Save,Export,Print,Forward,Reply,ReplyAll,Extract,AllowMacros,ViewRightsData,EditRightsData"
	 
    if($ExchangeVersion -ge $Exchange2010)
    {
        Output "Enable IRM features for messages sent to internal recipients." "White"
        $IRMInfo= Get-IRMConfiguration
        if($IRMInfo.InternalLicensingEnabled -eq $true)
        {
            Output "Already enabled the licensing for internal messages."   "Yellow"
        }
        else
        {
            Set-IRMConfiguration -InternalLicensingEnabled $true 
        }
        Output "Enable IRM in Microsoft Office Outlook Web App and in Microsoft Exchange ActiveSync." "White"
        if($IRMInfo.ClientAccessServerEnabled -eq $true)
        {
            Output "Already enabled the IRM in Microsoft Office Outlook Web App and in Microsoft Exchange ActiveSync."  "Yellow"
        }
        else
        {
            Set-IRMConfiguration -ClientAccessServerEnabled $true
        }
    }
}
if(($ExchangeVersion -eq $Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    Set-AdminAuditLogConfig -AdminAuditLogEnabled $true -WarningAction SilentlyContinue
}
	
#----------------------------------------------------------------------------------------
# Install Active Directory Certificate Services role.
#---------------------------------------------------------------------------------------- 
InstallADCSRole 

#---------------------------------------------------------------------------------------------------------------------------------------------------
# Create a personal certificate for mailbox user $MSASCMDUser03 then mailbox user $MSASCMDUser03 sends a secure email to mailbox user $MSASCMDUser09
#--------------------------------------------------------------------------------------------------------------------------------------------------- 
Output "Create the personal certificate for mailbox user $MSASCMDUser03 and then mailbox user $MSASCMDUser03 sends a secure email to mailbox user $MSASCMDUser09."
$scriptPath = & {Split-Path $MyInvocation.scriptName}
$domainControllerInfo = Get-ADDomainController | where {($_.Name -eq $sutComputerName) -and ($_.Domain -eq $domain)}
$securePassword = ConvertTo-Securestring $userPassword -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(("$domain\$MSASCMDUser03"),$securePassword)
$session = New-PSSession -ComputerName $sutComputerName -credential $credential
if(($ExchangeVersion -eq $Exchange2007) -or (($domainControllerInfo -ne $null) -and ($domainControllerInfo -ne "")))
{
    Invoke-Command -Session $session -ScriptBlock $createCert -ArgumentList $MSASCMDUser03, $userPassword, $scriptPath, $MSASCMDPfxFileName     
}
else
{
    if(Test-Path "$scriptPath\$MSASCMDPfxFileName")
    {
        Remove-Item "$scriptPath\$MSASCMDPfxFileName"
    }
    #Set up a PowerShell process to create a personal encryption certificate of $MSASCMDUser03.
    $objProcess =New-Object System.Diagnostics.ProcessStartInfo("PowerShell")
    $objProcess.Domain=$domain
    $objProcess.UserName=$MSASCMDUser03
    $objProcess.Password=$securePassword
    $objProcess.UseShellExecute=$false
    $objProcess.WorkingDirectory=$scriptPath
    $filePath=".\CreateCertificate.ps1"
    $argument=$filePath + " " + $MSASCMDUser03 + " " + $userPassword + " " + $MSASCMDPfxFileName
    $objProcess.Arguments=$argument
    [System.Diagnostics.Process]::Start($objProcess)

    #Wait until the certificate is created.
    $retryCount=0
    $time=0
    while($true)
    {
        if(Test-Path "$scriptPath\$MSASCMDPfxFileName")
        {
            Output "The personal encryption certificate of $MSASCMDUser03 was created successfully." "Green"
            break;
        }
        else
        {
            if($retryCount -gt 5)
			{
                throw "Failed to create the personal encryption certificate for $MSASCMDUser03."
			}
			$retryCount=$retryCount + 1
            $time=$time + 10
            Start-Sleep 10
            Output "Elapsed $time seconds to wait for the personal encryption certificate of $MSASCMDUser03." "White" 
        }
    }
}
Invoke-Command -Session $session -ScriptBlock $SendSecureEmail -ArgumentList "$sutComputerName.$domain", "$MSASCMDUser03@$domain", $userPassword,  "$MSASCMDUser09@$domain", $scriptPath, $MSASCMDPfxFileName, $MSASCMDEmailSubjectName, $ExchangeVersion
Output "Mailbox user $MSASCMDUser03 sends the secure email to $MSASCMDUser09 successfully." "Green"

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
Output "Server configuration script was executed successfully." "Green"
Stop-Transcript
exit 0