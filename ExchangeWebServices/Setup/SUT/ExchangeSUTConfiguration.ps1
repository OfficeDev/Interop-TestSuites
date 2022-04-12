$ErrorActionPreference = "Stop"
[String]$containerPath = Get-Location
[String]$logPath       = $containerPath + "\SetupLogs"
[String]$logFile       = $logPath + "\ExchangeSUTConfiguration.ps1.log"
[String]$debugLogFile  = $logPath + "\ExchangeSUTConfiguration.ps1.debug.log"
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
# Default Variables for Configuration 
#----------------------------------------------------------------------------
$sutComputerName                    = $env:ComputerName
$environmentResourceFile            = "$commonScriptDirectory\ExchangeTestSuite.config"

$MSOXWSATTUser01              = ReadConfigFileNode "$environmentResourceFile" "MSOXWSATTUser01"
$MSOXWSATTUser01Password      = ReadConfigFileNode "$environmentResourceFile" "MSOXWSATTUser01Password"

$MSOXWSBTRFUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSBTRFUser01"
$MSOXWSBTRFUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSBTRFUser01Password"

$MSOXWSCONTUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCONTUser01"
$MSOXWSCONTUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCONTUser01Password"

$MSOXWSCOREPublicFolderMailboxPrefix = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREPublicFolderMailboxPrefix"
$MSOXWSCOREPublicFolderMailbox = $MSOXWSCOREPublicFolderMailboxPrefix + $sutComputerName
$MSOXWSCOREPublicFolder       = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREPublicFolder"
$MSOXWSCOREUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser01"
$MSOXWSCOREUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser01Password"
$MSOXWSCOREUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser02"
$MSOXWSCOREUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSCOREUser02Password"

$MSOXWSFOLDUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser01"
$MSOXWSFOLDUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser01Password"
$MSOXWSFOLDUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser02"
$MSOXWSFOLDUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDUser02Password"
$MSOXWSFOLDManagedFolder1     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDManagedFolder1"
$MSOXWSFOLDManagedFolder2     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSFOLDManagedFolder2"

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
$MSOXWSMTGSUser03             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser03"
$MSOXWSMTGSUser03Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSMTGSUser03Password"

$MSOXWSSYNCUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser01"
$MSOXWSSYNCUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser01Password"
$MSOXWSSYNCUser02             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser02"
$MSOXWSSYNCUser02Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSSYNCUser02Password"

$MSOXWSTASKUser01             = ReadConfigFileNode "$environmentResourceFile" "MSOXWSTASKUser01"
$MSOXWSTASKUser01Password     = ReadConfigFileNode "$environmentResourceFile" "MSOXWSTASKUser01Password"

$Exchange2013                 = "Microsoft Exchange Server 2013"
$Exchange2010                 = "Microsoft Exchange Server 2010"
$Exchange2007                 = "Microsoft Exchange Server 2007"
$Exchange2016                 = "Microsoft Exchange Server 2016"

#-----------------------------------------------------------------------------------
# <summary>
# Check if one managedFolder already exists in the server. 
# </summary>
# <param name="$managedFolderName">The name of ManagedFolder.</param>
# <returns>
# Return true if managedFolder exists.
# Return false if managedFolder does not exist.
# </returns>
#-----------------------------------------------------------------------------------
function CheckManagedFolderExistOrNot
{
    param(
    [string]$managedFolderName
    )
    if($ExchangeVersion -le $Exchange2010)
    {
        $folderArray = Get-ManagedFolder
        if(($folderArray.length -ne $null) -and ($folderArray.length -ne ""))
        {
            for($i = 0; $i -lt $folderArray.length; $i++)
            {
                if($folderArray[$i].Name -eq $managedFolderName)
                {
                    return $true
                }
            }
            return $false
        }
       else
       {
           if($folderArray.Name -eq $managedFolderName)
           {
               return $true
           }
           else
           {
               return $false
           }
       }
    }
    else
    {
        $OrganaztionConfig=Get-OrganizationConfig
        $OrganizationDN=$OrganaztionConfig.DistinguishedName
        $identity="CN=$managedFolderName,CN=ELC Folders Container,"+$OrganizationDN
       try
       { 
           $folder =Get-ADObject -Identity $identity          
           return $true
       }
       catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
       {
            return $false
       }
    } 
}
#--------------------------------------------------------------------------------------
# <summary>
# Create a new managedFolder. 
# </summary>
# <param name="$managedFolderName">The name of managedFolder.</param>
#-----------------------------------------------------------------------------------
function CreateManagedFolder
{
    param(
    [string]$managedFolderName
    )

    $exist = CheckManagedFolderExistOrNot $managedFolderName

    if($exist -eq $true)
    {
        OutputWarning "Folder $managedFolderName already exists."
    }
    else
    {
        if($ExchangeVersion -le "Microsoft Exchange Server 2010")
        {
            New-ManagedFolder -Name $managedFolderName -FolderName $managedFolderName | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
			OutputSuccess "Created the managed folder $managedFolderName successfully."
        }
        else
        {
            $OrganaztionConfig=Get-OrganizationConfig
            $OrganizationDN=$OrganaztionConfig.DistinguishedName
            $path="CN=ELC Folders Container,"+$OrganizationDN
            New-ADObject -Name $managedFolderName -Type msExchELCFolder -Path $path -OtherAttributes @{showInAdvancedViewOnly=$TRUE;msExchVersion=4535486012416;msExchELCFolderType=13;msExchELCFolderName=$managedFolderName} | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
            OutputSuccess "Created the managed folder $managedFolderName successfully."
        }
    }    
}

#-----------------------------------------------------------------------------------
# <summary>
# Check if impersonation rights already exists in the server. 
# </summary>
# <param name="impersonationAssigmentName">The assignment name of impersonation.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of user.</param>
# <returns>
# Return true if impersonation rights is granted.
# Return false if impersonation rights is not granted.
# Return rename if impersonationAssignmentName is already used by other user.
# </returns>
#-----------------------------------------------------------------------------------
function CheckImpersonationRightsExistOrNot
{
    param(
    [string]$impersonationAssignmentName,
    [string]$ExchangeVersion,
    [string]$userName
    )
    ValidateParameter 'impersonationAssignmentName' $impersonationAssignmentName
    ValidateParameter 'ExchangeVersion' $ExchangeVersion
    ValidateParameter 'userName' $userName
    
    if($ExchangeVersion -eq $Exchange2007)
    {
        $ExchangeServerNameArray=Get-ExchangeServer
        if($ExchangeServerNameArray -is [Array])
        {
            for($i = 0; $i -lt $ExchangeServerNameArray.length; $i++)
            {
                if($ExchangeServerNameArray[$i].Name -eq $Env:ComputerName)
                {
                    $global:ExchangeServerName = (Get-ExchangeServer)[$i].distinguishedname
                    $global:mailboxDatabaseName = (Get-MailboxDatabase)[$i].distinguishedname            
                    break
                }
            }
        }
        else
        {
            $global:ExchangeServerName = (Get-ExchangeServer).distinguishedname
            $global:mailboxDatabaseName =(Get-MailboxDatabase).distinguishedname    
        }
        
        $domain= $Env:UserDNSDomain.split(".")[0]
        $impersonation= Get-ADPermission -Identity $global:ExchangeServerName | where {($_.ExtendedRights -like "ms-Exch-EPI-Impersonation") -and ($_.User -like "$domain\$userName")}
        $mayImpersonation= Get-ADPermission -Identity $global:mailboxDatabaseName | where {($_.ExtendedRights -like "ms-Exch-EPI-May-Impersonate") -and ($_.User -like "$domain\$userName")}
        if(($impersonation -ne $null -and $impersonation -ne "") -and ($mayImpersonation -ne $null -and $mayImpersonation -ne ""))
        {
            return $true
        }
        else
        {
            return $false
        }
    }
    elseif($ExchangeVersion -ge $Exchange2010)
    {
        $assignments = Get-ManagementRoleAssignment -Role ApplicationImpersonation
        foreach($assignment in $assignments)
        {
            if($assignment.Name -eq $impersonationAssignmentName)
            {
                if($assignment.RoleAssigneeName -eq $userName)
                {
                    return $true
                }
                else
                {
                    return "rename"
                }
            }
                
        }
    }

    return $false
}

#--------------------------------------------------------------------------------------
# <summary>
# Grant impersonation rights for specified user. 
# </summary>
# <param name="impersonationAssigmentName">The assignment name of impersonation.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of user.</param>
#-----------------------------------------------------------------------------------
function GrantImpersonationRights
{
    param(
    [string]$impersonationAssignmentName,
    [string]$ExchangeVersion,
    [string]$userName
    )
    ValidateParameter 'impersonationAssignmentName' $impersonationAssignmentName
    ValidateParameter 'ExchangeVersion' $ExchangeVersion
    ValidateParameter 'userName' $userName
    
    $exist = CheckImpersonationRightsExistOrNot $impersonationAssignmentName $ExchangeVersion $userName 
    if($exist -eq $true)
    {
        OutputWarning "Impersonation rights for the user $userName is already granted."
    }
    else
    {
        if($ExchangeVersion -eq $Exchange2007)
        {
            
            Add-ADPermission -Identity $global:ExchangeServerName -user $userName -extendedRight ms-Exch-EPI-Impersonation | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
            Add-ADPermission -Identity $global:mailboxDatabaseName -User $userName -ExtendedRights ms-Exch-EPI-May-Impersonate | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        elseif($ExchangeVersion -ge $Exchange2010)
        {
            if($exist -eq "rename")
            {
                $impersonationAssignmentName = [System.Guid]::NewGuid().toString()        
            }
            
            New-ManagementRoleAssignment -Name:$impersonationAssignmentName -Role:ApplicationImpersonation -User:"$userName@$env:userdnsdomain" | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }

        $check = CheckImpersonationRightsExistOrNot $impersonationAssignmentName $ExchangeVersion $userName 
        if($check)
        {
            OutputSuccess "Granted the impersonation rights for the user $userName successfully."
        }
        else
        {
            Throw("Failed to granted impersonation rights for the user $userName.")
        }
    }
}

#-----------------------------------------------------
# Begin to configure Exchange server
#-----------------------------------------------------
OutputText "Begin to configure Exchange server ..."
OutputWarning "Steps for manual configuration:"
OutputWarning "Enable remoting in Powershell."
Invoke-Command {
    $ErrorActionPreference = "Continue"
    Enable-PSRemoting -Force
}

#-----------------------------------------------------
# Get Exchange server basic information
#-----------------------------------------------------
OutputText "The basic information of the Exchange server:"

$domain          = $Env:UserDNSDomain
OutputText "Domain name: $domain"
$sutComputerName = $Env:ComputerName
OutputText "The name of the Exchange server: $sutComputerName"
$userName        = $Env:UserName
OutputText "The logon name of the current user: $userName "
$room            = "ResourceMailbox"

$ExchangeVersion = GetExchangeServerVersion


#-----------------------------------------------------
# Add Exchange PowerShell snapin
#-----------------------------------------------------
AddExchangePSSnapIn
#-----------------------------------------------------
# Check whether Exchange server is installed on a domain controller.
#-----------------------------------------------------
InstallWindowsFeature RSAT-AD-PowerShell
CheckExchangeInstalledOnDCOrNot

#-----------------------------------------------------
# Start Exchange automatic started services.
#-----------------------------------------------------
StartService "msexchange*" "auto"

$adminAuditLogEnabledChanged = $false

if(($ExchangeVersion -eq $Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    $adminAuditLogConfig=Get-AdminAuditLogConfig
    if($adminAuditLogConfig.AdminAuditLogEnabled -eq $true)
    {
        Set-AdminAuditLogConfig -AdminAuditLogEnabled $False -WarningAction SilentlyContinue | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        $adminAuditLogEnabledChanged = $true 
    } 
}

#-----------------------------------------------------
# Start to create mailbox users
#-----------------------------------------------------
OutputText "Mailbox users are currently being created on the Exchange server; please wait..."
[System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
$mailboxDatabases = Get-MailboxDatabase -Server $sutComputerName
if(@($mailboxDatabases).count -gt 1)
{
    $mailboxDatabaseName = $mailboxDatabases[0].Name
}
else
{
    $mailboxDatabaseName = $mailboxDatabases.Name
}

CreateMailboxUser  $MSOXWSATTUser01   $MSOXWSATTUser01Password        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSBTRFUser01  $MSOXWSBTRFUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSCONTUser01  $MSOXWSCONTUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSCOREUser01  $MSOXWSCOREUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSCOREUser02  $MSOXWSCOREUser02Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSFOLDUser01  $MSOXWSFOLDUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSFOLDUser02  $MSOXWSFOLDUser02Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMSGUser01   $MSOXWSMSGUser01Password        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMSGUser02   $MSOXWSMSGUser02Password        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMSGUser03   $MSOXWSMSGUser03Password        $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMSGRoom01   $MSOXWSMSGRoom01Password        $mailboxDatabaseName $domain $room
CreateMailboxUser  $MSOXWSMTGSUser01  $MSOXWSMTGSUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMTGSUser02  $MSOXWSMTGSUser02Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMTGSRoom01  $MSOXWSMTGSRoom01Password       $mailboxDatabaseName $domain $room
CreateMailboxUser  $MSOXWSSYNCUser01  $MSOXWSSYNCUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSSYNCUser02  $MSOXWSSYNCUser02Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSTASKUser01  $MSOXWSTASKUser01Password       $mailboxDatabaseName $domain
CreateMailboxUser  $MSOXWSMTGSUser03  $MSOXWSMTGSUser03Password       $mailboxDatabaseName $domain

#-------------------------------------------------------------
# Add delegate for specified mailbox user
#--------------------------------------------------------------
OutputText "Add delegate of mailbox user $MSOXWSMTGSUser03 to mailbox user $MSOXWSMTGSUser02."
AddDelegateForMaiboxUser $MSOXWSMTGSUser03 $MSOXWSMTGSUser03Password $MSOXWSMTGSUser02 $sutComputerName $domain $ExchangeVersion

if($ExchangeVersion -le $Exchange2010)
{
    CreatePublicFolderDatabase "PublicFolderDatabase" "$sutComputerName"
}
elseif($ExchangeVersion -ge $Exchange2013)
{
    $publicFolderMailboxInfo = Get-Mailbox -PublicFolder -filter {Name -eq $MSOXWSCOREPublicFolderMailbox}
    if(($publicFolderMailboxInfo -ne $null) -and ($publicFolderMailboxInfo -ne ""))
    {    
        OutputWarning "Public Folder Mailbox $MSOXWSCOREPublicFolderMailbox already exists!"
    }
    else
    {
        New-Mailbox -PublicFolder -Name $MSOXWSCOREPublicFolderMailbox | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        OutputSuccess "Public Folder Mailbox $MSOXWSCOREPublicFolderMailbox created!"
    }
}

if($ExchangeVersion -le $Exchange2010)
{
    $publicFolders=Get-PublicFolder -Server $sutComputerName -Recurse
}
elseif($ExchangeVersion -ge $Exchange2013)
{
    $publicFolders=Get-PublicFolder -Recurse
}

if($publicFolders -is [array])
{
    $i=$publicFolders.length-1
    while($i -ge 0)
    {
        if($publicFolders[$i].Name -eq $MSOXWSCOREPublicFolder)
        {
            OutputWarning "Public Folder already exists! Delete and re-create it"
            Remove-PublicFolder -Identity "\$MSOXWSCOREPublicFolder" -Recurse -Confirm:$false
            break
        }
        $i--
    }
}
elseif($publicFolders.Name -eq $MSOXWSCOREPublicFolder)
{
    OutputWarning "Public Folder already exists! Delete and re-create it"
    Remove-PublicFolder -Identity "\$MSOXWSCOREPublicFolder" -Recurse
}

OutputText "Creating a new public folder ..."
if($ExchangeVersion -le $Exchange2010)
{
    New-PublicFolder -Name $MSOXWSCOREPublicFolder -Server $sutComputerName | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
}
else
{
    New-PublicFolder -Name $MSOXWSCOREPublicFolder -Mailbox $MSOXWSCOREPublicFolderMailbox | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
}
OutputSuccess "New Public Folder MSOXWSCORE_PublicFolder created!"

OutputText "check the public folder's client permission..."
if($ExchangeVersion -le $Exchange2010)
{
    $permissionUser = Get-PublicFolderClientPermission -Identity "\$MSOXWSCOREPublicFolder" -User $MSOXWSCOREUser01 
}
else
{
    $permissionUser = Get-PublicFolderClientPermission -Identity "\$MSOXWSCOREPublicFolder" | where {$_.User.DisplayName -eq $MSOXWSCOREUser01}    
}
if($permissionUser -ne $null)
{
    OutputWarning "$MSOXWSCOREPublicFolder's client permission for $MSOXWSCOREUser01 already exists! Delete and re-create it"
    Remove-PublicFolderClientPermission -Identity "\$MSOXWSCOREPublicFolder" -User $MSOXWSCOREUser01 -Confirm:$false
}

Add-PUblicFolderClientPermission -Identity "\$MSOXWSCOREPublicFolder" -User $MSOXWSCOREUser01 -AccessRights Owner | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
OutputSuccess "Added the user $MSOXWSCOREUser01 Owner permission to $MSOXWSCOREPublicFolder successfully."

OutputText "Creating new managed folders ..."
CreateManagedFolder $MSOXWSFOLDManagedFolder1
CreateManagedFolder $MSOXWSFOLDManagedFolder2

OutputText "Granting impersonation permissions for the specified user..."
GrantImpersonationRights "MSOXWSATTExchangeImpersonation"  $ExchangeVersion $MSOXWSATTUser01
GrantImpersonationRights "MSOXWSBTRFExchangeImpersonation" $ExchangeVersion $MSOXWSBTRFUser01
GrantImpersonationRights "MSOXWSCOREExchangeImpersonation" $ExchangeVersion $MSOXWSCOREUser01
GrantImpersonationRights "MSOXWSFOLDExchangeImpersonation" $ExchangeVersion $MSOXWSFOLDUser01
GrantImpersonationRights "MSOXWSSYNCExchangeImpersonation" $ExchangeVersion $MSOXWSSYNCUser01

if($ExchangeVersion -eq $Exchange2007)
{
    $pfAdminGroup = "PublicFolderAdmin"  
}
elseif($ExchangeVersion -ge $Exchange2010)
{
    $pfAdminGroup = "Public Folder Management"
}
AddUserToExchangeAdminGroup $ExchangeVersion $MSOXWSFOLDUser01 $pfAdminGroup
if ($ExchangeVersion -ge $Exchange2016)
{
    $cONTGroup = "Recipient Management"
AddUserToExchangeAdminGroup $ExchangeVersion $MSOXWSCONTUser01 $cONTGroup
}
OutputText "Start Microsoft Exchange Transport service ..."
StartService "MSExchangeTransport"

if($ExchangeVersion -ge $Exchange2013)
{
    OutputText "Starting the Microsoft Exchange Mailbox Transport Delivery service..."
    StartService "MSExchangeDelivery"
    
    OutputText "Starting the Microsoft Exchange Search Host Controller service..."
    StartService "HostControllerService"
}

if($ExchangeVersion -le $Exchange2010)
{
    OutputText "Starting the Microsoft Exchange Mail Submission service..."
    StartService "MSExchangeMailSubmission"
}

OutputText "Configuring Exchange web services without SSL ..."
cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/EWS" /commit:APPHOST /section:system.webServer/security/access /sslFlags:"None"

if(($global:ExchangeVersion -eq $global:Exchange2010) -and ($PSVersionTable.PSVersion.Major -ge 3))
{
    if($adminAuditLogEnabledChanged -eq $true)
    {
        Set-AdminAuditLogConfig -AdminAuditLogEnabled $true -WarningAction SilentlyContinue | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    }
}

IISReset /restart

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
OutputSuccess "[ExchangeSUTConfiguration.PS1] has run successfully."
Stop-Transcript
exit 0