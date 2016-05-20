#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

#-----------------------------------------------------------------------------------
# <summary>
# Return new valid ports for RPC over HTTP, which will be used for the Configuration of RPC over HTTP.
# </summary>
# <returns>
# The new valid ports format is "SUTName:593;SUTName:50000-60535;SUTName:6001-6002;SUTName:6004;SUTIPv4s:6001-6002;SUTIPv4s:6004". 
# </returns>
#-----------------------------------------------------------------------------------
function GetNewRpchValidPorts
{
    $validPortsNumberforAdd1 = "6001-6002"
    $validPortsNumberforAdd2 = "6004"
    $validPortsAdd1 = "$env:COMPUTERNAME`:$validPortsNumberforAdd1;$env:COMPUTERNAME`:$validPortsNumberforAdd2"
    $validPortsAdd2 = ""
    $ipConfigSet = Get-WmiObject Win32_NetworkAdapterConfiguration

    foreach ($ipConfig in $ipConfigSet)
    {
        if ($ipConfig.IPEnabled)
        {
            $ip = $ipConfig.IPAddress
            $ipv4 = $ip[0]
            $validPortsAdd = "$ipv4`:$validPortsNumberforAdd1;$ipv4`:$validPortsNumberforAdd2`;" 
            if($validPortsAdd2 -eq "")
            {
                $validPortsAdd2 = $validPortsAdd
            }
            else
            {                
                $validPortsAdd2 = $validPortsAdd2 + ";" + $validPortsAdd
            }        
        }
    }
    $validPortsForEx = "$env:COMPUTERNAME`:593;$env:COMPUTERNAME`:50000-60535"
    $validPortsNew = $validPortsForEx + ";" + $validPortsAdd1 + ";" + $validPortsAdd2
    return $validPortsNew
}

#-----------------------------------------------------------------------------------
# <summary>
# Compare the recommended Exchange minor version with the installed Exchange minor version.
# </summary>
# <param name="actualVersion">The display version of the Exchange installed currently.</param>
# <param name="recommendedVersion">An array with three elements, the recommended Exchange display version.</param> 
# <returns>
# A Boolean value, true if the server installed the recommended service pack, otherwise false.
# </returns>
#-----------------------------------------------------------------------------------          
function CompareExchangeMinorVersion
{
    param(
    [String]$actualVersion,
    [String]$recommendedVersion
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'actualVersion' $actualVersion
    ValidateParameter 'recommendedVersion' $recommendedVersion
    
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
    $ExchangeServer2007             = "$global:Exchange2007",   "8.3.83.6",      "SP3"
    $ExchangeServer2010             = "$global:Exchange2010",   "14.3.123.4",    "SP3"
    $ExchangeServer2013             = "$global:Exchange2013",   "15.0.847.32",   "SP1"
    $ExchangeServer2016             = "$global:Exchange2016",   "15.1.280.0",   "0"
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
        if($item.DisplayName.StartsWith($ExchangeServer2016[0]))
        {
            $version = $item.DisplayVersion
            $ExchangeVersion = $ExchangeServer2016[0]
            $recommendVersion = $ExchangeServer2016[1]
            $recommendMinorVersion = $ExchangeServer2016[2]
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
# Add Exchange PowerShell SnapIn. 
# </summary>
#-----------------------------------------------------------------------------------
function AddExchangePSSnapIn
{
    if($global:ExchangeVersion -ge $global:Exchange2010)
    {
        $ExchangeShellSnapIn = "Microsoft.Exchange.Management.PowerShell.E2010"    
    }
    if($global:ExchangeVersion -eq $global:Exchange2007)
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
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether Exchange server is installed on a domain controller.
# </summary>
#-----------------------------------------------------
function CheckExchangeInstalledOnDCOrNot
{
    Import-module ActiveDirectory
    $domainControllerInfo = Get-ADDomainController | where {($_.Name -eq $env:COMPUTERNAME ) -and ($_.Domain -eq $env:USERDNSDOMAIN)}
    if(($domainControllerInfo -ne $null) -and ($domainControllerInfo -ne ""))
    {
        Output "An Exchange server installed on a domain controller is not recommended." "Yellow"
        Output "Would you like to continue running the SUT setup script on this machine or exit?" "Cyan"    
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
# Check if one Mailbox User already exists in the server. 
# </summary>
# <param name="mailboxUserName">The username of mailbox.</param>
# <returns>
# Return true if mailbox exists.
# Return false if mailbox does not exist.
# </returns>
#-----------------------------------------------------------------------------------
function CheckMailboxUserExistOrNot
{
    param(
    [string]$mailboxUserName
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'mailboxUserName' $mailboxUserName
    
    $mailboxArray = Get-Mailbox
    for($i = 0; $i -lt $mailboxArray.length; $i++)
    {
        if($mailboxArray[$i].Name -eq $mailboxUserName)
        {
            return $true
        }
    }
    return $false
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a new mailbox user. 
# </summary>
# <param name="mailboxUserName">The username of mailbox.</param>
# <param name="mailboxUserPassword">The password of mailbox user.</param>
# <param name="mailboxUserDatabase">The database of mailbox user.</param>
# <param name="mailboxUserDomain">The domain that mailbox user belongs to.</param>
#-----------------------------------------------------------------------------------
function CreateMailboxUser
{
    param(
    [string]$mailboxUserName,
    [string]$mailboxUserPassword,
    [string]$mailboxUserDatabase,
    [string]$mailboxUserDomain
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'mailboxUserName' $mailboxUserName
    ValidateParameter 'mailboxUserPassword' $mailboxUserPassword
    ValidateParameter 'mailboxUserDatabase' $mailboxUserDatabase
    ValidateParameter 'mailboxUserDomain' $mailboxUserDomain
    
    Output "Create a mailbox user $mailboxUserName." "White"
    $exist = CheckMailboxUserExistOrNot $mailboxUserName
    if($exist -eq $true)
    {
        Output "Mailbox for $mailboxUserName already exists!" "Yellow"        
    }
    else
    {
        if($mailboxUserName.Length -ge 21)
        {
            Output "The mailbox username should be below 20 characters." "Yellow"
            Output "Warning: The mailbox user name $mailboxUserName exceeds 20 characters. This may cause an RPC connection failure." "Yellow"
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
        New-Mailbox -UserPrincipalName "$mailboxUserName@$mailboxUserDomain" -Database $mailboxUserDatabase -Name $mailboxUserName -Password $securePassword | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        $check = CheckMailboxUserExistOrNot $mailboxUserName
        if($check)
        {
            Output "Mailbox for $mailboxUserName created successfully." "Green"
        }
        else
        {
            Throw("Failed to create the mailbox for $mailboxUserName.")
        }        
    }
    SetPasswordNeverExpires $mailboxUserName
}

#--------------------------------------------------------------------------------------
# <summary>
# Create a public folder database if there is no public folder database on Exchange server 2007 or Exchange server 2010. 
# </summary>
# <param name="publicFolderDatabaseName">The name of public folder database.</param>
# <param name="server">The Exchange server name where public folder is located.</param>
# <returns>
# The name of existed public folder database.
# </returns>
#-----------------------------------------------------------------------------------
function CreatePublicFolderDatabase
{
    param(
    [string]$publicFolderDatabaseName,
    [string]$server
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderDatabaseName' $publicFolderDatabaseName
    ValidateParameter 'server' $server
   
    if($global:ExchangeVersion -le $global:Exchange2010)
    {
        $publicFolderDatabase = Get-Publicfolderdatabase -Server $server
        if (($publicFolderDatabase -eq $null) -or ($publicFolderDatabase -eq ""))
        {
            Output "A public folder database is currently being created on the Exchange server; please wait..." "White"    
            if($global:ExchangeVersion -eq $global:Exchange2007)
            {
                $storageGroup = Get-StorageGroup -Server $server
                if ($storageGroup -is [array])
                {
                    $storageGroupName = $storageGroup[0].Identity.ToString()
                }
                else
                {
                    $storageGroupName = $storageGroup.Identity.ToString()
                }
                $publicFolderDatabaseName = New-PublicFolderDatabase -Name $publicFolderDatabaseName -StorageGroup "$storageGroupName"
            }
            if($global:ExchangeVersion -eq $global:Exchange2010)
            {
                $publicFolderDatabaseName = New-PublicFolderDatabase -Name $publicFolderDatabaseName -Server $server
            }
            Output "The public folder database $publicFolderDatabaseName was created successfully." "Green"       
        }
        else
        {
            Output "A public folder database already exists." "Yellow"
            $publicFolderDatabaseName = $publicFolderDatabase
        }

        Output "Mounting the public folder database $publicFolderDatabaseName." "White"
        Mount-Database -Identity $publicFolderDatabaseName
        Output "Mounted the public folder database $publicFolderDatabaseName successfully." "Green"
        return $publicFolderDatabaseName
    }
}

#--------------------------------------------------------------------------------------
# <summary>
# Create a public folder mailbox if there is no public folder mailbox on Exchange server 2013. 
# </summary>
# <param name="publicFolderMailboxName">The name of public folder mailbox.</param>
# <param name="server">The Exchange server name where public folder is located.</param>
# <param name="mailboxDatabaseName">The name of mailbox database.</param>
# <returns>
# The name of existed public folder mailbox.
# </returns>
#-----------------------------------------------------------------------------------
function CreatePublicFolderMailbox
{
    param(
    [string]$publicFolderMailboxName,
    [string]$server,
    [string]$mailboxDatabaseName
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderMailboxName' $publicFolderMailboxName
    ValidateParameter 'server' $server
    ValidateParameter 'mailboxDatabaseName' $mailboxDatabaseName

    if($global:ExchangeVersion -ge $global:Exchange2013)
    {
        $publicFolderMailbox = Get-Mailbox -PublicFolder -Server $server
        if($publicFolderMailbox.Count -le 0)
        {
            Output "A public folder is currently being created on the Exchange server; please wait..." "White"
            $publicFolderMailboxName = New-Mailbox -PublicFolder -Name $publicFolderMailboxName -Database $mailboxDatabaseName
            Output "Created the public folder mailbox $publicFolderMailboxName successfully." "Green"
        }
        else
        {
            Output "A public folder already exists." "Yellow" 
            if($publicFolderMailbox -is [array])
            {
               $publicFolderMailboxName = $publicFolderMailbox[0].Name
            }  
            else 
            {
               $publicFolderMailboxName = $publicFolderMailbox.Name 
            }            
        }
        return $publicFolderMailboxName
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Enable mail on Public Folder
# </summary>
# <param name="publicFolderName">The name of public folder.</param>
#-----------------------------------------------------------------------------------
function EnableMailOnPublicFolder
{
    param(
    [string]$publicFolderName
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderName' $publicFolderName
    
    if(!(Get-PublicFolder -Identity $publicFolderName).MailEnabled)
    {
        Output "Set the public folder $publicFolderName to mail-enabled." "White"
        Enable-MailPublicFolder -Identity $publicFolderName
        Output "Set the public folder $publicFolderName to mail-enabled successfully." "Green"
    }
    else
    {
        Output "Mail is already enabled on the public folder $publicFolderName." "Yellow"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check if one public folder already exists in the server
# </summary>
# <param name="publicFolderName">The name of public folder.</param>
# <param name="server">The Exchange server name where public folder is located, used in Exchange 2007 and Exchange 2010.</param>
# <returns>
# Return true if exist.
# Return false if not exist.
# </returns>
#-----------------------------------------------------------------------------------
function CheckPublicFolderExistOrNot
{
    param(
    [string]$publicFolderName,
    [string]$server
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderName' $publicFolderName
    if($global:ExchangeVersion -le $global:Exchange2010)
    {
        ValidateParameter 'server' $server
    }

    if($global:ExchangeVersion -le $global:Exchange2010)
    {
        $publicFolders = Get-PublicFolder -Server $server -Recurse
    }
    elseif($global:ExchangeVersion -ge $global:Exchange2013)
    {
         $publicFolders = Get-PublicFolder -Recurse
    }
    if($publicFolders -is [array])
    {
        $i = $publicFolders.length - 1
        while($i -ge 0)
        {
            if($publicFolders[$i].Name -eq $publicFolderName)
            {
                return $true
                break
            }
            $i--
        }
    }
    else
    {
        if($publicFolders.Name -eq $publicFolderName)
        {
            return $true
        }
    }
    return $false
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a public folder. 
# </summary>
# <param name="publicFolderName">The name of public folder.</param>
# <param name="server">The name of Exchange server where public folder is located, used in Exchange 2007 and Exchange 2010.</param>
# <param name="publicFolderMailboxName">The name of public folder mailbox where public folder is located, optional, used in Exchange 2013.</param>
#-----------------------------------------------------------------------------------
function CreatePublicFolder
{
    param(
    [string]$publicFolderName,
    [string]$server,
    [string]$publicFolderMailboxName
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderName' $publicFolderName
    ValidateParameter 'server' $server
    if($global:ExchangeVersion -ge $global:Exchange2013)
    {
        ValidateParameter 'publicFolderMailboxName' $publicFolderMailboxName
    }

    $exist = CheckPublicFolderExistOrNot $publicFolderName $server
    if($exist -eq $true)
    {
        Output "$publicFolderName already exists!" "Yellow"        
    }
    else
    {
        if($global:ExchangeVersion -le $global:Exchange2010)
        {
            New-PublicFolder -Name $publicFolderName -Server $server | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        elseif($global:ExchangeVersion -ge $global:Exchange2013)
        {
            New-PublicFolder -Name $publicFolderName -Mailbox $publicFolderMailboxName | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }

        $check = CheckPublicFolderExistOrNot $publicFolderName $server
        if($check)
        {
            Output "$publicFolderName created successfully." "Green"
        }
        else
        {
            Throw("Failed to create $publicFolderName.")
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check ghosted public folder status. 
# </summary>
# <param name="publicFolderName">The name of public folder.</param>
# <param name="server">The name of Exchange server where ghosted public folder is located, used in Exchange 2007 and Exchange 2010.</param>
#-----------------------------------------------------------------------------------
function CheckGhostedPublicFolderStatus
{
    param(
    [string]$publicFolderName,
    [string]$server
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'publicFolderName' $publicFolderName
    ValidateParameter 'server' $server

    $count = 0
    $time = 0
    while($true)
    {
        if($count -gt 20)
        {        
            Output "The command Get-PublicFolder has timed out even after waiting for 100 minutes. To resolve this issue, increase the performance of your local machine such as changing the settings of the CPU or Memory, and then restart the script." "Red"
            Stop-Transcript
            exit 2
        }

        Get-PublicFolder -Server $server "\$publicFolderName" -ErrorAction silentlyContinue | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        if(!$?)
        {
            #MapiObjectNotFoundException is the specific exception of Get-PublicFolder when the time is not up.
            if($error[0].CategoryInfo.Reason -eq "MapiObjectNotFoundException")
            {
                $count = $count + 1
                $time = $count * 5
                Start-Sleep -s 300
                Output "Elapsed time is $time minutes : Waiting for $publicFolderName on $server to become a ghosted public folder." "yellow"
            }
            else
            {
                throw $error[0]
            }
        }                    
        else
        {
            Output "$publicFolderName on $server has become a ghosted public folder after $time minutes." "Green"
            break
        }
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Check if the user is already in the specified Exchange Admin group.
# </summary>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of user.</param>
# <param name="groupName">The name of specified Exchange Admin group.</param>
# <returns>
# Return true if user is already in the specified Exchange Admin group.
# Return false if user is not in the specified Exchange Admin group.
# </returns>
#----------------------------------------------------------------------------------
Function CheckExchangeAdminMember
{
    param(
    [string]$ExchangeVersion,
    [string]$userName,
    [string]$groupName
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'ExchangeVersion' $ExchangeVersion
    ValidateParameter 'userName' $userName
    ValidateParameter 'groupName' $groupName
    
    if($ExchangeVersion -eq $global:Exchange2007)
    {
        $adminRole = Get-ExchangeAdministrator -Identity $userName | where {$_.Role -eq $groupName}  
        if($adminRole -ne $null)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
    elseif($ExchangeVersion -ge $global:Exchange2010)
    {
        $adminMember = Get-RoleGroupMember -Identity $groupName | where {$_.Name -eq $userName}         
        if($adminMember -ne $null)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Add user to specified Exchange Admin group.
# </summary>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
# <param name="userName">The name of user.</param>
# <param name="groupName">The name of specified Exchange Admin group.</param>
#-----------------------------------------------------------------------------------
Function AddUserToExchangeAdminGroup
{
    param(
    [string]$ExchangeVersion,
    [string]$userName,
    [string]$groupName
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'ExchangeVersion' $ExchangeVersion
    ValidateParameter 'userName' $userName
    ValidateParameter 'groupName' $groupName
 
    $exist = CheckExchangeAdminMember $ExchangeVersion $userName $groupName
    
    if($exist)
    {
        Output "The user $userName is already in the $groupName group." "Yellow"
    }
    else
    {
        if($ExchangeVersion -eq $global:Exchange2007)
        {            
            Add-ExchangeAdministrator -Role $groupName -Identity $userName | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }
        
        elseif($ExchangeVersion -ge $global:Exchange2010)
        {
            Add-RoleGroupMember -Identity $groupName -member $userName -BypassSecurityGroupManagerCheck | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        }

        $check = CheckExchangeAdminMember $ExchangeVersion $userName $groupName
        if($check)
        {
            Output "Added user $userName to $groupName group successfully." "Green"
        }
        else
        {
            Throw("Failed to add user $useName to $groupName group!")
        }
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Add user specific public folder client permission.
# </summary> 
# <param name="userName">The name of user.</param>
# <param name="publicFolderName">The name of public folder.</param>
# <param name="accessRights"> The rights being removed.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#-----------------------------------------------------------------------------------
Function AddUserPublicFolderClientPermission
{
    param(
    [string]$userName,
    [string]$publicFolderName,
    [string]$accessRights,
    [string]$ExchangeVersion
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'userName' $userName
    ValidateParameter 'publicFolderName' $publicFolderName
    ValidateParameter 'accessRights' $accessRights
    ValidateParameter 'ExchangeVersion' $ExchangeVersion
    
    if($ExchangeVersion -le $global:Exchange2010)
    {        
        $permissionUser = Get-PublicFolderClientPermission -Identity $publicFolderName -User $userName
        if($permissionUser -ne $null)
        {
            Remove-PublicFolderClientPermission -Identity $publicFolderName -User $userName -AccessRights $permissionUser.AccessRights -Confirm:$false
        }
    }
    else
    {
        $permissionUser = Get-PublicFolderClientPermission -Identity $publicFolderName | where {$_.User.DisplayName -eq $userName}
        if($permissionUser -ne $null)
        {
            Remove-PublicFolderClientPermission -Identity $publicFolderName -User $userName -Confirm:$false
        }
    }
    Add-PublicFolderClientPermission -Identity $publicFolderName -User $userName -AccessRights $accessRights | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    if($publicFolderName -eq "\")
    {
        Output "Added the user $userName $accessRights permission to the root public folder successfully." "Green"
    } 
    else
    {
        Output "Added the user $userName $accessRights permission to the public folder $publicFolderName successfully." "Green" 
    }
}

#----------------------------------------------------------------------------------
# <summary>
# Create a Distribution Group.
# </summary>
# <param name="distributionGroupName">The name of the Distribution Group that will be created.</param>
# <param name="distributionGroupType">The group type of the Distribution Group that will be created.</param>
# <param name="samAccountName">The name for clients of the object running older operating systems.</param>
#-----------------------------------------------------------------------------------
Function CreateDistrbutionGroup
{
    param(
    [string]$distributionGroupName,
    [string]$distributionGroupType,
    [string]$samAccountName
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'distributionGroupName' $distributionGroupName
    ValidateParameter 'distributionGroupType' $distributionGroupType
    ValidateParameter 'samAccountName' $samAccountName
    
    $distributionGroupArray = Get-DistributionGroup -Filter {Name -eq $distributionGroupName}
    if($distributionGroupArray -eq $null)
    {
        Output "Creating a distribution group $distributionGroupName." "White"
        New-DistributionGroup -Name $distributionGroupName -SAMAccountName $samAccountName -Type $distributionGroupType | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
        Output "Created the distribution group $distributionGroupName successfully." "Green"
    }
    else
    {
        Output "Distribution group $distributionGroupName already exists." "Yellow"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Disable encryption on the Microsoft Exchange server. 
# </summary>
# <param name="server">The Exchange server name.</param>
# <param name="ExchangeVersion">The version of Microsoft Exchange Server.</param>
#-----------------------------------------------------------------------------------
Function DisableEncryption
{
    param(
    [string]$server,
    [string]$ExchangeVersion
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'server' $server
    ValidateParameter 'ExchangeVersion' $ExchangeVersion

    Output "Disable encryption on the Microsoft Exchange server..." "White"
    if($ExchangeVersion -eq $global:Exchange2007)
    {
        Set-MailboxServer $server -MapiEncryptionRequired $false
        net stop MSExchangeIS    
        net start MSExchangeIS
    }
    elseif($ExchangeVersion -ge $global:Exchange2010)
    {
        Set-RpcClientAccess -Server $server -EncryptionRequired $false
    }
    Output "Disabled encryption on the Exchange server successfully." "Green"    
}

#-----------------------------------------------------------------------------------
# <summary>
# Start Configuration of RPC over HTTP for Exchange server. 
# </summary>
#-----------------------------------------------------------------------------------
function ConfigureRPCOverHTTP
{
    # Get New Valid Ports
    $validPortsNew = GetNewRpchValidPorts

    Output "Configure RPC over HTTP on the Exchange server:" "White"
    Output "Steps for manual configuration:" "Yellow"
    $step = 1 
    Output "$step. In HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc, create a key named `"DefaultChannelLifetime`" and set the key value to 0x20000" "Yellow"
    Output "   In HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc, create a key named `"ActAsWebFarm`" and set the key value to 0" "Yellow"
    Output "   In HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy, create a key named `"AllowAnonymous`" and set the key value to 1" "Yellow"  
    Output "   In HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy, find the key named `"Enabled`" and set the key value to 1" "Yellow"
    Output "   In HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy, find the key named `"ValidPorts`" and set the key value to $validPortsNew" "Yellow"
    Output "   In HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\MSExchangeServiceHost\RpcHttpConfigurator, find the key named `"PeriodicPollingMinutes`" and set the key value to 0" "Yellow" 
    $step++
    Output "$step. In the `"SSL Settings`" page of `"Default Web Site/Rpc`" in IIS, disable `"Require SSL`", and set `"Ignore`" for Client certificates" "Yellow"  
    Output "   In the `"Authentication`" page of `"Default Web Site/Rpc`" in IIS, enable `"Anonymous Authentication`", `"Basic Authentication`" and `"Windows Authentication`", and disable the rest of the options." "Yellow"  
    Output "   In the `"Authentication`" page of `"Default Web Site/Rpc`" in IIS, edit the Anonymous Authentication credentials settings and set user name as IUSR with an empty password" "Yellow" 
    $step++
    Output "$step. Restart IIS" "Yellow"
    $step++
    Output "$step. Export the default site's HTTPS certificate, and save the certificate to system drive, e.g. c:\" "Yellow"

    # Step 1
    cmd /c reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc" /v "DefaultChannelLifetime" /t "REG_DWORD" /d 0x20000 /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    cmd /c reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc" /v "ActAsWebFarm" /t "REG_DWORD" /d 0 /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    cmd /c reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy" /v "AllowAnonymous" /t "REG_DWORD" /d 1 /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 
    cmd /c reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy" /v "Enabled" /t "REG_DWORD" /d 1 /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    cmd /c reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Rpc\RpcProxy" /v "ValidPorts" /t "REG_SZ" /d "$validPortsNew" /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100
    cmd /c reg add "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\MSExchangeServiceHost\RpcHttpConfigurator" /v "PeriodicPollingMinutes" /t "REG_DWORD" /d 0 /f 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100

    # Step 2
    cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/Rpc" /commit:APPHOST /section:system.webServer/security/access /sslFlags:"None" 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 
    cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/Rpc" /commit:APPHOST /section:system.webServer/security/authentication/basicAuthentication /Enabled:"true" 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 
    cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/Rpc" /commit:APPHOST /section:system.webServer/security/authentication/windowsAuthentication /Enabled:"true" 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 
    cmd /c $env:windir\system32\inetsrv\appcmd.exe set config "Default Web Site/Rpc" /commit:APPHOST /section:system.webServer/security/authentication/anonymousAuthentication /Enabled:"true" /userName:"IUSR" 2>&1 | Out-File -FilePath $logFile -Append -encoding ASCII -width 100 

    # Step 3
    IISReset /restart

    # Step 4: Export certificates
    $sysDrive = $env:SystemDrive
    (dir cert:\LocalMachine\my) | ForEach-Object{[system.IO.file]::WriteAllBytes("$sysDrive\$($_.Subject).cer", ($_.Export('Cert','secret')))}
    
    Output "Configured RPC over HTTP successfully." "Green"
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the value of LegacyExchangeDN of a mailbox user.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The name of a mailbox user.</param>
# <param name="credentialUserName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="credentialPassword">The password of the domain user name.</param>
# <returns>The value of LegacyExchangeDN of a mailbox user.</returns> 
#-----------------------------------------------------------------------------------
function GetUserDN
{
    param(
    [string]$computerName,
    [string]$userName,
    [string]$credentialUserName,
    [string]$credentialPassword 
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'computerName' $computerName	
    ValidateParameter 'userName' $userName	
    ValidateParameter 'credentialUserName' $credentialUserName	
    ValidateParameter 'credentialPassword' $credentialPassword	

    $securePassword = ConvertTo-SecureString $credentialPassword -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($credentialUserName,$securePassword)
    $userDN = Invoke-Command -ComputerName $computerName -Credential $credential -ErrorAction SilentlyContinue -ScriptBlock {
        # Create a New ADSI Call
        $dnsDomain = $args[1].Split("\")[0]
        $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$dnsDomain",$args[1],$args[2])
        # Create a New DirectorySearcher Object
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
        # Set the filter to search for a specific CNAME
        $temp = $args[0]
        $searcher.filter = "(&(objectClass=user) (CN=$temp))"
        # Set results in $adFind variable
        $adFind = $searcher.findall()
        $adFind[0].Properties.legacyexchangedn
    } -Args $userName,$credentialUserName,$credentialPassword

    if($userDN -eq "" -or $userDN -eq $null)
    {
        #-----------------------------------------------------
        # Get $userName's ESSDN manually
        #-----------------------------------------------------
        Output "Can't get $userName's ESSDN automatically." "Yellow"
        Output "For Windows platform, refer to the package test suite deployment guide to obtain the ESSDN value." "Yellow"
        Output "Enter the ESSDN of user $userName" "Cyan"    
        $userDN = CheckForEmptyUserInput "the ESSDN of user $userName" $userName
        Output "Your input is $userDN" "White"
    }
    return $userDN
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the Exchange server version. 
# </summary>
# <param name="computerName">The computer name which the Exchange server is installed on.</param>
# <param name="userName">The user name of the SUTComputerName, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <returns>
# The version name of the Exchange server.
# Return "Microsoft Exchange Server 2007" if Exchange version is "Microsoft Exchange Server 2007".
# Return "Microsoft Exchange Server 2010" if Exchange version is "Microsoft Exchange Server 2010".
# Return "Microsoft Exchange Server 2013" if Exchange version is "Microsoft Exchange Server 2013".
# </returns>
#-----------------------------------------------------------------------------------
function GetExchangeServerVersionOnSUT
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'computerName' $computerName
    ValidateParameter 'userName' $userName
    ValidateParameter 'password' $password

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    #-----------------------------------------------------
    # Test if $computerName exists
    #-----------------------------------------------------
    Invoke-Command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock {"Test if remote computer exists." | Out-Null}
    if(!$?)
    {
        $errMsg = $Error[0]
        if($errMsg.ErrorDetails.Message.Contains("the server name cannot be resolved"))
        {
            Output $errMsg.ErrorDetails.Message "yellow"
            Output "The specified server $computerName may not exist." "Cyan"
            Output "Would you like to continue to run the client setup script?" "Cyan"   
            Output "1: CONTINUE." "Cyan"
            Output "2: EXIT." "Cyan"    
            $runWhenConnectingToSUTFailedChoices = @('1','2')
            $runWhenConnectingToSUTFailed = ReadUserChoice $runWhenConnectingToSUTFailedChoices "runWhenConnectingToSUTFailed"
            if($runWhenConnectingToSUTFailed -eq "2")
            {
                Stop-Transcript
                exit 0
            }
        }            
    }

    $sutVersion = Invoke-Command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock {    
        #-----------------------------------------------------
        # Exchange display name in registry.
        #-----------------------------------------------------
        $Exchange2007 = "Microsoft Exchange Server 2007", "ExchangeServer2007"
        $Exchange2010 = "Microsoft Exchange Server 2010", "ExchangeServer2010"
        $Exchange2013 = "Microsoft Exchange Server 2013", "ExchangeServer2013"
        $Exchange2016 = "Microsoft Exchange Server 2016", "ExchangeServer2016"   
        $ExchangeVersion  = "Unknown Version"
        $keys = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $items = $keys | foreach-object {Get-ItemProperty $_.PsPath}    
        foreach ($item in $items)
        {
            if(($item.DisplayName -eq $null) -or ($item.DisplayName -eq ""))
            {
                continue
            }
            if($item.DisplayName -eq $Exchange2007[0])
            {
                $ExchangeVersion = $Exchange2007[1]
                break
            }
            if($item.DisplayName -eq $Exchange2010[0])
            {
                $ExchangeVersion = $Exchange2010[1]
                break
            }            
            if($item.DisplayName.StartsWith($Exchange2013[0]))
            {
                $ExchangeVersion = $Exchange2013[1]
                break
            }
            if($item.DisplayName.StartsWith($Exchange2016[0]))
            {
                $ExchangeVersion = $Exchange2016[1]
                break
            }        
        }    
        return $ExchangeVersion
    }

    $ExchangeVersions = @("ExchangeServer2007","ExchangeServer2010","ExchangeServer2013","ExchangeServer2016")
    if($ExchangeVersions -notcontains $sutVersion )
    {
        Output "Cannot get the Exchange version automatically." "Yellow"
        $sutVersionChoices = @('1: Microsoft Exchange Server 2007',
                               '2: Microsoft Exchange Server 2010',
                               '3: Microsoft Exchange Server 2013',
                               '4: Microsoft Exchange Server 2016')   
        Output "Select the Exchange version: " "Cyan"
        Output ($sutVersionChoices[0]) "Cyan"    
        Output ($sutVersionChoices[1]) "Cyan"    
        Output ($sutVersionChoices[2]) "Cyan"
        Output ($sutVersionChoices[3]) "Cyan"
            
        $sutVersion = ReadUserChoice $sutVersionChoices "sutVersion"
        Switch ($sutVersion)
        {
            "1" { $sutVersion = $ExchangeVersions[0]; break }
            "2" { $sutVersion = $ExchangeVersions[1]; break }
            "3" { $sutVersion = $ExchangeVersions[2]; break }
            "4" { $sutVersion = $ExchangeVersions[3]; break }
        }
    }
    else
    {
        Output "The Exchange version installed on the server is $sutVersion." "Green"
    }
    return $sutVersion
}

$global:Exchange2007 = "Microsoft Exchange Server 2007"
$global:Exchange2010 = "Microsoft Exchange Server 2010"
$global:Exchange2013 = "Microsoft Exchange Server 2013"
$global:Exchange2016 = "Microsoft Exchange Server 2016"
[void][System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")