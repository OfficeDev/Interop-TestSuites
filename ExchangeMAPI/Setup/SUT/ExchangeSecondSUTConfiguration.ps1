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
$logFile               = $logPath+"\ExchangeSecondSUTConfiguration.ps1.log"
$debugLogFile          = $logPath+"\ExchangeSecondSUTConfiguration.ps1.debug.log"
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

#----------------------------------------------------------------------------
# Values for Configuration 
#----------------------------------------------------------------------------
$domain                             = $env:USERDNSDOMAIN
$sut2ComputerName                   = $env:ComputerName

$environmentResourceFile            = "$commonScriptDirectory\ExchangeTestSuite.config"

$defaultPublicFolderDatabase2       = ReadConfigFileNode "$environmentResourceFile" "defaultPublicFolderDatabaseOnSecondSUT"
$defaultPublicFolderMailbox2Prefix  = ReadConfigFileNode "$environmentResourceFile" "defaultPublicFolderMailboxPrefixOnSecondSUT"
$defaultPublicFolderMailbox2        = $defaultPublicFolderMailbox2Prefix + $sut2ComputerName

$MSOXCFOLDPublicFolderGhosted       = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDPublicFolderGhosted"
$MSOXCROPSPublicFolderGhosted       = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSPublicFolderGhosted"
$MSOXCFXICSGhostedPublicFolder      = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSGhostedPublicFolder"

$MSOXCFXICSUser2                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSUser2"
$MSOXCFXICSUser2Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSUser2Password"

$MSOXCSTORMailboxOnServer2          = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer2"
$MSOXCSTORMailboxOnServer2Password  = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer2Password"

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
# Begin to configure second server
#-----------------------------------------------------
Output "Begin to configure the second Exchange server ..." "White"
Output "Steps for manual configuration:" "Yellow" 
Output "Enable remoting in Powershell." "Yellow"
Invoke-Command {
    $ErrorActionPreference = "Continue"
    Enable-PSRemoting -Force
}

InstallWindowsFeature RSAT-AD-PowerShell
StartService "msexchange*" "auto"

#-----------------------------------------------------
# Get Second Exchange server basic information
#-----------------------------------------------------
$global:ExchangeVersion = GetExchangeServerVersion
Output "The basic information of the second Exchange server:" "White"
Output "Domain name: $domain" "White"
Output "The name of the second Exchange server: $sut2ComputerName" "White"
Output "The version of the second Exchange server: $global:ExchangeVersion" "White"

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

#-----------------------------------------------------
# Get Main Exchange server Name and Check
#-----------------------------------------------------
$checkMainServerOK = $false
Output "Enter the computer name of the first SUT." "cyan"
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported." "Cyan"    
while(1)
{
    $nodeInXml = "sutComputerName"
    [String]$sutComputerName = GetUserInput $nodeInXml
    if($sutComputerName -as [ipaddress])
    {
        Output "IP address is not supported." "Yellow"
    }
    elseif ($sutComputerName -imatch '[`~!@#$%^&*()=+_\[\]{}\\|;:.''",<>/?]')
    {
        Output """$sutComputerName"" contains characters that are not allowed, such as `` ~ ! @ # $ % ^ & * ( ) = + _ [ ] { } \ | ; : . ' "" , < > / and ?." "Yellow"
    }
    elseif ($sutComputerName.Length -lt 1 -or $sutComputerName.Length -gt 15)
    {
        Output "Computer name has a minimum length (1 character) and maximum length (15 characters)." "Yellow"
    }
    elseif ($sutComputerName -eq $sut2ComputerName)
    {
        Output """$sutComputerName"" is the computer name of the current SUT." "Yellow"
    }
    else
    {
        $ExchangeServers = Get-ExchangeServer 
        foreach($_ in $ExchangeServers)
        {
            if($_.Name -eq $sutComputerName -and $_.Name -ne $sut2ComputerName)
            {
                $checkMainServerOK = $true
                break
            }
        }
        if($checkMainServerOK -eq $true)
        {
            break
        }
        else
        {
            Output "Exchange server $sutComputerName doesn't exist in this domain environment! Check and enter the right name:" "Yellow"
        }
    }
    if($unattendedXmlName -eq "")
    {
        Output "Retry with a valid computer name." "Yellow"
    }
    else
    {
        Write-Warning "Change the value of `"$nodeInXml`" with a valid computer name in the ExchangeSecondSUTConfigurationAnswers.xml file and then run the script again.`r`n"
        Stop-Transcript
        exit 2
    }
}

#----------------------------------------------------------------------------
# Start Configuration of RPC over HTTP for Second Exchange server
#----------------------------------------------------------------------------
ConfigureRPCOverHTTP
Output "The second Exchange server is now configured for RPC over HTTP transport." "Green"

#----------------------------------------------------------------------------
# Start Configuration of Mailbox Users for Second Exchange server
#----------------------------------------------------------------------------
# Create Mailbox Users
Output "Mailbox users are currently being created on the Exchange server; please wait..." "White"
$mailboxDatabases2 = Get-MailboxDatabase -Server $sut2ComputerName
if(@($mailboxDatabases2).count -gt 1)
{
    $mailboxDatabaseName2 = $mailboxDatabases2[0].Identity.ToString()
}
else
{
    $mailboxDatabaseName2 = $mailboxDatabases2.Identity.ToString()
}
CreateMailboxUser  $MSOXCFXICSUser2 $MSOXCFXICSUser2Password $mailboxDatabaseName2 $domain
CreateMailboxUser  $MSOXCSTORMailboxOnServer2  $MSOXCSTORMailboxOnServer2Password  $mailboxDatabaseName2 $domain

Output "All required mailbox users on the second Exchange server are created and configured." "Green"

#----------------------------------------------------------------------------
# Start Configuration of Public Folder Databases and Public Folders for Second Exchange server
#----------------------------------------------------------------------------

if($global:ExchangeVersion -le $global:Exchange2010)
{
    # Create Public Folder Database
    $publicFolderDatabase2Name = CreatePublicFolderDatabase $defaultPublicFolderDatabase2 $sut2ComputerName

    # Set $publicFolderDatabase2Name to be the associated public folder database for the mailbox database of second server
    Set-Mailboxdatabase -Identity $mailboxDatabaseName2 -publicFolderdatabase $publicFolderDatabase2Name        
    output "$publicFolderDatabase2Name is now the associated public folder database for the mailbox database of $sut2ComputerName." "Green"

    # Create Public Folders
    CreatePublicFolder $MSOXCROPSPublicFolderGhosted $sut2ComputerName
    CreatePublicFolder $MSOXCFOLDPublicFolderGhosted $sut2ComputerName

    CheckGhostedPublicFolderStatus $MSOXCROPSPublicFolderGhosted  $sutComputerName
    CheckGhostedPublicFolderStatus $MSOXCFOLDPublicFolderGhosted  $sutComputerName
    CheckGhostedPublicFolderStatus $MSOXCFXICSGhostedPublicFolder $sut2ComputerName

    # Configuration of Public Folders    
    Output "On the $sutComputerName, replicate the�$MSOXCFXICSGhostedPublicFolder�with $publicFolderDatabase2Name database." "Yellow"
    Set-PublicFolder -Server $sutComputerName  -Identity "\$MSOXCFXICSGhostedPublicFolder" -Replicas $publicFolderDatabase2Name  
}
elseif($global:ExchangeVersion -ge $global:Exchange2013)
{
    Output "Public folder replication is not supported in $global:ExchangeVersion" "Yellow"
    $publicFolderMailbox2Name = CreatePublicFolderMailbox $defaultPublicFolderMailbox2 $sut2ComputerName $mailboxDatabaseName2

    # Create Public Folders
    CreatePublicFolder $MSOXCROPSPublicFolderGhosted $sut2ComputerName $publicFolderMailbox2Name
    CreatePublicFolder $MSOXCFOLDPublicFolderGhosted $sut2ComputerName $publicFolderMailbox2Name

}
Output "All required public folder databases and public folders on the second Exchange server are created and configured." "Green"

#----------------------------------------------------------------------------
# Start Other Configuration for Exchange server
#----------------------------------------------------------------------------
#Disable encryption on the Microsoft Exchange server
DisableEncryption $sut2ComputerName $global:ExchangeVersion

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