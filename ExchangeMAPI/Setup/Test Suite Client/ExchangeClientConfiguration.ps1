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

#----------------------------------------------------------------------------
# Start script
#----------------------------------------------------------------------------
$ErrorActionPreference  = "Stop"
[String]$containerPath  = Get-Location
$logPath                = $containerPath + "\SetupLogs"
$logFile                = $logPath + "\ExchangeClientConfiguration.ps1.log"
$debugLogFile           = $logPath + "\ExchangeClientConfiguration.ps1.debug.log"

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
# Paths for all PTF configuration files.
#-----------------------------------------------------
$commonDeploymentFile        = "..\..\Source\Common\Common\ExchangeCommonConfiguration.deployment.ptfconfig"
$MSOXCFOLDDeploymentFile     = "..\..\Source\MS-OXCFOLD\TestSuite\MS-OXCFOLD_TestSuite.deployment.ptfconfig"
$MSOXCFXICSDeploymentFile    = "..\..\Source\MS-OXCFXICS\TestSuite\MS-OXCFXICS_TestSuite.deployment.ptfconfig"
$MSOXCMAPIHTTPDeploymentFile = "..\..\Source\MS-OXCMAPIHTTP\TestSuite\MS-OXCMAPIHTTP_TestSuite.deployment.ptfconfig"
$MSOXCMSGDeploymentFile      = "..\..\Source\MS-OXCMSG\TestSuite\MS-OXCMSG_TestSuite.deployment.ptfconfig"
$MSOXCNOTIFDeploymentFile    = "..\..\Source\MS-OXCNOTIF\TestSuite\MS-OXCNOTIF_TestSuite.deployment.ptfconfig"
$MSOXCPERMDeploymentFile     = "..\..\Source\MS-OXCPERM\TestSuite\MS-OXCPERM_TestSuite.deployment.ptfconfig"
$MSOXCPRPTDeploymentFile     = "..\..\Source\MS-OXCPRPT\TestSuite\MS-OXCPRPT_TestSuite.deployment.ptfconfig" 
$MSOXCROPSDeploymentFile     = "..\..\Source\MS-OXCROPS\TestSuite\MS-OXCROPS_TestSuite.deployment.ptfconfig" 
$MSOXCRPCDeploymentFile      = "..\..\Source\MS-OXCRPC\TestSuite\MS-OXCRPC_TestSuite.deployment.ptfconfig"
$MSOXCSTORDeploymentFile     = "..\..\Source\MS-OXCSTOR\TestSuite\MS-OXCSTOR_TestSuite.deployment.ptfconfig"
$MSOXCTABLDeploymentFile     = "..\..\Source\MS-OXCTABL\TestSuite\MS-OXCTABL_TestSuite.deployment.ptfconfig"
$MSOXNSPIDeploymentFile      = "..\..\Source\MS-OXNSPI\TestSuite\MS-OXNSPI_TestSuite.deployment.ptfconfig"
$MSOXORULEDeploymentFile     = "..\..\Source\MS-OXORULE\TestSuite\MS-OXORULE_TestSuite.deployment.ptfconfig"

$environmentResourceFile     = "$commonScriptDirectory\ExchangeTestSuite.config"

#-----------------------------------------------------
# Check and make sure that the SUT configuration is finished before running the client setup script.
#-----------------------------------------------------
Output "The SUT must be configured before running the client setup script." "Cyan"
Output "Did you either run the SUT setup script or configure the SUT as described by the Test Suite Deployment Guide? (Y/N)" "Cyan"
$isSutConfiguredChoices = @("Y","N")
$isSutConfigured = ReadUserChoice $isSutConfiguredChoices "isSutConfigured"
if($isSutConfigured -eq "N")
{
    Output "You input `"N`"." "White"
    Output "Exiting the client setup script now." "Yellow"
    Output "Configure the SUT and run the client setup script again." "Yellow"
    Stop-Transcript
    exit 0
}

#-----------------------------------------------------
# Check the operating system (OS) version
#-----------------------------------------------------
Output "Check the operating system (OS) version of the local machine ..." "White"
CheckOSVersion -computer localhost

#-----------------------------------------------------
# Check the Application environment
#-----------------------------------------------------
Output "Check whether the required applications have been installed ..." "White"
$vsInstalledStatus = CheckVSVersion "12.0"
$ptfInstalledStatus = CheckPTFVersion "1.0.2220.0"
$seInstalledStatus = CheckSEVersion "3.6.14230.01"
if(!$vsInstalledStatus -or !$ptfInstalledStatus -or !$seInstalledStatus)
{
    Output "Would you like to exit and install the application(s) that highlighted as yellow in above or continue without installing the application(s)?" "Cyan"    
    Output "1: CONTINUE (Without installing the recommended application(s) , it may cause some risk on running the test cases)." "Cyan"
    Output "2: EXIT." "Cyan"    
    $runWithoutRequiredAppInstalledChoices = @('1','2')
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
Output "Configure the ExchangeCommonConfiguration.deployment.ptfconfig file ..." "White"
Output "Enter the computer name of the first SUT:" "Cyan"    
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported." "Cyan"    
$sutComputerName = ReadComputerName $false "sutComputerName"
Output "The computer name of the first SUT you entered: $sutComputerName" "White"

Output "Enter the computer name of the second SUT. Press `"Enter`" if it doesn't exist." "Cyan"    
Output "The computer name must be valid. Fully qualified domain name(FQDN) or IP address is not supported." "Cyan"    
$sut2ComputerName = ReadComputerName $true "sut2ComputerName"
if($sut2ComputerName -ne "")
{
    Output "The computer name of the second SUT you entered: $sut2ComputerName" "White"
}

if($env:USERDNSDOMAIN -ne $null)
{
    Output "Current logon user:" "Cyan"    
    Output "Domain: $env:USERDOMAIN" "Cyan"    
    Output "Name:   $env:USERNAME" "Cyan"    
    Output "Would you like to use this user? (Y/N)" "Cyan"
    $useCurrentUserChoices = @("Y","N")
    $useCurrentUser = ReadUserChoice $useCurrentUserChoices "useCurrentUser"   
    if($useCurrentUser -eq "Y")
    {
        $dnsDomain = $ENV:USERDNSDOMAIN
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
    Output "Enter the name of DNS domain where the SUT belongs to (for example: contoso.com):" "Cyan"
    [String]$dnsDomain = CheckForEmptyUserInput "Domain name" "dnsDomain"
    Output "The DNS Domain name you entered: $dnsDomain" "White"
    Output "Enter the user name. It must be the SUT administrator." "Cyan"
    $userName = CheckForEmptyUserInput "User name" "userName"
    Output "The user name you entered: $userName" "White"
}

Output "Enter password:" "Cyan"    
$password = CheckForEmptyUserInput "Password" "password"
Output "Password you entered: $password" "White"

Output "Steps for manual configuration:" "Yellow" 
Output "Add SUT machines to the TrustedHosts configuration setting to ensure WinRM client can process remote calls against SUT machines." "Yellow"
$service = "WinRM"
$serviceStatus = (Get-Service $service).Status
if($serviceStatus -ne "Running")
{
    Start-Service $service
}
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

Output "Try to get the Exchange server version on the selected server $sutComputerName ..." "White"
$sutVersion = GetExchangeServerVersionOnSUT $sutComputerName "$dnsDomain\$userName" $password
if($sut2ComputerName -ne "")
{
    Output "Try to get the Exchange server version on the selected server $sut2ComputerName ..." "White"
    $sutVersionOnServer2 = GetExchangeServerVersionOnSUT $sut2ComputerName "$dnsDomain\$userName" $password
    if($sutVersion -ne $sutVersionOnServer2)
    {
        Output "The Exchange server version on $sutComputerName and $sut2ComputerName are different." "Yellow"
        Output "Would you like to continue executing client setup script?" "Cyan"   
        Output "1: CONTINUE(Will use Exchange server version $sutVersion on $sutComputerName for the ExchangeCommonConfiguration.deployment.ptfconfig file." "Cyan"
        Output "2: EXIT." "Cyan"    
        $runWhenExchangeVersionDifferentChoices = @('1','2')
        $runWhenExchangeVersionDifferent = ReadUserChoice $runWhenExchangeVersionDifferentChoices "runWhenExchangeVersionDifferent"
        if($runWhenExchangeVersionDifferent -eq "2")
        {
            Stop-Transcript
            exit 0
        }
    }
}

$transportSeqs = @('1: ncacn_http, to use RPC over HTTP as transport',
                   '2: ncacn_ip_tcp, to use RPC over TCP as transport',
                   '3: mapi_http, to use MAPIHTTP as transport')
Output "Select transport method for Exchange RPC calls." "Cyan"    
if($sutVersion -lt "ExchangeServer2013")
{
    Output "This value could be `"ncacn_http`" or `"ncacn_ip_tcp`". For more information, refer to section 2.1 in [MS-OXRPC]." "Cyan"
    Output ($transportSeqs[0]) "Cyan"    
    Output ($transportSeqs[1]) "Cyan"
    $transportSeqs = $transportSeqs[0,1]
    $transportSeq = ReadUserChoice $transportSeqs "transportSeq"
    Switch ($transportSeq)
    {
        "1" { $transportSeq = "ncacn_http";   break }
        "2" { $transportSeq = "ncacn_ip_tcp"; break }
    }
}
else
{  
    Output "Exchange Server 2013 does not support RPC over TCP." "Yellow"
    Output "This value could be `"ncacn_http`" or `"mapi_http`". For ncacn_http, refer to section 2.1 in [MS-OXRPC]. mapi_http indicates the use of MS-OXCMAPIHTTP, For mapi_http, refer to [MS-OXCMAPIHTTP]." "Cyan"
    Output ($transportSeqs[0]) "Cyan"
    Output ($transportSeqs[2] + ", MAPIHTTP is supported from Exchange Server 2013 SP1") "Cyan"
    $transportSeqs = $transportSeqs[0,2]
    $transportSeq = ReadUserChoice $transportSeqs "transportSeq"
    Switch ($transportSeq)
    {
        "1" { $transportSeq = "ncacn_http"; break }
        "3" { $transportSeq = "mapi_http";  break}
    }
}

if($transportSeq -eq "mapi_http")
{
    Output "MAPIHTTP requires to use Autodiscover to get the service's url for mailbox server endpoint and address book server endpoint." "Yellow"
    Output "Set the value of property useAutodiscover to true." "Yellow"
    $useAutodiscover = "true"
}
else
{
    $useAutodiscovers = @('1: true, to use Autodiscover',
                            '2: false, to not use Autodiscover')
    Output "Select whether to use Autodiscover for mailbox setting." "Cyan"
    Output ($useAutodiscovers[0]) "Cyan"
    Output ($useAutodiscovers[1]) "Cyan"
    $useAutodiscover = ReadUserChoice $useAutodiscovers "useAutodiscover"
    Switch ($useAutodiscover)
    {
        "1" { $useAutodiscover = "true";  break }
        "2" { $useAutodiscover = "false"; break }
    }
}

$compressRpcRequests = @('1: true, to enable the compression in the RPC request',
                         '2: false, to disable the compression in the RPC request')
Output "Select whether compression should be enabled in the RPC request sent by the client." "Cyan"    
Output "For more information, refer to RPC operations ""EcDoConnectEx"" and ""EcDoRpcExt2"" in [MS-OXCRPC]." "Cyan"
Output ($compressRpcRequests[0]) "Cyan"
Output ($compressRpcRequests[1]) "Cyan"
$compressRpcRequest = ReadUserChoice $compressRpcRequests "compressRpcRequest"
Switch ($compressRpcRequest)
{
    "1" { $compressRpcRequest = "true";  break }
    "2" { $compressRpcRequest = "false"; break }
}

$xorRpcRequests = @('1: true, RPC request sent by the client needs obfuscation',
                    '2: false, RPC request sent by the client does not need obfuscation')
Output "Select whether an RPC request sent by the client needs obfuscation." "Cyan"    
Output "For more information, refer to RPC operations ""EcDoConnectEx"" and ""EcDoRpcExt2"" in [MS-OXCRPC]." "Cyan"
Output ($xorRpcRequests[0]) "Cyan"
Output ($xorRpcRequests[1]) "Cyan"
$xorRpcRequest = ReadUserChoice $xorRpcRequests "xorRpcRequest"
Switch ($xorRpcRequest)
{
    "1" { $xorRpcRequest = "true";  break }
    "2" { $xorRpcRequest = "false"; break }
}

if($transportSeq -eq "ncacn_http")
{
    $rpchUseSsls = @('1: true, to use RPC over HTTP with SSL',
                        '2: false, to use RPC over HTTP without SSL')
    Output "Select whether to use RPC over HTTP with SSL." "Cyan"
    Output ($rpchUseSsls[0]) "Cyan"
    Output ($rpchUseSsls[1]) "Cyan"
    $rpchUseSsl = ReadUserChoice $rpchUseSsls "rpchUseSsl"
    Switch ($rpchUseSsl)
    {
        "1" { $rpchUseSsl = "true";  break }
        "2" { $rpchUseSsl = "false"; break }
    }

    $rpchAuthSchemes = @('1: Basic, to use Basic authentication scheme',
                            '2: NTLM, to use Windows authentication scheme')
    Output "Select authentication scheme used in the http authentication for RPC over HTTP." "Cyan"   
    Output ($rpchAuthSchemes[0]) "Cyan"
    Output ($rpchAuthSchemes[1]) "Cyan"
    $rpchAuthScheme = ReadUserChoice $rpchAuthSchemes "rpchAuthScheme"
    Switch ($rpchAuthScheme)
    {
        "1" { $rpchAuthScheme = "Basic"; break }
        "2" { $rpchAuthScheme = "NTLM";  break }
    }
}

if($transportSeq -ne "mapi_http")
{
    $authenticationServices = @('0: RPC_C_AUTHN_NONE, no authentication, only Exchange Server 2013 support RPC_C_AUTHN_NONE',
                                '9: RPC_C_AUTHN_GSS_NEGOTIATE, to use the Microsoft Negotiate security support provider (SSP)',
                                '10: RPC_C_AUTHN_WINNT, to use the Microsoft NT LAN Manager (NTLM) SSP',
                                '16: RPC_C_AUTHN_GSS_KERBEROS, to use the Microsoft Kerberos SSP')
    Output "Select the authentication service for RPC calls to Exchange server." "Cyan"   
    if($sutVersion -ge "ExchangeServer2013")
    {
        Output ($authenticationServices[0]) "Cyan"    
    }
    Output ($authenticationServices[1]) "Cyan"    
    Output ($authenticationServices[2]) "Cyan"    
    Output ($authenticationServices[3]) "Cyan"    
    if($sutVersion -lt "ExchangeServer2013")
    {
        $authenticationServices = $authenticationServices[1..($authenticationServices.Length-1)]
    }

    $authenticationService = ReadUserChoice $authenticationServices "authenticationService"

    if($authenticationService -eq "0")
    {
        Output "RPC_C_AUTHN_NONE requires authentication level RPC_C_AUTHN_LEVEL_NONE." "Yellow"
        Output "Set authentication level to 1: RPC_C_AUTHN_LEVEL_NONE automatically." "Yellow"
        $authenticationLevel = "1"
    }
    else
    {
        $authenticationLevels = @('0: RPC_C_AUTHN_LEVEL_DEFAULT, same as RPC_C_AUTHN_LEVEL_CONNECT',
                                  '2: RPC_C_AUTHN_LEVEL_CONNECT',
                                  '3: RPC_C_AUTHN_LEVEL_CALL',
                                  '4: RPC_C_AUTHN_LEVEL_PKT',
                                  '5: RPC_C_AUTHN_LEVEL_PKT_INTEGRITY',
                                  '6: RPC_C_AUTHN_LEVEL_PKT_PRIVACY')
        Output "Select authentication level for creating an RPC binding." "Cyan"    
        Output "For more information on setting the property value, refer to section 2.2.1.1.8 in [MS-RPCE]." "Cyan"    
        Output ($authenticationLevels[0] + ", same as RPC_C_AUTHN_LEVEL_CONNECT") "Cyan"    
        Output ($authenticationLevels[1] +", for disable encryption, authenticates the credentials of the client and server") "Cyan"    
        Output ($authenticationLevels[2] +", same as RPC_C_AUTHN_LEVEL_PKT") "Cyan"    
        Output ($authenticationLevels[3] +", same as RPC_C_AUTHN_LEVEL_CONNECT but also prevents replay attacks") "Cyan"    
        Output ($authenticationLevels[4] +", same as RPC_C_AUTHN_LEVEL_PKT but also verifies that none of the data transferred between the client and server has been modified") "Cyan"    
        Output ($authenticationLevels[5] +", for encryption, same as RPC_C_AUTHN_LEVEL_PKT_INTEGRITY but also ensures that the data transferred can only be seen unencrypted by the client and the server") "Cyan"    
        Output "The recommended value is 6. You need to disable encryption on the Exchange server if you use any other value." "Cyan"    
        $authenticationLevel = ReadUserChoice $authenticationLevels "authenticationLevel"
    }

    $setUuids = @('1: true, to set the field',
                  '2: false, to not set the field')
    Output "Select whether to set PFC_OBJECT_UUID(0x80) field of the RPC header pfc_flags." "Cyan"
    Output ($setUuids[0]) "Cyan"
    Output ($setUuids[1]) "Cyan"
    Output "For more information about pfc_flags, refer to DCE RPC C706 section 12.6.3.1 Declarations." "cyan"
    $setUuid = ReadUserChoice $setUuids "setUuid"
    Switch ($setUuid)
    {
        "1" { $setUuid = "true";  break }
        "2" { $setUuid = "false"; break }
    }

    $rpcForceShutdownAssociations = @('1: true, to shut down the TCP connection after the last binding handle is freed',
                                      '2: false, to reuse the existing TCP connection')
    Output "Select whether to set RPC_C_OPT_DONT_LINGER option on the binding handle." "Cyan"
    Output ($rpcForceShutdownAssociations[0]) "cyan"
    Output ($rpcForceShutdownAssociations[1]) "cyan"
    $rpcForceShutdownAssociation = ReadUserChoice $rpcForceShutdownAssociations "RpcForceShutdownAssociation"
    Switch ($rpcForceShutdownAssociation)
    {
        "1" { $rpcForceShutdownAssociation = "true";  break }
        "2" { $rpcForceShutdownAssociation = "false"; break }
    }
}

# Get Local machine IPv4/IPv6 Address
$ipAddress = ipconfig | Select-String "ip" | Select-String "address"
if($ipAddress -eq $null)
{
    Output "Your local machine does not have a valid IP address, check your network environment and run the script again." "red"
    Stop-Transcript
    exit 2
}
else
{
    $ipv4Address = GetIpAddress "IPv4"
    $ipv6Address = GetIpAddress "IPv6"
}

Output "Modify the properties as necessary in the ExchangeCommonConfiguration.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $commonDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"SutComputerName`", and set the value as $sutComputerName" "Yellow"
$step++
Output "$step.Find the property `"Sut2ComputerName`", and set the value as $sut2ComputerName" "Yellow"
$step++
Output "$step.Find the property `"Domain`", and set the value as $dnsDomain" "Yellow"
$step++
Output "$step.Find the property `"SutVersion`", and set the value as $sutVersion" "Yellow"
$step++
Output "$step.Find the property `"TransportSeq`", and set the value as $transportSeq" "Yellow"
if($transportSeq -eq "ncacn_http")
{
    $step++
    Output "$step.Find the property `"RpchUseSsl`", and set the value as $rpchUseSsl" "Yellow"
    $step++
    Output "$step.Find the property `"RpchAuthScheme`", and set the value as $rpchAuthScheme" "Yellow"
}
if($transportSeq -ne "mapi_http")
{
    $step++
    Output "$step.Find the property `"RpcAuthenticationLevel`", and set the value as $authenticationLevel" "Yellow"
    $step++
    Output "$step.Find the property `"RpcAuthenticationService`", and set the value as $authenticationService" "Yellow"
    $step++
    Output "$step.Find the property `"SetUuid`", and set the value as $setUuid" "Yellow"
    $step++
    Output "$step.Find the property `"RpcForceShutdownAssociation`", and set the value as $rpcForceShutdownAssociation" "Yellow"
}
$step++
Output "$step.Find the property `"CompressRpcRequest`", and set the value as $compressRpcRequest" "Yellow"
$step++
Output "$step.Find the property `"XorRpcRequest`", and set the value as $xorRpcRequest" "Yellow"
$step++
Output "$step.Find the property `"useAutodiscover`", and set the value as $useAutodiscover" "Yellow"
$step++
Output "$step.Find the property `"NotificationIP`", and set the value as $ipv4Address" "Yellow"
$step++
Output "$step.Find the property `"NotificationIPv6`", and set the value as $ipv6Address" "Yellow"

ModifyConfigFileNode $commonDeploymentFile "SutComputerName"             $sutComputerName
ModifyConfigFileNode $commonDeploymentFile "Sut2ComputerName"            $sut2ComputerName
ModifyConfigFileNode $commonDeploymentFile "Domain"                      $dnsDomain
ModifyConfigFileNode $commonDeploymentFile "SutVersion"                  $sutVersion
ModifyConfigFileNode $commonDeploymentFile "TransportSeq"                $transportSeq
if($transportSeq -eq "ncacn_http")
{
    ModifyConfigFileNode $commonDeploymentFile "RpchUseSsl"              $rpchUseSsl
    ModifyConfigFileNode $commonDeploymentFile "RpchAuthScheme"          $rpchAuthScheme
}
if($transportSeq -ne "mapi_http")
{
    ModifyConfigFileNode $commonDeploymentFile "RpcAuthenticationLevel"      $authenticationLevel
    ModifyConfigFileNode $commonDeploymentFile "RpcAuthenticationService"    $authenticationService
    ModifyConfigFileNode $commonDeploymentFile "SetUuid"                     $setUuid
    ModifyConfigFileNode $commonDeploymentFile "RpcForceShutdownAssociation" $rpcForceShutdownAssociation
}
ModifyConfigFileNode $commonDeploymentFile "CompressRpcRequest"          $compressRpcRequest
ModifyConfigFileNode $commonDeploymentFile "XorRpcRequest"               $xorRpcRequest
ModifyConfigFileNode $commonDeploymentFile "useAutodiscover"             $useAutodiscover
ModifyConfigFileNode $commonDeploymentFile "NotificationIP"              $ipv4Address
ModifyConfigFileNode $commonDeploymentFile "NotificationIPv6"            $ipv6Address

Output "Configuration for ExchangeCommonConfiguration.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCFOLD ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCFOLD_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCFOLDCommonUser                = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDCommonUser"
$MSOXCFOLDCommonUserPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDCommonUserPassword"
$MSOXCFOLDAdminUser                 = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDAdminUser"
$MSOXCFOLDAdminUserPassword         = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDAdminUserPassword"
$MSOXCFOLDPublicFolderMailEnabled   = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDPublicFolderMailEnabled"
# Get User DN for MS-OXCFOLD
Output "Get Exchange user($MSOXCFOLDCommonUser)'s ESSDN automatically." "White"
$MSOXCFOLDCommonUserEssdn = GetUserDN $sutComputerName $MSOXCFOLDCommonUser "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCFOLDAdminUser)'s ESSDN automatically." "White"
$MSOXCFOLDAdminUserEssdn = GetUserDN $sutComputerName $MSOXCFOLDAdminUser "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCFOLD_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCFOLDDeploymentFile" "Yellow"
if($sut2ComputerName -ne "")
{
    $MSOXCFOLDPublicFolderGhosted = ReadConfigFileNode "$environmentResourceFile" "MSOXCFOLDPublicFolderGhosted"
    $step++
    Output "$step.Find the property `"GhostedPublicFolder`", and set the value as $MSOXCFOLDPublicFolderGhosted" "Yellow"
}
$step++
Output "$step.Find the property `"MailEnabledPublicFolder`", and set the value as $MSOXCFOLDPublicFolderMailEnabled" "Yellow"
$step++
Output "$step.Find the property `"CommonUser`", and set the value as $MSOXCFOLDCommonUser" "Yellow"
$step++
Output "$step.Find the property `"CommonUserPassword`", and set the value as $MSOXCFOLDCommonUserPassword" "Yellow"
$step++
Output "$step.Find the property `"CommonUserEssdn`", and set the value as $MSOXCFOLDCommonUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCFOLDAdminUser" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCFOLDAdminUserPassword" "Yellow"
$step++
Output "$step.Find the property `"AdminUserEssdn`", and set the value as $MSOXCFOLDAdminUserEssdn" "Yellow"

if($sut2ComputerName -ne "")
{
    ModifyConfigFileNode $MSOXCFOLDDeploymentFile "GhostedPublicFolder" $MSOXCFOLDPublicFolderGhosted
}    
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "MailEnabledPublicFolder" $MSOXCFOLDPublicFolderMailEnabled  
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "CommonUser"              $MSOXCFOLDCommonUser 
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "CommonUserPassword"      $MSOXCFOLDCommonUserPassword  
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "CommonUserEssdn"         $MSOXCFOLDCommonUserEssdn  
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "AdminUserName"           $MSOXCFOLDAdminUser  
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "AdminUserPassword"       $MSOXCFOLDAdminUserPassword  
ModifyConfigFileNode $MSOXCFOLDDeploymentFile "AdminUserEssdn"          $MSOXCFOLDAdminUserEssdn

Output "Configuration for MS-OXCFOLD_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCFXICS ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCFXICS_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCFXICSAdminUser                = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSAdminUser"
$MSOXCFXICSAdminUserPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSAdminUserPassword"
$MSOXCFXICSGhostedPublicFolder      = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSGhostedPublicFolder"
$MSOXCFXICSPublicFolder             = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSPublicFolder"
# Get User DN for MS-OXCFXICS
Output "Get Exchange user($MSOXCFXICSAdminUser)'s ESSDN automatically." "White"
$MSOXCFXICSAdminUserEssdn = GetUserDN $sutComputerName $MSOXCFXICSAdminUser "$dnsDomain\$userName" $password
if($sut2ComputerName -ne "")
{
    $MSOXCFXICSUser2                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSUser2"
    $MSOXCFXICSUser2Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCFXICSUser2Password"
    # Get User DN for MS-OXCFXICS
    Output "Get Exchange user($MSOXCFXICSUser2)'s ESSDN automatically." "White"
    $MSOXCFXICSUser2Essdn = GetUserDN $sutComputerName $MSOXCFXICSUser2 "$dnsDomain\$userName" $password
}

Output "Modify the properties as necessary in the MS-OXCFXICS_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCFXICSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCFXICSAdminUser" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCFXICSAdminUserPassword" "Yellow"
$step++
Output "$step.Find the property `"AdminUserESSDN`", and set the value as $MSOXCFXICSAdminUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"PublicFolderName`", and set the value as $MSOXCFXICSPublicFolder" "Yellow"
if($sut2ComputerName -ne "")
{
    $step++
    Output "$step.Find the property `"User2Name`", and set the value as $MSOXCFXICSUser2" "Yellow"
    $step++
    Output "$step.Find the property `"User2Password`", and set the value as $MSOXCFXICSUser2Password" "Yellow"
    $step++
    Output "$step.Find the property `"TestUser2ESSDN`", and set the value as $MSOXCFXICSUser2Essdn" "Yellow"
    $step++
    Output "$step.Find the property `"GhostedPublicFolderName`", and set the value as $MSOXCFXICSGhostedPublicFolder" "Yellow"
}

ModifyConfigFileNode $MSOXCFXICSDeploymentFile "PublicFolderName"       $MSOXCFXICSPublicFolder
ModifyConfigFileNode $MSOXCFXICSDeploymentFile "AdminUserName"          $MSOXCFXICSAdminUser
ModifyConfigFileNode $MSOXCFXICSDeploymentFile "AdminUserPassword"      $MSOXCFXICSAdminUserPassword    
ModifyConfigFileNode $MSOXCFXICSDeploymentFile "AdminUserESSDN"         $MSOXCFXICSAdminUserEssdn
if($sut2ComputerName -ne "")
{
    ModifyConfigFileNode $MSOXCFXICSDeploymentFile "User2Name"               $MSOXCFXICSUser2
    ModifyConfigFileNode $MSOXCFXICSDeploymentFile "User2Password"           $MSOXCFXICSUser2Password    
    ModifyConfigFileNode $MSOXCFXICSDeploymentFile "TestUser2ESSDN"          $MSOXCFXICSUser2Essdn
    ModifyConfigFileNode $MSOXCFXICSDeploymentFile "GhostedPublicFolderName" $MSOXCFXICSGhostedPublicFolder
}    

Output "Configuration for MS-OXCFXICS_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCMAPIHTTP ptfconfig file.
#-------------------------------------------------------
if($sutVersion -ge "ExchangeServer2013")
{
    Output "Configure MS-OXCMAPIHTTP_TestSuite.deployment.ptfconfig file ..." "White"
    $MSOXCMAPIHTTPAdminUser             = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPAdminUser"
    $MSOXCMAPIHTTPAdminUserPassword     = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPAdminUserPassword"
    $MSOXCMAPIHTTPGeneralUser           = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPGeneralUser"
    $MSOXCMAPIHTTPDistributionGroup     = ReadConfigFileNode "$environmentResourceFile" "MSOXCMAPIHTTPDistributionGroup"

    #Get the value for AmbiguousName
    $MSOXCMAPIHTTPAmbiguousName = GetFirstSameSubstring $MSOXCMAPIHTTPAdminUser $MSOXCMAPIHTTPGeneralUser
    if($MSOXCMAPIHTTPAmbiguousName -eq "")
    {
        Output "The values of the properties AdminUserName and GeneralUserName do not have the same prefix." "Yellow"
        Output "Would you like to continue executing client setup script?" "Cyan"   
        Output "1: CONTINUE(Which will set an empty value for property AmbiguousName. In this case, MS-OXCMAPIHTTP test cases may fail." "Cyan"
        Output "2: EXIT." "Cyan"    
        $runWhenMSOXCMAPIHTTPAmbiguousNameisEmptyChoices = @('1','2')
        $runWhenMSOXCMAPIHTTPAmbiguousNameisEmpty = ReadUserChoice $runWhenMSOXCMAPIHTTPAmbiguousNameisEmptyChoices "runWhenMSOXCMAPIHTTPAmbiguousNameisEmpty"
        if($runWhenMSOXCMAPIHTTPAmbiguousNameisEmpty -eq "2")
        {
            Stop-Transcript
            exit 0
        }       
    }

    # Get User DN for MS-OXCMAPIHTTP
    Output "Get Exchange user($MSOXCMAPIHTTPAdminUser)'s ESSDN automatically." "White"
    $MSOXCMAPIHTTPAdminUserEssdn = GetUserDN $sutComputerName $MSOXCMAPIHTTPAdminUser "$dnsDomain\$userName" $password
    Output "Get Exchange user($MSOXCMAPIHTTPGeneralUser)'s ESSDN automatically." "White"
    $MSOXCMAPIHTTPGeneralUserEssdn = GetUserDN $sutComputerName $MSOXCMAPIHTTPGeneralUser "$dnsDomain\$userName" $password

    $serverDN = $MSOXCMAPIHTTPAdminUserEssdn.Substring(0,$MSOXCMAPIHTTPAdminUserEssdn.ToLower().IndexOf("/cn=recipients")) + "/cn=Configuration/cn=Servers/cn=$sutComputerName@$dnsDomain"

    Output "Modify the properties as necessary in the MS-OXCMAPIHTTP_TestSuite.deployment.ptfconfig file..." "White"
    Output "Steps for manual configuration:" "Yellow"
    $step = 1
    Output "$step.Open $MSOXCMAPIHTTPDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCMAPIHTTPAdminUser" "Yellow"
    $step++
    Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCMAPIHTTPAdminUserPassword" "Yellow"
    $step++
    Output "$step.Find the property `"AdminUserEssdn`", and set the value as $MSOXCMAPIHTTPAdminUserEssdn" "Yellow"
    $step++
    Output "$step.Find the property `"GeneralUserName`", and set the value as $MSOXCMAPIHTTPGeneralUser" "Yellow"
    $step++
    Output "$step.Find the property `"GeneralUserEssdn`", and set the value as $MSOXCMAPIHTTPGeneralUserEssdn" "Yellow"
    $step++
    Output "$step.Find the property `"ServerDN`", and set the value as $serverDN" "Yellow"
    $step++
    Output "$step.Find the property `"DistributionListName`", and set the value as $MSOXCMAPIHTTPDistributionGroup" "Yellow"
    $step++
    Output "$step.Find the property `"AmbiguousName`", and set the value as $MSOXCMAPIHTTPAmbiguousName" "Yellow"
        
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "AdminUserName"        $MSOXCMAPIHTTPAdminUser
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "AdminUserPassword"    $MSOXCMAPIHTTPAdminUserPassword
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "AdminUserEssdn"       $MSOXCMAPIHTTPAdminUserEssdn
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "GeneralUserName"      $MSOXCMAPIHTTPGeneralUser
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "GeneralUserEssdn"     $MSOXCMAPIHTTPGeneralUserEssdn
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "ServerDN"             $serverDN
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "DistributionListName" $MSOXCMAPIHTTPDistributionGroup
    ModifyConfigFileNode $MSOXCMAPIHTTPDeploymentFile "AmbiguousName"        $MSOXCMAPIHTTPAmbiguousName

    Output "Configuration for MS-OXCMAPIHTTP_TestSuite.deployment.ptfconfig file is complete" "Green"
}

#-------------------------------------------------------
# Configuration for MS-OXCMSG ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCMSG_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCMSGCommonUser           = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGCommonUser"
$MSOXCMSGCommonUserPassword   = ReadConfigFileNode "$environmentResourceFile" "MSOXCMSGCommonUserPassword"
$MSOXCMSGAdminUser            = ReadConfigFileNode  "$environmentResourceFile" "MSOXCMSGAdminUser"
$MSOXCMSGAdminUserPassword    = ReadConfigFileNode  "$environmentResourceFile" "MSOXCMSGAdminUserPassword"
# Get User DN for MS-OXCMSG
Output "Get Exchange user($MSOXCMSGCommonUser)'s ESSDN automatically." "White"
$MSOXCMSGCommonUserEssdn = GetUserDN $sutComputerName $MSOXCMSGCommonUser "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCMSGAdminUser)'s ESSDN automatically." "White"
$MSOXCMSGAdminUserEssdn = GetUserDN $sutComputerName $MSOXCMSGAdminUser "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCMSG_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCMSGDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"CommonUser`", and set the value as $MSOXCMSGCommonUser" "Yellow"
$step++
Output "$step.Find the property `"CommonUserPassword`", and set the value as $MSOXCMSGCommonUserPassword" "Yellow"
$step++
Output "$step.Find the property `"CommonUserEssdn`", and set the value as $MSOXCMSGCommonUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCMSGAdminUser" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCMSGAdminUserPassword" "Yellow"
$step++
Output "$step.Find the property `"AdminUserEssdn`", and set the value as $MSOXCMSGAdminUserEssdn" "Yellow"
    
ModifyConfigFileNode $MSOXCMSGDeploymentFile "CommonUser"         $MSOXCMSGCommonUser
ModifyConfigFileNode $MSOXCMSGDeploymentFile "CommonUserPassword" $MSOXCMSGCommonUserPassword
ModifyConfigFileNode $MSOXCMSGDeploymentFile "CommonUserEssdn"    $MSOXCMSGCommonUserEssdn
ModifyConfigFileNode $MSOXCMSGDeploymentFile "AdminUserName"      $MSOXCMSGAdminUser
ModifyConfigFileNode $MSOXCMSGDeploymentFile "AdminUserPassword"  $MSOXCMSGAdminUserPassword
ModifyConfigFileNode $MSOXCMSGDeploymentFile "AdminUserEssdn"     $MSOXCMSGAdminUserEssdn

Output "Configuration for MS-OXCMSG_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCNOTIF ptfconfig file
#-------------------------------------------------------
Output "Configure MS-OXCNOTIF_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCNOTIFUser             = ReadConfigFileNode "$environmentResourceFile" "MSOXCNOTIFUser"
$MSOXCNOTIFUserPassword     = ReadConfigFileNode "$environmentResourceFile" "MSOXCNOTIFUserPassword"
$MSOXCNOTIFNotificationPort = ReadConfigFileNode "$environmentResourceFile" "MSOXCNOTIFNotificationPort"

AddFirewallInboundRule "MAPI MS-OXCNOTIF" "UDP" $MSOXCNOTIFNotificationPort $true

# Get User DN for MS-OXCNOTIF
Output "Get Exchange user($MSOXCNOTIFUser)'s ESSDN automatically." "White"
$MSOXCNOTIFUserEssdn = GetUserDN $sutComputerName $MSOXCNOTIFUser "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCNOTIF_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCNOTIFDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSOXCNOTIFUser" "Yellow" 
$step++
Output "$step.Find the property `"User1Password`", and set the value as $MSOXCNOTIFUserPassword" "Yellow"
$step++
Output "$step.Find the property `"User1Essdn`", and set the value as $MSOXCNOTIFUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"NotificationPort`", and set the value as $MSOXCNOTIFNotificationPort" "Yellow"   

ModifyConfigFileNode $MSOXCNOTIFDeploymentFile "User1Name"         $MSOXCNOTIFUser
ModifyConfigFileNode $MSOXCNOTIFDeploymentFile "User1Password"     $MSOXCNOTIFUserPassword
ModifyConfigFileNode $MSOXCNOTIFDeploymentFile "User1Essdn"        $MSOXCNOTIFUserEssdn
ModifyConfigFileNode $MSOXCNOTIFDeploymentFile "NotificationPort"  $MSOXCNOTIFNotificationPort

Output "Configuration for MS-OXCNOTIF_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCPERM ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCPERM_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCPERMUser1                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser1"
$MSOXCPERMUser1Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser1Password"
$MSOXCPERMUser2                     = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser2"
$MSOXCPERMUser2Password             = ReadConfigFileNode "$environmentResourceFile" "MSOXCPERMUser2Password"
# Get User DN for MS-OXCPERM
Output "Get Exchange user($MSOXCPERMUser1)'s ESSDN automatically." "White"
$MSOXCPERMUser1Essdn = GetUserDN $sutComputerName $MSOXCPERMUser1 "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCPERMUser2)'s ESSDN automatically." "White"
$MSOXCPERMUser2Essdn = GetUserDN $sutComputerName $MSOXCPERMUser2 "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCPERM_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCPERMDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSOXCPERMUser1" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $MSOXCPERMUser1Password" "Yellow"
$step++
Output "$step.Find the property `"User1Essdn`", and set the value as $MSOXCPERMUser1Essdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCPERMUser2" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCPERMUser1Essdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserEssdn`", and set the value as $MSOXCPERMUser2Essdn" "Yellow"
    
ModifyConfigFileNode $MSOXCPERMDeploymentFile "User1Name"      $MSOXCPERMUser1
ModifyConfigFileNode $MSOXCPERMDeploymentFile "User1Password"  $MSOXCPERMUser1Password
ModifyConfigFileNode $MSOXCPERMDeploymentFile "User1Essdn"     $MSOXCPERMUser1Essdn
ModifyConfigFileNode $MSOXCPERMDeploymentFile "AdminUserName"      $MSOXCPERMUser2
ModifyConfigFileNode $MSOXCPERMDeploymentFile "AdminUserPassword"  $MSOXCPERMUser2Password
ModifyConfigFileNode $MSOXCPERMDeploymentFile "AdminUserEssdn"     $MSOXCPERMUser2Essdn

Output "Configuration for MS-OXCPERM_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCPRPT ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCPRPT_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCPRPTUser         = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTUser"
$MSOXCPRPTUserPassword = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTUserPassword"
$MSOXCPRPTPublicFolder = ReadConfigFileNode "$environmentResourceFile" "MSOXCPRPTPublicFolder"
if($sutVersion -ge "ExchangeServer2013")
{
     $logonPropertyID0 = "26141"
     $folderPropertyID2 = "16355"
}
else
{
     $logonPropertyID0 = "3585"
     $folderPropertyID2 = "16353"
}

# Get User DN for MS-OXCPRPT
Output "Get Exchange user($MSOXCPRPTUser)'s ESSDN automatically." "White"
$MSOXCPRPTUserEssdn = GetUserDN $sutComputerName $MSOXCPRPTUser "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCPRPT_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCPRPTDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"UserName`", and set the value as $MSOXCPRPTUser" "Yellow"
$step++
Output "$step.Find the property `"Password`", and set the value as $MSOXCPRPTUserPassword" "Yellow"
$step++
Output "$step.Find the property `"UserEssdn`", and set the value as $MSOXCPRPTUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"LogonPropertyID0`", and set the value as $logonPropertyID0" "Yellow"
$step++
Output "$step.Find the property `"FolderPropertyID2`", and set the value as $folderPropertyID2" "Yellow"
$step++
Output "$step.Find the property `"PublicFolderName`", and set the value as $MSOXCPRPTPublicFolder" "Yellow" 
   
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "UserName"           $MSOXCPRPTUser
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "Password"           $MSOXCPRPTUserPassword
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "UserEssdn"          $MSOXCPRPTUserEssdn
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "LogonPropertyID0"   $logonPropertyID0
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "FolderPropertyID2"  $folderPropertyID2
ModifyConfigFileNode $MSOXCPRPTDeploymentFile "PublicFolderName"   $MSOXCPRPTPublicFolder

Output "Configuration for MS-OXCPRPT_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCROPS ptfconfig file.
#-------------------------------------------------------
Output "Configure MS-OXCROPS_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCROPSUser                = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSUser"
$MSOXCROPSUserPassword        = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSUserPassword"
$MSOXCROPSEmailAlias          = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSEmailAlias"
$MSOXCROPSEmailAliasPassword  = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSEmailAliasPassword"

# Get User DN for MS-OXCROPS
Output "Get Exchange user($MSOXCROPSUser)'s ESSDN automatically." "White"
$MSOXCROPSUserEssdn = GetUserDN $sutComputerName $MSOXCROPSUser "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCROPSEmailAlias)'s ESSDN automatically." "White"
$MSOXCROPSEmailAliasEssdn = GetUserDN $sutComputerName $MSOXCROPSEmailAlias "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCROPS_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCROPSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCROPSUser" "Yellow"
$step++
Output "$step.Find the property `"Password`", and set the value as $MSOXCROPSUserPassword" "Yellow"
$step++
Output "$step.Find the property `"UserEssdn`", and set the value as $MSOXCROPSUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"EmailAlias`", and set the value as $MSOXCROPSEmailAlias" "Yellow"
$step++
Output "$step.Find the property `"EmailAliasPassword`", and set the value as $MSOXCROPSEmailAliasPassword" "Yellow"
$step++
Output "$step.Find the property `"EmailAliasEssdn`", and set the value as $MSOXCROPSEmailAliasEssdn" "Yellow"
if($sut2ComputerName -ne "")
{
    $MSOXCROPSPublicFolderGhosted = ReadConfigFileNode "$environmentResourceFile" "MSOXCROPSPublicFolderGhosted"
    $step++
    Output "$step.Find the property `"GhostedPublicFolderDisplayName`", and set the value as $MSOXCROPSPublicFolderGhosted" "Yellow"
}

ModifyConfigFileNode $MSOXCROPSDeploymentFile "AdminUserName"   $MSOXCROPSUser
ModifyConfigFileNode $MSOXCROPSDeploymentFile "Password"   $MSOXCROPSUserPassword
ModifyConfigFileNode $MSOXCROPSDeploymentFile "UserEssdn"  $MSOXCROPSUserEssdn
ModifyConfigFileNode $MSOXCROPSDeploymentFile "EmailAlias" $MSOXCROPSEmailAlias
ModifyConfigFileNode $MSOXCROPSDeploymentFile "EmailAliasPassword" $MSOXCROPSEmailAliasPassword
ModifyConfigFileNode $MSOXCROPSDeploymentFile "EmailAliasEssdn"    $MSOXCROPSEmailAliasEssdn
if($sut2ComputerName -ne "")
{
    ModifyConfigFileNode $MSOXCROPSDeploymentFile "GhostedPublicFolderDisplayName"  $MSOXCROPSPublicFolderGhosted
}

Output "Configuration for MS-OXCROPS_TestSuite.deployment.ptfconfig file is complete" "Green"

#-------------------------------------------------------
# Configuration for MS-OXCRPC ptfconfig file
#-------------------------------------------------------
Output "Configure MS-OXCRPC_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCRPCNormalUser            = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCNormalUser"
$MSOXCRPCNormalUserPassword    = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCNormalUserPassword"
$MSOXCRPCAdminUser             = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCAdminUser"
$MSOXCRPCAdminUserPassword     = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCAdminUserPassword"
$MSOXCRPCNotificationPort      = ReadConfigFileNode "$environmentResourceFile" "MSOXCRPCNotificationPort"

AddFirewallInboundRule "MAPI MS-OXCRPC" "UDP" $MSOXCRPCNotificationPort $true

# Get User DN for MS-OXCRPC
Output "Get Exchange user($MSOXCRPCNormalUser)'s ESSDN automatically." "White"
$MSOXCRPCNormalUserEssdn = GetUserDN $sutComputerName $MSOXCRPCNormalUser "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCRPCAdminUser)'s ESSDN automatically." "White"
$MSOXCRPCAdminUserEssdn = GetUserDN $sutComputerName $MSOXCRPCAdminUser "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCRPC_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCRPCDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"NormalUserName`", and set the value as $MSOXCRPCNormalUser" "Yellow"
$step++
Output "$step.Find the property `"NormalUserPassword`", and set the value as $MSOXCRPCNormalUserPassword" "Yellow"
$step++
Output "$step.Find the property `"NormalUserEssdn`", and set the value as $MSOXCRPCNormalUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCRPCAdminUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXCRPCNormalUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"AdminUserEssdn`", and set the value as $MSOXCRPCAdminUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"NotificationPort`", and set the value as $MSOXCRPCNotificationPort" "Yellow"  

ModifyConfigFileNode $MSOXCRPCDeploymentFile "NormalUserName"               $MSOXCRPCNormalUser
ModifyConfigFileNode $MSOXCRPCDeploymentFile "NormalUserPassword"           $MSOXCRPCNormalUserPassword
ModifyConfigFileNode $MSOXCRPCDeploymentFile "NormalUserEssdn"              $MSOXCRPCNormalUserEssdn
ModifyConfigFileNode $MSOXCRPCDeploymentFile "AdminUserName"                $MSOXCRPCAdminUser
ModifyConfigFileNode $MSOXCRPCDeploymentFile "AdminUserPassword"            $MSOXCRPCAdminUserPassword
ModifyConfigFileNode $MSOXCRPCDeploymentFile "AdminUserEssdn"               $MSOXCRPCAdminUserEssdn
ModifyConfigFileNode $MSOXCRPCDeploymentFile "NotificationPort"             $MSOXCRPCNotificationPort

Output "Configuration for MS-OXCRPC_TestSuite.deployment.ptfconfig file is complete" "Green"

#-----------------------------------------------------
# Configuration for MS-OXCSTOR ptfconfig file
#-----------------------------------------------------
Output "Configure MS-OXCSTOR_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCSTORUser                      = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORUser"
$MSOXCSTORUserPassword              = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORUserPassword"
$MSOXCSTORMailboxOnServer1          = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer1"
$MSOXCSTORMailboxOnServer1Password  = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer1Password"
$MSOXCSTORDisableMailbox            = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORDisableMailbox"
$MSOXCSTORDisableMailboxPassword    = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORDisableMailboxPassword"
# Get User DN for MS-OXCSTOR
Output "Get Exchange user($MSOXCSTORUser)'s ESSDN automatically." "White"
$MSOXCSTORUserEssdn = GetUserDN $sutComputerName $MSOXCSTORUser "$dnsDomain\$userName" $password
Output "Get Exchange user($MSOXCSTORMailboxOnServer1)'s ESSDN automatically." "White"
$MSOXCSTORMailboxOnServer1Essdn = GetUserDN $sutComputerName $MSOXCSTORMailboxOnServer1 "$dnsDomain\$userName" $password
if($sut2ComputerName -ne "")
{
    $MSOXCSTORMailboxOnServer2          = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer2"
    $MSOXCSTORMailboxOnServer2Password  = ReadConfigFileNode "$environmentResourceFile" "MSOXCSTORMailboxOnServer2Password"
    Output "Get Exchange user($MSOXCSTORMailboxOnServer2)'s ESSDN automatically." "White"
    $MSOXCSTORMailboxOnServer2Essdn = GetUserDN $sutComputerName $MSOXCSTORMailboxOnServer2 "$dnsDomain\$userName" $password
}

Output "Modify the properties as necessary in the MS-OXCSTOR_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCSTORDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXCSTORUser" "Yellow"
$step++
Output "$step.Find the property `"UserPassword`", and set the value as $MSOXCSTORUserPassword" "Yellow"
$step++
Output "$step.Find the property `"UserEssdn`", and set the value as $MSOXCSTORUserEssdn" "Yellow"
$step++
Output "$step.Find the property `"UserNameOfMailboxOnServer1`", and set the value as $MSOXCSTORMailboxOnServer1" "Yellow"
$step++
Output "$step.Find the property `"UserPasswordOfMailboxOnServer1`", and set the value as $MSOXCSTORMailboxOnServer1Password" "Yellow"
$step++
Output "$step.Find the property `"UserEssdnOfMailboxOnServer1`", and set the value as $MSOXCSTORMailboxOnServer1Essdn" "Yellow"
$step++
Output "$step.Find the property `"UserNameForDisableMailbox`", and set the value as $MSOXCSTORDisableMailbox" "Yellow"
$step++
Output "$step.Find the property `"UserPasswordForDisableMailbox`", and set the value as $MSOXCSTORDisableMailboxPassword" "Yellow"

if($sut2ComputerName -ne "")
{
    $step++
    Output "$step.Find the property `"UserNameOfMailboxOnServer2`", and set the value as $MSOXCSTORMailboxOnServer2" "Yellow"
    $step++
    Output "$step.Find the property `"UserPasswordOfMailboxOnServer2`", and set the value as $MSOXCSTORMailboxOnServer2Password" "Yellow"
    $step++
    Output "$step.Find the property `"UserEssdnOfMailboxOnServer2`", and set the value as $MSOXCSTORMailboxOnServer2Essdn" "Yellow"
}

ModifyConfigFileNode $MSOXCSTORDeploymentFile "AdminUserName"                     $MSOXCSTORUser
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserPassword"                      $MSOXCSTORUserPassword
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserEssdn"                         $MSOXCSTORUserEssdn
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserNameOfMailboxOnServer1"        $MSOXCSTORMailboxOnServer1
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserPasswordOfMailboxOnServer1"    $MSOXCSTORMailboxOnServer1Password
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserEssdnOfMailboxOnServer1"       $MSOXCSTORMailboxOnServer1Essdn
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserNameForDisableMailbox"         $MSOXCSTORDisableMailbox
ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserPasswordForDisableMailbox"     $MSOXCSTORDisableMailboxPassword

if($sut2ComputerName -ne "")
{
    ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserNameOfMailboxOnServer2"      $MSOXCSTORMailboxOnServer2
    ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserPasswordOfMailboxOnServer2"  $MSOXCSTORMailboxOnServer2Password
    ModifyConfigFileNode $MSOXCSTORDeploymentFile "UserEssdnOfMailboxOnServer2"     $MSOXCSTORMailboxOnServer2Essdn
}

Output "Configuration for MS-OXCSTOR_TestSuite.deployment.ptfconfig file is complete" "Green"

#-----------------------------------------------------
# Configuration for MS-OXCTABL ptfconfig file
#-----------------------------------------------------
Output "Configure MS-OXCTABL_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXCTABLSender1           = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender1"
$MSOXCTABLSender1Password   = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender1Password"
$MSOXCTABLSender2           = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender2"
$MSOXCTABLSender2Password   = ReadConfigFileNode "$environmentResourceFile" "MSOXCTABLSender2Password"
# Get User DN for MS-OXCTABL
Output "Get Exchange user($MSOXCTABLSender1)'s ESSDN automatically." "White"
$MSOXCTABLSender1Essdn = GetUserDN $sutComputerName $MSOXCTABLSender1 "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXCTABL_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXCTABLDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"Sender1Name`", and set the value as $MSOXCTABLSender1" "Yellow"
$step++
Output "$step.Find the property `"Sender1Password`", and set the value as $MSOXCTABLSender1Password" "Yellow"
$step++
Output "$step.Find the property `"Sender1Essdn`", and set the value as $MSOXCTABLSender1Essdn" "Yellow"
$step++
Output "$step.Find the property `"Sender2Name`", and set the value as $MSOXCTABLSender2" "Yellow"

ModifyConfigFileNode $MSOXCTABLDeploymentFile "Sender1Name"          $MSOXCTABLSender1
ModifyConfigFileNode $MSOXCTABLDeploymentFile "Sender1Password"      $MSOXCTABLSender1Password
ModifyConfigFileNode $MSOXCTABLDeploymentFile "Sender1Essdn"         $MSOXCTABLSender1Essdn
ModifyConfigFileNode $MSOXCTABLDeploymentFile "Sender2Name"          $MSOXCTABLSender2

Output "Configuration for MS-OXCTABL_TestSuite.deployment.ptfconfig file is complete" "Green"

#-----------------------------------------------------
# Configuration for MS-OXNSPI ptfconfig file
#-----------------------------------------------------
if($sutVersion -ge "ExchangeServer2010")
{
    Output "Configure MS-OXNSPI_TestSuite.deployment.ptfconfig file ..." "White"
    $MSOXNSPIUser1                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser1"
    $MSOXNSPIUser1Password              = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser1Password"
    $MSOXNSPIUser2                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser2"
    $MSOXNSPIUser3                      = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIUser3"
    $MSOXNSPIPublicFolderMailEnabled    = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIPublicFolderMailEnabled"
    $MSOXNSPIDynamicDistributionGroup   = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIDynamicDistributionGroup"
    $MSOXNSPIDistributionGroup          = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIDistributionGroup"
    $MSOXNSPIMailContact                = ReadConfigFileNode "$environmentResourceFile" "MSOXNSPIMailContact"

    #Get the value for AmbiguousName
    $MSOXNSPIAmbiguousName = GetFirstSameSubstring $MSOXNSPIUser1 $MSOXNSPIUser2
    if($MSOXNSPIAmbiguousName -eq "" -or $MSOXNSPIUser1.StartsWith($MSOXNSPIAmbiguousName) -ne $true -or $MSOXNSPIUser2.StartsWith($MSOXNSPIAmbiguousName) -ne $true )
    {
        $MSOXNSPIAmbiguousName = ""
        Output "The values of the properties User1Name and User2Name do not have the same prefix." "Yellow"
        Output "Would you like to continue executing client setup script?" "Cyan"   
        Output "1: CONTINUE(Which will set an empty value for property AmbiguousName. In this case, MS-OXNSPI test cases may fail." "Cyan"
        Output "2: EXIT." "Cyan"    
        $runWhenMSOXNSPIAmbiguousNameisEmptyChoices = @('1','2')
        $runWhenMSOXNSPIAmbiguousNameisEmpty = ReadUserChoice $runWhenMSOXNSPIAmbiguousNameisEmptyChoices "runWhenMSOXNSPIAmbiguousNameisEmpty"
        if($runWhenMSOXNSPIAmbiguousNameisEmpty -eq "2")
        {
            Stop-Transcript
            exit 0
        }       
    }

    # Get User DN for MS-OXNSPI
    Output "Get Exchange user($MSOXNSPIUser1)'s ESSDN automatically." "White"
    $MSOXNSPIUser1Essdn = GetUserDN $sutComputerName $MSOXNSPIUser1 "$dnsDomain\$userName" $password
    Output "Get Exchange user($MSOXNSPIUser2)'s ESSDN automatically." "White"
    $MSOXNSPIUser2Essdn = GetUserDN $sutComputerName $MSOXNSPIUser2 "$dnsDomain\$userName" $password
    Output "Get Exchange user($MSOXNSPIUser3)'s ESSDN automatically." "White"
    $MSOXNSPIUser3Essdn = GetUserDN $sutComputerName $MSOXNSPIUser3 "$dnsDomain\$userName" $password

    Output "Modify the properties as necessary in the MS-OXNSPI_TestSuite.deployment.ptfconfig file..." "White"
    Output "Steps for manual configuration:" "Yellow"
    $step = 1
    Output "$step.Open $MSOXNSPIDeploymentFile" "Yellow"
    $step++
    Output "$step.Find the property `"User1Name`", and set the value as $MSOXNSPIUser1" "Yellow"
    $step++
    Output "$step.Find the property `"User1Password`", and set the value as $MSOXNSPIUser1Password" "Yellow"
    $step++
    Output "$step.Find the property `"User1Essdn`", and set the value as $MSOXNSPIUser1Essdn" "Yellow"
    $step++
    Output "$step.Find the property `"User2Name`", and set the value as $MSOXNSPIUser2" "Yellow"
    $step++
    Output "$step.Find the property `"User2Essdn`", and set the value as $MSOXNSPIUser2Essdn" "Yellow"
    $step++
    Output "$step.Find the property `"User3Name`", and set the value as $MSOXNSPIUser3" "Yellow"
    $step++
    Output "$step.Find the property `"User3NameEssdn`", and set the value as $MSOXNSPIUser3Essdn" "Yellow"
    $step++
    Output "$step.Find the property `"DistributionListName`", and set the value as $MSOXNSPIDistributionGroup" "Yellow"
    $step++
    Output "$step.Find the property `"AgentName`", and set the value as $MSOXNSPIDynamicDistributionGroup" "Yellow"
    $step++
    Output "$step.Find the property `"ForumName`", and set the value as $MSOXNSPIPublicFolderMailEnabled" "Yellow"
    $step++
    Output "$step.Find the property `"RemoteMailUserName`", and set the value as $MSOXNSPIMailContact" "Yellow"
    $step++
    Output "$step.Find the property `"AmbiguousName`", and set the value as $MSOXNSPIAmbiguousName" "Yellow"

    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User1Name"                 $MSOXNSPIUser1
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User1Password"             $MSOXNSPIUser1Password
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User1Essdn"                $MSOXNSPIUser1Essdn
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User2Name"                 $MSOXNSPIUser2
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User2Essdn"                $MSOXNSPIUser2Essdn
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User3Name"                 $MSOXNSPIUser3
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "User3Essdn"                $MSOXNSPIUser3Essdn
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "DistributionListName"      $MSOXNSPIDistributionGroup
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "AgentName"                 $MSOXNSPIDynamicDistributionGroup
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "ForumName"                 $MSOXNSPIPublicFolderMailEnabled
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "RemoteMailUserName"        $MSOXNSPIMailContact
    ModifyConfigFileNode $MSOXNSPIDeploymentFile "AmbiguousName"             $MSOXNSPIAmbiguousName

    Output "Configuration for MS-OXNSPI_TestSuite.deployment.ptfconfig file is complete" "Green"
}

#-----------------------------------------------------
# Configuration for MS-OXORULE ptfconfig file
#-----------------------------------------------------
Output "Configure MS-OXORULE_TestSuite.deployment.ptfconfig file ..." "White"
$MSOXORULEUser1           = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser1"
$MSOXORULEUser1Password   = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser1Password"
$MSOXORULEUser2           = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser2"
$MSOXORULEUser2Password   = ReadConfigFileNode "$environmentResourceFile" "MSOXORULEUser2Password"

# Get User1DN for MS-OXORULE
Output "Get Exchange user($MSOXORULEUser1)'s ESSDN automatically." "White"
$MSOXORULEUser1Essdn = GetUserDN $sutComputerName $MSOXORULEUser1 "$dnsDomain\$userName" $password

# Get User2DN for MS-OXORULE
Output "Get Exchange user($MSOXORULEUser2)'s ESSDN automatically." "White"
$MSOXORULEUser2Essdn = GetUserDN $sutComputerName $MSOXORULEUser2 "$dnsDomain\$userName" $password

Output "Modify the properties as necessary in the MS-OXORULE_TestSuite.deployment.ptfconfig file..." "White"
Output "Steps for manual configuration:" "Yellow"
$step = 1
Output "$step.Open $MSOXORULEDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"AdminUserName`", and set the value as $MSOXORULEUser1" "Yellow"
$step++
Output "$step.Find the property `"AdminUserPassword`", and set the value as $MSOXORULEUser1Password" "Yellow"
$step++
Output "$step.Find the property `"AdminUserESSDN`", and set the value as $MSOXORULEUser1Essdn" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSOXORULEUser2" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $MSOXORULEUser2Password" "Yellow"
$step++
Output "$step.Find the property `"User2ESSDN`", and set the value as $MSOXORULEUser2Essdn" "Yellow"

ModifyConfigFileNode $MSOXORULEDeploymentFile "AdminUserName"              $MSOXORULEUser1
ModifyConfigFileNode $MSOXORULEDeploymentFile "AdminUserPassword"          $MSOXORULEUser1Password
ModifyConfigFileNode $MSOXORULEDeploymentFile "AdminUserESSDN"             $MSOXORULEUser1Essdn
ModifyConfigFileNode $MSOXORULEDeploymentFile "User2Name"              $MSOXORULEUser2
ModifyConfigFileNode $MSOXORULEDeploymentFile "User2Password"          $MSOXORULEUser2Password
ModifyConfigFileNode $MSOXORULEDeploymentFile "User2ESSDN"             $MSOXORULEUser2Essdn

Output "Configuration for MS-OXORULE_TestSuite.deployment.ptfconfig file is complete" "Green"

#----------------------------------------------------------------------------
# Import certificates
#----------------------------------------------------------------------------
Output "Find the *.cer certificate files located on the SUT server's system root directory, and import to the local computer's Trusted Root Certification Authorities store." "White"
ImportCertificates $sutComputerName "$dnsDomain\$userName" $password
if($sut2ComputerName -ne "")
{
    ImportCertificates $sut2ComputerName "$dnsDomain\$userName" $password
}

#----------------------------------------------------------------------------
# End script
#----------------------------------------------------------------------------
Output "[ExchangeClientConfiguration.ps1] has run sucessfully." "Green"
AddTimesStampsToLogFile "End" "$logFile"
Stop-Transcript
exit 0