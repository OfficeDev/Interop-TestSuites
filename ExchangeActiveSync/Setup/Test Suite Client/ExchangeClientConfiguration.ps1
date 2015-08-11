#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

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
Start-Transcript $debugLogFile -force -append

#----------------------------------------------------------------------------
# Default Variables for Configuration 
#----------------------------------------------------------------------------
$userPassword                        = "Password01!"

$MSASAIRSUser01                      = "MSASAIRS_User01"
$MSASAIRSUser02                      = "MSASAIRS_User02"

$MSASCALUser01                       = "MSASCAL_User01"
$MSASCALUser02                       = "MSASCAL_User02"

$MSASCMDUser01                       = "MSASCMD_User01"
$MSASCMDUser02                       = "MSASCMD_User02"
$MSASCMDUser03                       = "MSASCMD_User03"
$MSASCMDUser07                       = "MSASCMD_User07"
$MSASCMDUser08                       = "MSASCMD_User08"
$MSASCMDUser09                       = "MSASCMD_User09"
$MSASCMDTestGroup                    = "MSASCMD_TestGroup"
$MSASCMDLargeGroup                   = "MSASCMD_LargeGroup"
$MSASCMDSharedFolder                 = "MSASCMD_SharedFolder"
$MSASCMDNonEmptyDocument             = "MSASCMD_Non-emptyDocument.txt"
$MSASCMDEmptyDocument                = "MSASCMD_EmptyDocument.txt"
$MSASCMDEmailSubjectName             = "MSASCMD_SecureEmailForTest"

$MSASCNTCUser01                      = "MSASCNTC_User01"
$MSASCNTCUser02                      = "MSASCNTC_User02"

$MSASCONUser01                       = "MSASCON_User01"
$MSASCONUser02                       = "MSASCON_User02"
$MSASCONUser03                       = "MSASCON_User03"

$MSASDOCUser01                       = "MSASDOC_User01"
$MSASDOCSharedFolder                 = "MSASDOC_SharedFolder"
$MSASDOCVisibleFolder                = "MSASDOC_VisibleFolder"
$MSASDOCHiddenFolder                 = "MSASDOC_HiddenFolder"
$MSASDOCVisibleDocument              = "MSASDOC_VisibleDocument.txt"
$MSASDOCHiddenDocument               = "MSASDOC_HiddenDocument.txt"

$MSASEMAILUser01                     = "MSASEMAIL_User01"
$MSASEMAILUser02                     = "MSASEMAIL_User02"
$MSASEMAILUser03                     = "MSASEMAIL_User03"

$MSASHTTPUser01                      = "MSASHTTP_User01"
$MSASHTTPUser02                      = "MSASHTTP_User02"
$MSASHTTPUser03                      = "MSASHTTP_User03"
$MSASHTTPUser04                      = "MSASHTTP_User04"

$MSASNOTEUser01                      = "MSASNOTE_User01"

$MSASPROVUser01                      = "MSASPROV_User01"
$MSASPROVUser02                      = "MSASPROV_User02"
$MSASPROVUser03                      = "MSASPROV_User03"

$MSASRMUser01                        = "MSASRM_User01"
$MSASRMUser02                        = "MSASRM_User02"
$MSASRMUser03                        = "MSASRM_User03"
$MSASRMUser04                        = "MSASRM_User04"

$MSASTASKUser01                      = "MSASTASK_User01"

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

#-----------------------------------------------------------------------------------
# <summary>
# Throw expection if the required parameter is empty.
# </summary>
# <param name="parameterName">The name of parameter to be checked.</param>
# <param name="parameterValue">The value of parameter to be checked.</param>
#-----------------------------------------------------------------------------------
function ValidateParameter
{
    param(
    [string]$parameterName,
    [string]$parameterValue
    )
    
    if ($parameterValue -eq $null -or $parameterValue -eq "")
    {
        Throw "Parameter $parameterName cannot be empty."
    }    
}

#-----------------------------------------------------------------------------------
# <summary>
# Print the content with the specified color and add the content to the log file. 
# </summary>
# <param name="content">The content to be printed.</param>
# <param name="color">The color type of the content.</param>
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
# Modify the value of the specified node in the specified ptfconfig file.
# </summary>
# <param name="sourceFileName">The name of the configuration file containing the node.</param>
# <param name="nodeName">The name of the node.</param>
# <param name="nodeValue">The new value of the node.</param>
#-----------------------------------------------------------------------------------
function ModifyConfigFileNode
{
    Param(
    [string]$sourceFileName, 
    [string]$nodeName, 
    [string]$nodeValue
    )

    #----------------------------------------------------------------------------
    # Verify required parameters
    #----------------------------------------------------------------------------
    ValidateParameter 'sourceFileName' $sourceFileName
    ValidateParameter 'nodeName' $nodeName

    #----------------------------------------------------------------------------
    # Modify the content of the node
    #----------------------------------------------------------------------------
    $isFileAvailable = $false
    $isNodeFound = $false

    $isFileAvailable = Test-Path $sourceFileName
    if($isFileAvailable -eq $true)
    {    
        [xml]$configContent = New-Object XML
        $configContent.Load($sourceFileName)
        $propertyNodes = $configContent.GetElementsByTagName("Property")
        foreach($node in $propertyNodes)
        {
            if($node.GetAttribute("name") -eq $nodeName)
            {
                $node.SetAttribute("value",$nodeValue)
                $isNodeFound = $true
                break
            }
        }
        
        if($isNodeFound)
        {
            $configContent.save($sourceFileName)
        }
        else
        {
            Throw "Failed while changing configuration file $sourceFileName : Could not find node with name attribute $nodeName." 
        }
    }
    else
    {
        Throw "Failed while changing configuration file $sourceFileName : File does not exist!" 
    }

    #----------------------------------------------------------------------------
    # Verify the result
    #----------------------------------------------------------------------------
    if($isFileAvailable -eq $true -and $isNodeFound)
    {
        [xml]$configContent = New-Object XML
        $configContent.Load($sourceFileName)
        $propertyNodes = $configContent.GetElementsByTagName("Property")
        foreach($node in $propertyNodes)
        {
            if($node.GetAttribute("name") -eq $nodeName)
            {
                if($node.GetAttribute("value") -eq $nodeValue)
                {
                    Output "Configuration is successful : $nodeName = $nodeValue" "Green"
                    return
                }
            }
        }

        Throw "Failed after changing the configuration file $sourceFileName : The actual value of node is not the same as the updated content value."
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Get user input by manually input or by reading unattended configuration XML.
# </summary>
# <param name="nodeName">Propery name in unattended configuration XML.</param>
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
                Output "$userInput (Received from the ExchangeClientConfigurationAnswers.xml file for property : $nodeName)." "White"
                break
            }
        }        
        if(!$isNodeFound)
        {
            Output "Could not find node with name attribute $nodeName in $unattendedXmlName, will use empty value instead." "Yellow"
        }
    }
    return $userInput
}

#-----------------------------------------------------------------------------------
# <summary>
# Read computer name from user's input. 
# <param name="nodeName">Propery name in unattended configuration XML.</param> 
# <returns>
# The valid computer name.
# </returns>
#-----------------------------------------------------------------------------------
function ReadComputerName
{
    param(
    [string]$nodeName
    )
    While(1)
    {
        [String]$computerName = GetUserInput $nodeName
        if($computerName -as [ipaddress])
        {
            Output "IP addresses are not supported." "Yellow"
        }
        elseif ($computerName -imatch '[`~!@#$%^&*()=+_\[\]{}\\|;:.''",<>/?]')
        {
            Output """$computerName"" contains characters that are not allowed, such as `` ~ ! @ # $ % ^ & * ( ) = + _ [ ] { } \ | ; : . ' "" , < > / and ?." "Yellow"
        }
        elseif ($computerName.Length -lt 1 -or $computerName.Length -gt 15)
        {
            Output "Computer name length must be between 1-15 characters." "Yellow"
        }
        else
        {
            return $computerName
        }
        if($unattendedXmlName -eq "")
        {
            Output "Retry with a valid computer name." "Yellow"
        }
        else
        {
            Write-Warning "Change the value of $nodeName with a valid computer name in the client configuration XML and then run the script again."
            Stop-Transcript
            exit 2
        }
    }    
}

#-----------------------------------------------------------------------------------
# <summary>
# Read user's input until it is a valid one. 
# </summary>
# <param name="userChoices">The available number list user can select from.</param> 
# <param name="nodeName">Propery name in unattended configuration XML.</param>
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
                Write-Warning "Change the value of $nodeName in the client configuration XML with the values listed and then run the script again."
                Stop-Transcript
                exit 2
            }
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check user input that should not be empty. 
# </summary>
# <param name="property">Property name that requires user to input the value.</param>
# <param name="nodeName">Propery name in unattended configuration XML.</param>
# <returns>
# Valid user input or exit script in unattended mode if the value provide in configuration XML is empty.
# </returns>
#-----------------------------------------------------------------------------------
function CheckForEmptyUserInput
{
    param(
    [string]$property,
    [string]$nodeName
    )
    While(1)
    {
        [String]$userInput = GetUserInput $nodeName
        if($userInput -ne "" -and $userInput -ne $null )
        {
            return $userInput
        }
        else
        {
            Output """$property"" can not be empty" "Yellow"
            if($unattendedXmlName -eq "")
            {
                Output "Retry with a non-empty one." "Yellow"
            }
            else
            {
                Write-Warning "Change the value of $nodeName in the client configuration XML and then run the script again."
                Stop-Transcript
                exit 2
            }
        }
     }    
}

#-----------------------------------------------------------------------------------
# <summary>
# Check if the operating system (OS) version of the specified computer is the recommended Windows 7 SP1/Windows 2008 R2 SP1 and above.
# </summary>
# <param name="computer">The computer name of the machine to be checked.</param>
#-----------------------------------------------------------------------------------
Function CheckOSVersion($computer)
{
    $os = Get-WmiObject -class Win32_OperatingSystem -computerName $computer
    if([int]$os.BuildNumber -ge 7601) 
    {
        Output "You are using the recommended operating system." "Green"
    }
    else
    {
        Output "Your operating system is not the recommended version, the recommended operating system is Windows 7 SP1/Windows 2008 R2 SP1 and above." "Yellow"
        Output "Would you like to continue to run the test suite on this machine or exit?" "Cyan"    
        Output "1: CONTINUE." "Cyan"    
        Output "2: EXIT." "Cyan"    
        $runOnNonRecommendedOSChoices = @('1','2')
        $runOnNonRecommendedOS = ReadUserChoice $runOnNonRecommendedOSChoices "runOnNonRecommendedOS"
        if($runOnNonRecommendedOS -eq "2")
        {
            Stop-Transcript
            exit 0
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Get installation path
# </summary>
# <param name="registryPaths">The registry path of software.</param>
# <returns>Return installation path if it is available, else return null.</returns>
#-----------------------------------------------------------------------------------
function GetInstalledPath
{
    param(
    [Array]$registryPaths
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'registryPaths' $registryPaths
    $installPath=$null
    foreach($registryPath in $registryPaths)
    {   if(Test-path $registryPath)
        {
            $application     = Get-Item $registryPath
            $psPath          = Get-ItemProperty $application.PsPath
            if($psPath -ne $null -and $psPath -ne "")
            {
                $installPath  = $psPath.InstallDir
                if($installPath -ne "" -and $installPath -ne $null -and (Test-Path $installPath))
                {
                    return $installPath 
                }
            }
        }
    }
    return $installPath
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the machine has installed the Visual Studio
# </summary>
# <param name="recommendVersion">The recommend version of Visual Studio.</param>
# <returns>A Boolean value, true if the Visual Studio has been installed, otherwise false.</returns>
#-----------------------------------------------------------------------------------
Function CheckVSVersion
{
    param(
    [string]$recommendVersion = "10.0"
    )
           
    $versions = @{"9.0" = "Visual Studio 2008";"10.0" = "Visual Studio 2010";"11.0" = "Visual Studio 2012";"12.0" = "Visual Studio 2013"}
    if($versions.Keys -notcontains $recommendVersion)
    {
        Throw "Parameter recommendVersion should be one of the following values: $($versions.Keys)!"
    }
    else
    {
        $recommendVersion = $versions[$recommendVersion]
    }
    $installedVersions = @()
    $installed = $false
    foreach($version in $versions.Keys)
    {  
        $registryPaths = "HKLM:\SOFTWARE\Microsoft\VisualStudio\$version", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\$version"
        $installPath = GetInstalledPath $registryPaths
        if($installPath -ne $null)
        {
            $installedVersions += $versions[$version]
        }  
    }   
    
    if($installedVersions)
    {        
        if($installedVersions -contains "$recommendVersion")
        {
            Output "The required application $recommendVersion have been installed" "Green"
            $installed = $true
        }
        else
        {   
            $flag = @{$true="are";$false="is"}[$installedVersions.Count -gt 1]
            $outPutWord = @()
            foreach($installedVersion in $installedVersions)
            {
                $outPutWord +=$installedVersion+","
            }
            $outPutWord[($outPutWord.Count)-1] = $outPutWord[($outPutWord.Count)-1].split(",")[0]
            Output ("Your installed Visual Studio """ + $outPutWord + """ $flag not recommended version, the recommended version is $recommendVersion")  "Yellow"
        }        
   } 
   else
   {
       Output "The application $recommendVersion is not installed" "Yellow"
   }
   
   Return $installed
   
}

#-----------------------------------------------------------------------------------
# <summary>
# Check whether the machine has installed the Protocol Test Framework
# </summary>
# <param name="recommendVersion">The recommend version of Protocol Test Framework.</param>
# <returns>A Boolean value, true if the Protocol Test Framework has been installed, otherwise false.</returns>
#-----------------------------------------------------------------------------------
Function CheckPTFVersion
{
    param(
    [string]$recommendVersion = "1.0.1128.0"
    )
    # The applications to be checked.
    $applicationName = "Protocol Test Framework"
    $registryPaths = "HKLM:\SOFTWARE\Microsoft\ProtocolTestFramework", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\ProtocolTestFramework"
    $applicationInstallPath= GetInstalledPath $registryPaths
    
    $installed = $false
    if($applicationInstallPath -ne $null)
    {
        $dllPaths = [System.IO.Directory]::GetFiles($applicationInstallPath,"Microsoft.Protocols.TestTools.dll",[System.IO.SearchOption]::AllDirectories)
        if($dllPaths -ne "" -and $dllPaths -ne $null)
        {
            if($dllPaths -is [array])
            {
                $dllPath = $dllPaths[0]
            }
            else
            {
                $dllPath = $dllPaths
            }
        }
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dllPath)
        if($versionInfo.ProductMajorPart -eq $recommendVersion.Split(".")[0] -and $versionInfo.ProductMinorPart -eq $recommendVersion.Split(".")[1] -and $versionInfo.ProductBuildPart -ge $recommendVersion.Split(".")[2])
        {
            $installed = $true
        }
                  
    }
            
    # Output the installed status.
    if($installed)
    {
        Output ("The required application " +$applicationName + " have been installed.") "Green"
    }
    else
    {
        Output ("The application " + $applicationName+ $recommendVersion + " or the newer version is not installed.") "Yellow"          
    }
    
    return $installed
       
}

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
            Output "The client setup script will run in unattended mode with information provided by the client configuration XML `"$unattendedXmlName`"." "White"
            $unattendedXmlName = Resolve-Path $unattendedXmlName
            break
        }
        else
        {
            Output "The client configuration XML path `"$unattendedXmlName`" is not correct." "Yellow"
            Output "Retry with the correct file path or press `"Enter`" if you want client setup script to run in attended mode?" "Cyan"
            $unattendedXmlName = Read-Host
        }
    }
}

#-----------------------------------------------------
# Check the application environment
#-----------------------------------------------------
Output "Check whether the required applications have been installed ..." "White"
$vsInstalledStatus = CheckVSVersion "12.0"
$ptfInstalledStatus = CheckPTFVersion "1.0.2220.0"
if(!$vsInstalledStatus -or !$ptfInstalledStatus)
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
# Check the Operating System (OS) version
#-----------------------------------------------------
Output "Check the Operating System (OS) version of the local machine ..." "White"
CheckOSVersion -computer localhost

#-----------------------------------------------------
# Configuration for common ptfconfig file.
#-----------------------------------------------------
Output "Configure the ExchangeCommonConfiguration.deployment.ptfconfig file ..." "White"
Output "Enter the computer name of the SUT:" "cyan"
Output "The computer name must be valid. Fully qualified domain name (FQDN) or IP address is not supported." "Yellow"
$sutComputerName = ReadComputerName "sutComputerName"
Output "Name of the SUT you entered: $sutcomputerName" "White"

Output "Enter the domain name of SUT(for example: contoso.com):" "cyan"
$dnsDomain = CheckForEmptyUserInput "Domain name" "dnsDomain"
Output "The domain name you entered: $dnsDomain" "White"

Output "Select the Microsoft Exchange Server version" "cyan"
Output "If you are running your own server implementation, choose the closest exchange server version which matches your implementation." "cyan"
Output "1: Microsoft Exchange Server 2007" "cyan"
Output "2: Microsoft Exchange Server 2010" "cyan"
Output "3: Microsoft Exchange Server 2013" "cyan"

$sutVersions =@('1','2','3')
$sutVersion = ReadUserChoice $sutVersions "sutVersion" 
Switch ($sutVersion)
{
    "1" { $sutVersion = "ExchangeServer2007"; $protocolVersion ="12.1"; break }
    "2" { $sutVersion = "ExchangeServer2010"; break }
    "3" { $sutVersion = "ExchangeServer2013"; break }
}
Output "The SUT version you selected is $sutVersion." "White"

if($sutVersion -ge "ExchangeServer2010")
{
    Output "Select ActiveSync protocol version. Test suites will use this version while sending requests." "cyan"
    Output "1: Protocol version is 12.1" "cyan"
    Output "2: Protocol version is 14.0" "cyan"
    Output "3: Protocol version is 14.1" "cyan"
    $protocolVersions =@('1','2','3')
    $protocolVersion = ReadUserChoice $protocolVersions "protocolVersion"
    Switch ($protocolVersion)
    {
        "1" {$protocolVersion = "12.1"; break}
        "2" {$protocolVersion = "14.0"; break}
        "3" {$protocolVersion = "14.1"; break}
    }
}
Output "The ActiveSync protocol version you selected is $protocolVersion." "White"

Output "Select the transport type" "cyan"
Output "1: HTTP" "cyan"
Output "2: HTTPS" "cyan"

$transportTypes =@('1','2')
$transportType = ReadUserChoice $transportTypes "transportType"
Switch ($transportType)
{
    "1" { $transportType = "HTTP";  break }
    "2" { $transportType = "HTTPS"; break }
}
Output "Transport type you entered: $transportType" "White"

Output "Select encoding scheme for the URL query string." "cyan"
Output "1: Test suites will use base64 encoding for the URL query string" "cyan"
Output "2: Test suites will use plaintext encoding for the URL query string" "cyan"
$headerEncodingTypes =@('1','2')
$headerEncodingType = ReadUserChoice $headerEncodingTypes "headerEncodingType"
Switch ($headerEncodingType)
{
    "1" {$headerEncodingType = "Base64"; break}
    "2" {$headerEncodingType = "PlainText"; break}
}
Output "Head encoding type you selected is $headerEncodingType." "White"

Output "Add SUT machine to the TrustedHosts configuration setting to ensure WinRM client can process remote calls against SUT machine." "Yellow"
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

Output "Modify the properties as necessary in the ExchangeCommonConfiguration.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $commonDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"Domain`", and set the value as $dnsDomain" "Yellow"
$step++
Output "$step.Find the property `"SutComputerName`", and set the value as $sutcomputerName" "Yellow"
$step++
Output "$step.Find the property `"SutVersion`", and set the value as $sutVersion" "Yellow"
$step++
Output "$step.Find the property `"TransportType`", and set the value as $transportType" "Yellow"
$step++
Output "$step.Find the property `"ActiveSyncProtocolVersion`", and set the value as $protocolVersion" "Yellow"
$step++
Output "$step.Find the property `"HeaderEncodingType`", and set the value as $headerEncodingType" "Yellow"

ModifyConfigFileNode $commonDeploymentFile "Domain"                      $dnsDomain
ModifyConfigFileNode $commonDeploymentFile "SutComputerName"             $sutComputerName
ModifyConfigFileNode $commonDeploymentFile "SutVersion"                  $sutVersion
ModifyConfigFileNode $commonDeploymentFile "TransportType"               $transportType
ModifyConfigFileNode $commonDeploymentFile "ActiveSyncProtocolVersion"   $protocolVersion
ModifyConfigFileNode $commonDeploymentFile "HeaderEncodingType"          $headerEncodingType

Output "Configuration for the ExchangeCommonConfiguration.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASAIRS ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASAIRS_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASAIRSDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASAIRSUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASAIRSUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User1Name"             "$MSASAIRSUser01"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User1Password"         "$userPassword"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User2Name"             "$MSASAIRSUser02"
ModifyConfigFileNode $MSASAIRSDeploymentFile    "User2Password"         "$userPassword"
Output "Configuration for the MS-ASAIRS_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASCAL ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASCAL_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASCALDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"OrganizerUserName`", and set the value as $MSASCALUser01" "Yellow"
$step++
Output "$step.Find the property `"OrganizerUserPassword`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"AttendeeUserName`", and set the value as $MSASCALUser02" "Yellow"
$step++
Output "$step.Find the property `"AttendeeUserPassword`", and set the value as $userPassword" "Yellow"
ModifyConfigFileNode $MSASCALDeploymentFile    "OrganizerUserName"            "$MSASCALUser01"
ModifyConfigFileNode $MSASCALDeploymentFile    "OrganizerUserPassword"        "$userPassword"
ModifyConfigFileNode $MSASCALDeploymentFile    "AttendeeUserName"             "$MSASCALUser02"
ModifyConfigFileNode $MSASCALDeploymentFile    "AttendeeUserPassword"         "$userPassword"

Output "Configuration for the MS-ASCAL_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASCMD ptfconfig file.
#-------------------------------------------------------
$MSASCMDSharedFolderPath = "\\[SutComputerName]\$MSASCMDSharedFolder"
$MSASCMDNonEmptyDocumentPath= "[SharedFolder]\$MSASCMDNonEmptyDocument"
$MSASCMDEmptyDocumentPath= "[SharedFolder]\$MSASCMDEmptyDocument"

Output "Modify the properties as necessary in the MS-ASCMD_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASCMDDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASCMDUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASCMDUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASCMDUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User7Name`", and set the value as $MSASCMDUser07" "Yellow"
$step++
Output "$step.Find the property `"User7Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User8Name`", and set the value as $MSASCMDUser08" "Yellow"
$step++
Output "$step.Find the property `"User8Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User9Name`", and set the value as $MSASCMDUser09" "Yellow"
$step++
Output "$step.Find the property `"User9Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"GroupDisplayName`", and set the value as $MSASCMDTestGroup " "Yellow"
$step++
Output "$step.Find the property `"LargeGroupDisplayName`", and set the value as $MSASCMDLargeGroup " "Yellow"
$step++
Output "$step.Find the property `"SharedFolder`", and set the value as $MSASCMDSharedFolderPath " "Yellow"
$step++
Output "$step.Find the property `"SharedDocument1`", and set the value as $MSASCMDNonEmptyDocumentPath " "Yellow"
$step++
Output "$step.Find the property `"SharedDocument2`", and set the value as $MSASCMDEmptyDocumentPath" "Yellow"
$step++
Output "$step.Find the property `"MIMEMailSubject`", and set the value as $MSASCMDEmailSubjectName" "Yellow"

ModifyConfigFileNode $MSASCMDDeploymentFile    "User1Name"                    "$MSASCMDUser01"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User1Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User2Name"                    "$MSASCMDUser02"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User2Password"                "$userPassword"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User3Name"                    "$MSASCMDUser03"
ModifyConfigFileNode $MSASCMDDeploymentFile    "User3Password"                "$userPassword"
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

Output "Configuration for the MS-ASCMD_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASCNTC ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASCNTC_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASCNTCDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASCNTCUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASCNTCUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASCNTCDeploymentFile    "User1Name"            $MSASCNTCUser01
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User1Password"        $userPassword
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User2Name"            $MSASCNTCUser02
ModifyConfigFileNode $MSASCNTCDeploymentFile    "User2Password"        $userPassword

Output "Configuration for the MS-ASCNTC_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASCON ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASCON_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASCONDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASCONUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASCONUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASCONUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASCONDeploymentFile "User1Name"      $MSASCONUser01
ModifyConfigFileNode $MSASCONDeploymentFile "User1Password"  $userPassword
ModifyConfigFileNode $MSASCONDeploymentFile "User2Name"      $MSASCONUser02
ModifyConfigFileNode $MSASCONDeploymentFile "User2Password"  $userPassword
ModifyConfigFileNode $MSASCONDeploymentFile "User3Name"      $MSASCONUser03
ModifyConfigFileNode $MSASCONDeploymentFile "User3Password"  $userPassword

Output "Configuration for the MS-ASCON_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASDOC ptfconfig file.
#-------------------------------------------------------
# Get the property value of MS-ASDOC ptfconfig file.
$MSASDOCSharedFolderPath = "\\[SutComputerName]\$MSASDOCSharedFolder"
$visibleDocumentPath = "[SharedFolder]\$MSASDOCVisibleDocument"
$hiddenDocumentPath = "[SharedFolder]\$MSASDOCHiddenDocument"
$hiddenFolderPath = "[SharedFolder]\$MSASDOCHiddenFolder"
$visibleFolderPath = "[SharedFolder]\$MSASDOCVisibleFolder"

Output "Modify the properties as necessary in the MS-ASDOC_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASDOCDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"UserName`", and set the value as $MSASDOCUser01" "Yellow"
$step++
Output "$step.Find the property `"UserPassword`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"SharedFolder`", and set the value as $MSASDOCSharedFolderPath" "Yellow"
$step++
Output "$step.Find the property `"SharedHiddenDocument`", and set the value as $hiddenDocumentPath" "Yellow"
$step++
Output "$step.Find the property `"SharedVisibleDocument`", and set the value as $visibleDocumentPath" "Yellow"
$step++
Output "$step.Find the property `"SharedHiddenFolder`", and set the value as $hiddenFolderPath" "Yellow"
$step++
Output "$step.Find the property `"SharedVisibleFolder`", and set the value as $visibleFolderPath" "Yellow"

ModifyConfigFileNode $MSASDOCDeploymentFile "UserName"                   "$MSASDOCUser01"
ModifyConfigFileNode $MSASDOCDeploymentFile "UserPassword"               "$userPassword"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedFolder"               "$MSASDOCSharedFolderPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedHiddenDocument"       "$hiddenDocumentPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedVisibleDocument"      "$visibleDocumentPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedHiddenFolder"         "$hiddenFolderPath"
ModifyConfigFileNode $MSASDOCDeploymentFile "SharedVisibleFolder"        "$visibleFolderPath"

Output "Configuration for the MS-ASDOC_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASEMAIL ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASEMAIL_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASEMAILDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASEMAILUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASEMAILUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASEMAILUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASEMAILDeploymentFile "User1Name"                   $MSASEMAILUser01
ModifyConfigFileNode $MSASEMAILDeploymentFile "User1Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User2Name"                   $MSASEMAILUser02
ModifyConfigFileNode $MSASEMAILDeploymentFile "User2Password"               $userPassword
ModifyConfigFileNode $MSASEMAILDeploymentFile "User3Name"                   $MSASEMAILUser03
ModifyConfigFileNode $MSASEMAILDeploymentFile "User3Password"               $userPassword

Output "Configuration for the MS-ASEMAIL_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASHTTP ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASHTTP_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASHTTPDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASHTTPUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASHTTPUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASHTTPUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User4Name`", and set the value as $MSASHTTPUser04" "Yellow"
$step++
Output "$step.Find the property `"User4Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASHTTPDeploymentFile "User1Name"                  $MSASHTTPUser01
ModifyConfigFileNode $MSASHTTPDeploymentFile "User1Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User2Name"                  $MSASHTTPUser02
ModifyConfigFileNode $MSASHTTPDeploymentFile "User2Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User3Name"                  $MSASHTTPUser03
ModifyConfigFileNode $MSASHTTPDeploymentFile "User3Password"              $userPassword
ModifyConfigFileNode $MSASHTTPDeploymentFile "User4Name"                  $MSASHTTPUser04
ModifyConfigFileNode $MSASHTTPDeploymentFile "User4Password"              $userPassword

Output "Configuration for the MS-ASHTTP_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASNOTE ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASNOTE_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASNOTEDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"UserName`", and set the value as $MSASNOTEUser01" "Yellow"
$step++
Output "$step.Find the property `"UserPassword`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASNOTEDeploymentFile    "UserName"             "$MSASNOTEUser01"
ModifyConfigFileNode $MSASNOTEDeploymentFile    "UserPassword"         "$userPassword"
Output "Configuration for the MS-ASNOTE_TestSuite.deployment.ptfconfig file is complete." "Green"
#-------------------------------------------------------
# Configuration for MS-ASPROV ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASPROV_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASPROVDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASPROVUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASPROVUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASPROVUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASPROVDeploymentFile "User1Name"                $MSASPROVUser01
ModifyConfigFileNode $MSASPROVDeploymentFile "User1Password"            $userPassword
ModifyConfigFileNode $MSASPROVDeploymentFile "User2Name"                $MSASPROVUser02
ModifyConfigFileNode $MSASPROVDeploymentFile "User2Password"            $userPassword
ModifyConfigFileNode $MSASPROVDeploymentFile "User3Name"                $MSASPROVUser03
ModifyConfigFileNode $MSASPROVDeploymentFile "User3Password"            $userPassword

Output "Configuration for the MS-ASPROV_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASRM ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASRM_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASRMDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"User1Name`", and set the value as $MSASRMUser01" "Yellow"
$step++
Output "$step.Find the property `"User1Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User2Name`", and set the value as $MSASRMUser02" "Yellow"
$step++
Output "$step.Find the property `"User2Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User3Name`", and set the value as $MSASRMUser03" "Yellow"
$step++
Output "$step.Find the property `"User3Password`", and set the value as $userPassword" "Yellow"
$step++
Output "$step.Find the property `"User4Name`", and set the value as $MSASRMUser04" "Yellow"
$step++
Output "$step.Find the property `"User4Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASRMDeploymentFile "User1Name"                $MSASRMUser01
ModifyConfigFileNode $MSASRMDeploymentFile "User1Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User2Name"                $MSASRMUser02
ModifyConfigFileNode $MSASRMDeploymentFile "User2Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User3Name"                $MSASRMUser03
ModifyConfigFileNode $MSASRMDeploymentFile "User3Password"            $userPassword
ModifyConfigFileNode $MSASRMDeploymentFile "User4Name"                $MSASRMUser04
ModifyConfigFileNode $MSASRMDeploymentFile "User4Password"            $userPassword

Output "Configuration for the MS-ASRM_TestSuite.deployment.ptfconfig file is complete." "Green"

#-------------------------------------------------------
# Configuration for MS-ASTASK ptfconfig file.
#-------------------------------------------------------
Output "Modify the properties as necessary in the MS-ASTASK_TestSuite.deployment.ptfconfig file..." "White"
$step=1
Output "Steps for manual configuration:" "Yellow"
Output "$step.Open $MSASTASKDeploymentFile" "Yellow"
$step++
Output "$step.Find the property `"UserName`", and set the value as $MSASTASKUser01" "Yellow"
$step++
Output "$step.Find the property `"Password`", and set the value as $userPassword" "Yellow"

ModifyConfigFileNode $MSASTASKDeploymentFile "UserName"           $MSASTASKUser01
ModifyConfigFileNode $MSASTASKDeploymentFile "Password"           $userPassword

Output "Configuration for the MS-ASTASK_TestSuite.deployment.ptfconfig file is complete." "Green"

#----------------------------------------------------------------------------
# End script
#----------------------------------------------------------------------------

Output "Client configuration script was executed successfully." "Green"
Stop-Transcript
exit 0