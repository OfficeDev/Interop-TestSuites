#-------------------------------------------------------------------------
# Configuration script exit code definition:
# 1. A normal termination will set the exit code to 0
# 2. An uncaught THROW will set the exit code to 1
# 3. Script execution warning and issues will set the exit code to 2
# 4. Exit code is set to the actual error code for other issues
#-------------------------------------------------------------------------

#-----------------------------------------------------------------------------------
# <summary>
# Add an user policy to a web application without name-prefixed.
# </summary>
# <param name="webApplicationUrl">The url of the web application.</param>
# <param name="user">A valid account which could be a domain user or a local machine user, the account format is as follows: "DomainName\UserName" or "Computername\UserName".</param>
# <param name="roleName">Optional. Name of the policy role. By default, it is "Full Control".</param>
#-----------------------------------------------------------------------------------
function AddUserPolicyWithoutNamePrefix
{
    Param(
    [String]$webApplicationUrl,
    [String]$user,
    [String]$roleName = "Full Control"
    )

    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($webApplicationUrl -eq $null -or $webApplicationUrl -eq "")
    {
        Throw "The parameter webApplicationUrl cannot be empty."
    }

    if($user -eq $null -or $user -eq "")
    {
        Throw "The parameter user cannot be empty."
    }

    #----------------------------------------------------------------------------
    # Main function.
    #----------------------------------------------------------------------------
    $uri = new-object System.Uri($webApplicationUrl)
    $webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($uri)
    $isWindowsAuthentication = $false
    try
    {
        $policyRole = $webApp.PolicyRoles[$roleName]
        if ($policyRole -eq $null)
        {
            Throw  "[AddUserPolicyWithoutNamePrefix] Cannot find the policy role '$roleName' in the web application '$webApplicationUrl'."
        }

        # Add the user policy and specify its permission role. 
        $policy = $webApp.Policies | where {$_.UserName -eq $user}
        
        if ($policy -eq $null -or $policy.PolicyRoleBindings -eq $null -or $policy.PolicyRoleBindings[$policyRole] -eq $null)
        {
            $useClaims = $webApp.UseClaimsAuthentication
            # If the web application uses claims-based authentication, change it to windows in 
            # order to add user policy without name prefixed.
            if ($useClaims -eq $true)
            {
                $webApp.UseClaimsAuthentication = $false
                $webApp.Update()
                $isWindowsAuthentication = $true
                Output "[AddUserPolicyWithoutNamePrefix] The web application '$webApplicationUrl' has been changed to Windows authentication in order to add a user policy without a prefixed name." "Green"
            }
            $policy = $webApp.Policies.Add($user, $user)
            $policy.PolicyRoleBindings.Add($policyRole)
            $webApp.Update()
            Output "[AddUserPolicyWithoutNamePrefix] The user '$user' has been added to the user policy collection of the web application '$webApplicationUrl' with the role '$roleName'." "Green"
        }
        else
        {
            Output "The specific user policy for the user $user already exists in the web application '$webApplicationUrl'." "Yellow"
        }
    }
    finally
    {
        # Restore if the web application uses claims-based authentication.
        if ($isWindowsAuthentication -eq $true)
        {
            $webApp.UseClaimsAuthentication = $true
            $webApp.Update()
            Start-Sleep 10
            Output "[AddUserPolicyWithoutNamePrefix] The web application '$webApplicationUrl' that uses claims-based authentication has been restored." "Green"
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Compare the recommended SharePoint minor version with the installed SharePoint minor version.
# </summary>
# <param name="actualVersion">The display version of the SharePoint installed currently.</param>
# <param name="recommendedVersion">An array with three elements, the recommended SharePoint display version.</param>
# <returns>
# A Boolean value, true if the server has the recommended service pack installed, otherwise false.
# </returns>
#-----------------------------------------------------------------------------------          
function CompareSharePointMinorVersion
{
    param(
    [String]$actualVersion,
    [String]$recommendedVersion
    )
    $actualVersionBuildNumber = $actualVersion.split(".")[2]
    $recommendedVersionBuildNumber = $recommendedVersion.split(".")[2]
    
    if($actualVersionBuildNumber -eq $recommendedVersionBuildNumber)
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
# Get the SharePoint Server Version. 
# </summary>
# <returns>
# Return an array include SharePoint version info:
# 1.String value:SharePoint Name.
# 2.Boolean value: true means Sharepoint version is recommended version;otherwise,false.
# 3.String value:SharePoint Service Pack name.
# </returns>
#-----------------------------------------------------------------------------------
function GetSharePointVersion
{
    $script:WindowsSharePointServices3     = "Microsoft Windows SharePoint Services 3.0", "12.0.6612.1000", "SP3"
    $script:SharePointServer2007           = "Microsoft Office SharePoint Server 2007",   "12.0.6612.1000", "SP3"
    $script:SharePointFoundation2010       = "Microsoft SharePoint Foundation 2010",      "14.0.7015.1000", "SP2"
    $script:SharePointServer2010           = "Microsoft SharePoint Server 2010",          "14.0.7015.1000", "SP2"
    $script:SharePointFoundation2013       = "Microsoft SharePoint Foundation 2013",      "15.0.4571.1502", "SP1"
    $script:SharePointServer2013           = "Microsoft SharePoint Server 2013",          "15.0.4571.1502", "SP1"
    $script:SharePointServer2016           = "Microsoft SharePoint Server 2016",          "16.0.4351.1000", ""
    $script:SharePointServer2019           = "Microsoft SharePoint Server 2019",          "16.0.10711.37301", ""
    $script:SharePointServerSubscriptionEdition           = "Microsoft SharePoint Server Subscription Edition",          "16.0.14326.20450", ""
    $SharePointVersion                     = "Unknown Version"   
    
    $keys = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
    $items = $keys | foreach-object {Get-ItemProperty $_.PsPath}
    $SharePointServer2013NameInKey = $script:SharePointServer2013[0] + " "
    $SharePointServer2007NameInKey = $script:SharePointServer2007[0] + " "
    foreach ($item in $items)
    {
        if($item.DisplayName -ne $null -and $item.DisplayName -ne "")
        {
            if($item.DisplayName -eq $script:WindowsSharePointServices3[0])
            {                
                $version = $item.DisplayVersion
                $SharePointVersion = $script:WindowsSharePointServices3[0]
                $recommendVersion = $script:WindowsSharePointServices3[1]
                $recommendMinorVersion = $script:WindowsSharePointServices3[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                foreach ($item in $items)
                {   
                    if($item.DisplayName -eq "$SharePointServer2007NameInKey")
                    {
                        $version = $item.DisplayVersion
                        $SharePointVersion = $script:SharePointServer2007[0]
                        $recommendVersion = $script:SharePointServer2007[1]
                        $recommendMinorVersion = $script:SharePointServer2007[2]
                        $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                        break
                    }
                }
                break
            }
            elseif($item.DisplayName -eq $script:SharePointFoundation2010[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointFoundation2010[0]
                $recommendVersion = $script:SharePointFoundation2010[1]
                $recommendMinorVersion = $script:SharePointFoundation2010[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }        
            elseif($item.DisplayName -eq $script:SharePointServer2010[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointServer2010[0]
                $recommendVersion = $script:SharePointServer2010[1]
                $recommendMinorVersion = $script:SharePointServer2010[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }        
            elseif($item.DisplayName -eq $script:SharePointFoundation2013[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointFoundation2013[0]
                $recommendVersion = $script:SharePointFoundation2013[1]
                $recommendMinorVersion = $script:SharePointFoundation2013[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }        
            elseif($item.DisplayName -eq "$SharePointServer2013NameInKey")
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointServer2013[0]
                $recommendVersion = $script:SharePointServer2013[1]
                $recommendMinorVersion = $script:SharePointServer2013[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }
            elseif($item.DisplayName -eq $script:SharePointServer2016[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointServer2016[0]
                $recommendVersion = $script:SharePointServer2016[1]
                $recommendMinorVersion = $script:SharePointServer2016[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }
            elseif($item.DisplayName -eq $script:SharePointServer2019[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointServer2019[0]
                $recommendVersion = $script:SharePointServer2019[1]
                $recommendMinorVersion = $script:SharePointServer2019[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }
            elseif($item.DisplayName -eq $script:SharePointServerSubscriptionEdition[0])
            {
                $version = $item.DisplayVersion
                $SharePointVersion = $script:SharePointServerSubscriptionEdition[0]
                $recommendVersion = $script:SharePointServerSubscriptionEdition[1]
                $recommendMinorVersion = $script:SharePointServerSubscriptionEdition[2]
                $isRecommendMinorVersion = CompareSharePointMinorVersion $version $recommendVersion
                break
            }
        }
    }
    
    $SharePointVersion = $SharePointVersion,$version,$isRecommendMinorVersion,$recommendVersion,$recommendMinorVersion    
    return $SharePointVersion    
}

#-----------------------------------------------------------------------------------
# <summary>
# Output the detailed sharepoint version info.
# </summary>
#-----------------------------------------------------------------------------------   
function OutPutSupportVersionInfo
{       
    $sharePointVersionInfo = GetSharePointVersion
    if($sharePointVersionInfo[2])
    {
        Output ("SharePoint Server Version: $($sharePointVersionInfo[0]) $($sharePointVersionInfo[4]).") "White"
    }
    else
    {
        Output "$($sharePointVersionInfo[0]) $($sharePointVersionInfo[1]) is not the recommended version." "Yellow"
        Output ("Please install the recommended $($sharePointVersionInfo[0]) $($sharePointVersionInfo[3]), otherwise some cases might fail.") "Yellow"
        Output "Would you like to continue configuring the server or exit?" "Cyan"
        Output "1: CONTINUE." "Cyan"
        Output "2: EXIT." "Cyan"
        $runOnNonRecommendedSUTChoices = @('1: CONTINUE','2: EXIT')
        $runOnNonRecommendedSUT = ReadUserChoice $runOnNonRecommendedSUTChoices "runOnNonRecommendedSUT"
        if ($runOnNonRecommendedSUT -eq "2")
        {
            Stop-Transcript
            exit 0
        }
    }
}
 
#-----------------------------------------------------------------------------------
# <summary>
# Configure SharePoint server to support HTTPS transport. 
# </summary>
# <param name="computerName">The computer name of the machine where SharePoint Server is installed.</param>
# <param name="SharePointVersion">The sharePoint server version.
# Note: The value is obtained by calling function "GetSharePointVersion".</param>
# <param name="webAppName">The Web Application name of the SharePoint Server,the default value of SUT web-site is SharePoint - 80.</param>
# <param name="httpsPortNumber">The port number is used in HTTPS binding,the default value of SUT web-site is 443.</param>
# <param name="isSUTCentralAdministration">Boolean value, true means binding https on SUT Central Administration; false means binding https on SUT web-site.</param>
# <param name="httpPortNumber">The port number of SUT web-site,the default value is 80.</param>
#-----------------------------------------------------------------------------------
function AddHTTPSBinding
{
    Param(
    [string]$computerName,
    [string]$SharePointVersion,
    [string]$webAppName = "SharePoint - 80",
    [string]$httpsPortNumber = "443",
    [bool]$isSUTCentralAdministration = $false,
    [string]$httpPortNumber = "80"
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($SharePointVersion -eq $null -or $SharePointVersion -eq "")
    {
        Throw "The parameter SharePointVersion cannot be empty."
    }
    
    # Set parameter according to $isSUTCentralAdministration.
    if($isSUTCentralAdministration)
    {
        $outPutKeyWords = "SUT Central Administration"
        $adminUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.sites[0].url.Split(":")
        $adminNum = $adminUrl[2]
        $siteURL = "http://${computerName}:${adminNum}"
        $SPAlternateUrl = "https://${computerName}:${httpsPortNumber}"        
    }
    else
    {
        $outPutKeyWords = "SUT web-site"
        $siteURL = "http://${computerName}:${httpPortNumber}"
        $SPAlternateUrl = "https://${computerName}:${httpsPortNumber}"
    }
    
    $spWebService = "Microsoft SharePoint Foundation Web Application"
    if($SharePointVersion -eq $SharePointServer2007[0] -or $SharePointVersion -eq $WindowsSharePointServices3[0])
    {
        $spWebService = "Windows SharePoint Services Web Application"
    }
    #----------------------------------------------------------------------------
    # Variable definition.
    #----------------------------------------------------------------------------
    Import-Module Servermanager
    Add-WindowsFeature Web-Mgmt-Service | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
    $cert = dir cert:\localmachine\my | where {$_.Subject -like "CN=WMSvc-*"}
    
    if($cert -eq $null -or $cert -eq "")
    {
        throw "[AddHTTPSBinding] Cannot get the WMSvc certificate."
    }
    else
    {     
        $certHash = $cert.GetCertHashString()
    }    

    #----------------------------------------------------------------------------
    # Binding HTTPS to web application.
    #----------------------------------------------------------------------------
    $guid = [Guid]::NewGuid()
    $appId = $guid.ToString()
    $httpsInfor = cmd.exe /c netsh http show sslcert ipport=0.0.0.0:$httpsPortNumber
    $siteBindingInfor = cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" list site $webAppName
    if($httpsInfor[4].Contains("IP:port") -and ($siteBindingInfor -like ("*https*")))
    {
        Output "HTTPS has already been configured." "Yellow"
    }
    else
    {
        $bindinginformation = "*:$httpsPortNumber"+":"
        cmd.exe /c netsh http add sslcert ipport=0.0.0.0:$httpsPortNumber certhash = $certHash appid = "{$appId}"
        cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" set site $webAppName /+"bindings.[protocol='https',bindinginformation="`'$bindinginformation`'"]" /commit:apphost

        if(!$?)
        {
            Throw "Failed to bind the site $webAppName over HTTPS (port $httpsPortNumber)."
        }
    }

    #----------------------------------------------------------------------------
    # Load SharePoint Snap-in.
    #----------------------------------------------------------------------------
    Output "Set an alternate access mapping for HTTPS." "White"
    
    if($isSUTCentralAdministration)
    {
        $alternateUrls = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.AlternateUrls
    }
    else
    {
        $webApplicationUri = new-object Uri($siteURL);
        $webApplication = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($webApplicationUri);
        if(!$webApplication)
        {
            $applicationWebService = [Microsoft.SharePoint.Administration.SPFarm]::Local.Services | where {$_.TypeName -eq $spWebService}
            $applicationWebServicePool = new-object Microsoft.SharePoint.Administration.SPApplicationPool($webAppName,$applicationWebService)
            $webApplication = new-object Microsoft.SharePoint.Administration.SPWebApplication($webAppName,$applicationWebService,$applicationWebServicePool)
            $webApplication.Update($true)
        }
        
        $alternateUrls = $webApplication.AlternateUrls;
        $defaultAlternateUrl = new-object Microsoft.SharePoint.Administration.SPAlternateUrl($webApplicationUri,[Microsoft.SharePoint.Administration.SPUrlZone]::Default)
        if(!$alternateUrls.Contains($defaultAlternateUrl))
        {
            $alternateUrls.Add($defaultAlternateUrl)
        }

        $alternateUrls.Update()        
    }
        
    $altUrl = new-Object Microsoft.SharePoint.Administration.SPAlternateUrl($SPAlternateUrl, [Microsoft.SharePoint.Administration.SPUrlZone]::Internet)
    if(!$alternateUrls.Contains($altUrl))
    {
        $alternateUrls.Add($altUrl)
    }
    Output "[AddHTTPSBinding] The access mapping is set to HTTPS." "Green"
    
    #----------------------------------------------------------------------------
    # Restart web application.
    #----------------------------------------------------------------------------
    Output "[AddHTTPSBinding] Restart the web application." "Yellow"
    RestartWebApplication $webAppName        
  
}

#-----------------------------------------------------------------------------------
# <summary>
# Restart the web application. 
# </summary>
# <param name="webAppName">The name of the web application.</param>
#-----------------------------------------------------------------------------------
function RestartWebApplication
{
    param(
    [string]$webAppName
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($webAppName -eq $null -or $webAppName -eq "")
    {
        Throw "The parameter webAppName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Restart web application.
    #----------------------------------------------------------------------------
    cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" stop site $webAppName
    if(!$?)
    {
        Throw "Failed to stop the site $webAppName"
    }
    
    cmd.exe /c "$env:windir\system32\inetsrv\appcmd.exe" start site $webAppName
    if(!$?)
    {
        Throw "Failed to start the site $webAppName"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the name of web application. 
# </summary>
# <param name="defaultWebAppName">The default name of the web application.</param>
#-----------------------------------------------------------------------------------
function GetWebAPPName
{
    param(
    [string]$defaultWebAppName = "SharePoint - 80"
    )    
   
    #----------------------------------------------------------------------------
    # Get the name of web application.
    #----------------------------------------------------------------------------
    import-module WebAdministration
    $webApplication = Get-Website
    if($webApplication | Where {$_.Name -eq $defaultWebAppName})
    {
        $webAppName = $defaultWebAppName
        return $webAppName
    }
    else
    {
        Output "The default SUT website $defaultWebAppName does not exist. Enter the website name on the web server:" "Cyan"
        $webAppName = CheckForEmptyUserInput "The web application name" "webAppName"
        if($webApplication | Where {$_.Name -eq $webAppName})
        {    
            return $webAppName
        }
        else
        {                 
            throw "The SUT website $webAppName does not exist on the web server."
        }        
    }
}


#-----------------------------------------------------------------------------------
# <summary>
# Delete the specified web and its sub webs. 
# </summary>
# <param name="web">The web to be deleted.</param>
#-----------------------------------------------------------------------------------
function DeleteWeb 
{
    Param ( 
    [Microsoft.SharePoint.SPWeb]$web
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    
    $subWebs = $web.GetSubwebsForCurrentUser()
    
    if($subWebs.Count -ne 0)
    {
        foreach($subweb in $subWebs)
        {
            DeleteWeb $subweb
            $subweb.Dispose()
        }
    }
    
    $web.Delete()
    Output ("[DeleteWeb] The web  """ + $web.Url + """ is deleted") "Green"
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a web with the specified website-relative URL, title and web template or 
# create a default web with the specified website-relative URL.  
# </summary>
# <param name="parentSite">The site collection where the web is located.</param>
# <param name="isDefault">>Boolean value, true means create a default web with the specified website-relative URL;
# otherwise,false.</param>
# <param name="webUrl">A string that contains the new website URL relative to the 
# root website in the site collection.</param>
# <param name="webTitle">A string that contains the title.</param>
# <param name="webTemplate">A string that represents the site definition or site template.</param>
# <param name="uniquePermission">Boolean value, true to create a subsite that does not inherit 
# permissions from another site; otherwise, false.</param>
# <param name="webDescription">A string that contains the description.</param>
# <param name="webLCID">A 32-bit unsigned integer that specifies the locale ID.</param>
# <returns>A Microsoft.SharePoint.SPWeb object that represents the web site.</returns>
#-----------------------------------------------------------------------------------
function CreateWeb 
{
    Param ( 
    [Microsoft.SharePoint.SPSite]$parentSite,
    [bool]$isDefault = $false,
    [string]$webUrl,
    [string]$webTitle,
    [string]$webTemplate,
    [bool]$uniquePermission = $false,
    [string]$webDescription = "",
    [uint32] $webLCID = $null
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($parentSite -eq $Null -or $parentSite -eq "")
    {
        Throw "The parameter parentSite cannot be empty."
    }
    if($webUrl -eq $null -or $webUrl -eq "")
    {
        Throw "The parameter webUrl cannot be empty."
    }
    if($webTitle -eq $null -or $webTitle -eq "")
    {
        $count = $webUrl.Split('/').count
        $webTitle = $webUrl.Split('/')[$count -1]
    }
    if($isDefault -eq $false)
    {
        if($webTemplate -eq $null -or $webTemplate -eq "")
        {
            Throw "The parameter webTemplate cannot be empty."
        }
    }
        
    $web = $parentSite.OpenWeb($webUrl)
    if ($web.Exists)
    {
        OutPut "[CreateWeb] The $webUrl already exists. Delete it first and then create a new one." "Yellow"
        DeleteWeb $web
        $web.Dispose()
    }
    if($isDefault)
    {
         $web = $parentSite.AllWebs.Add($webUrl)
    }
    else
    {
        $web = $parentSite.AllWebs.Add($webUrl, $webTitle, $webDescription, $webLCID, $webTemplate, $uniquePermission, $false)
    }
    Output ("[CreateWeb] The web ""$webUrl"" has been created under web """ + $parentSite.Url + """") "Green"
    return $web
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a folder that is located at the specified URL to the specified web. 
# </summary>
# <param name="web">The web where the folder is located.</param>
# <param name="folderUrl">A string that specifies the URL of the folder.</param>
# <param name="overwrite">Boolean value, true means overwrite the folder with the same url; otherwise false.</param>
# <returns>A Microsoft.SharePoint.SPFolder object that represents the folder.</returns>
#-----------------------------------------------------------------------------------
function CreateSharePointFolder
{
    Param ( 
    [Microsoft.SharePoint.SPWeb]$web,
    [string]$folderUrl,
    [bool]$overwrite = $false
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if ($folderUrl -eq $null -or $folderUrl -eq "")
    {
        Throw "The parameter folderUrl cannot be empty."
    }
    

    #----------------------------------------------------------------------------
    # Check if the specified folder already exists.
    #----------------------------------------------------------------------------
    $originalFolder = $web.GetFolder($folderUrl)
    if ($originalFolder.Exists)
    {
        if ($overwrite)
        {
            $originalFolder.Delete() | Out-Null
            Output "[CreateSharePointFolder] The folder '$folderUrl' is deleted." "Yellow"
            
            # Create the folder.
            $web.Folders.Add($folderUrl) | Out-Null
            Output "[CreateSharePointFolder] The folder '$folderUrl' is created on $webName." "Green"
        }
        else
        {
            Output "[CreateSharePointFolder] Failed to create the folder. The specified folder '$folderUrl' already exists. " "Yellow"
        }
    }
    else
    {
        # Create the folder.
        $web.Folders.Add($folderUrl) | Out-Null
        Output ("[CreateSharePointFolder] The folder '$folderUrl' has been created under """ + $web.Url + """") "Green"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Upload a file into a specific folder on the web site.
# </summary>
# <param name="web">The web where the folder is located.</param>
# <param name="folderUrl">A string that specifies the URL of the folder.</param>
# <param name="fileName">The name of the file.</param>
# <param name="fileLocalPath">The local path of the file.</param>
# <param name="overwrite">Boolean value, true means overwrite the file with the same name; otherwise false.</param>
#-----------------------------------------------------------------------------------
function UploadFileToSharePointFolder
{
    Param ( 
    [Microsoft.SharePoint.SPWeb]$web,
    [string]$folderUrl,
    [string]$fileName,
    [string]$fileLocalPath,
    [bool]$overwrite = $false
    )

    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if ($folderUrl -eq $null -or $folderUrl -eq "")
    {
        Throw "The parameter folderUrl cannot be empty."
    }
    if ($fileName -eq $null -or $fileName -eq "")
    {
        Throw "The parameter fileName cannot be empty."
    }
    if ($fileLocalPath -eq $null -or $fileLocalPath -eq "")
    {
        Throw "The parameter fileLocalPath cannot be empty."
    }

    #----------------------------------------------------------------------------
    # Upload the file into the specific folder of the specific site.
    #----------------------------------------------------------------------------
    $spFolder =$web.GetFolder($folderUrl)
    if ($spFolder.Exists)
    {
        $file = $spFolder.Files[$fileName]
        if($file -ne $null)
        {
            Output "[UploadFileToSharePointFolder] The file $fileName already exists. Delete it first and then create a new one." "Yellow"
        }
        
        # Get the stream that contains the file contents.
        $fileItem = Get-Item $fileLocalPath
        $spFolder.Files.Add($fileName, $fileItem.OpenRead(), $overwrite) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
        $addedFile = $spFolder.Files[$fileName]        
    }
    else
    {
        Throw "The folder '$folderUrl' does not exist."
    }

    Output "[UploadFileToSharePointFolder] The file '$fileName' has been uploaded to the folder '$folderUrl'." "Green"
}

#-----------------------------------------------------------------------------------
# <summary>
# Get a Microsoft.SharePoint.SPRoleDefinition object which represents a permission level.
# </summary>
# <param name="permissionLevelName">The name of the permission level.</param>
# <param name="permissionDetail">A string array that specifies permissions, such
# as ""Open","ViewListItems","ViewPages","UseRemoteAPIs","UseClientIntegration"".</param>
# <returns>A Microsoft.SharePoint.SPRoleDefinition object represents the permission level.</returns>
#-----------------------------------------------------------------------------------
Function GetPermissionObject
{
    Param(
    [String]$permissionLevelName,
    [System.Array]$permissionDetail
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($permissionLevelName -eq $null -or $permissionLevelName -eq "")
    {
        Throw "The parameter permissionLevelName cannot be empty."
    }

    if($permissionDetail -eq $null -or $permissionDetail -eq "")
    {
        Throw "The parameter permissionDetail cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Create new customized permission level definition.
    #----------------------------------------------------------------------------
    $newLevel = new-object Microsoft.SharePoint.SPRoleDefinition
    $newLevel.Name = $permissionLevelName
    $newLevel.Description = $permissionLevelName

    #----------------------------------------------------------------------------
    # Bind detailed permissions to customized permission level.
    #----------------------------------------------------------------------------
    $basePermissions = "AddListItems","ViewListItems","EditListItems","DeleteListItems","ApproveItems","OpenItems","ViewVersions",
    "DeleteVersions","CancelCheckout","ManagePersonalViews","ManageLists" ,"ViewFormPages" ,"Open","ViewPages","AddAndCustomizePages",
    "ApplyThemeAndBorder","ApplyStyleSheets","ViewUsageData","CreateSSCSite","CreateSSCSite","ManageSubwebs","CreateGroups","ManagePermissions",
    "BrowseDirectories","BrowseUserInfo","AddDelPrivateWebParts","UpdatePersonalWebParts","ManageWeb","UseClientIntegration","UseRemoteAPIs",
    "ManageAlerts","CreateAlerts","EditMyUserInfo","EnumeratePermissions","FullMask"
                       
    
    for($index = 0 ; $index -lt $permissionDetail.Length ; $index++)
    {
        $findMatchPermission = $false
        foreach ($permission in $basePermissions)
        {
            if($permission -eq $permissionDetail[$index])
            {
                $newLevel.BasePermissions = $newLevel.BasePermissions -bor [Microsoft.SharePoint.SPBasePermissions]::$permission
                $findMatchPermission = $true;
                break;
            }
        }
        if($findMatchPermission -eq $false)
        {
            throw "No such permission level : $permissionDetail[$index]"
        }
    }
    
    Output "The required permission level details are <$permissionDetail>." "White"
    return $newLevel 
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a customized permission level which specifics the built-in permissions
# available in SharePoint Server. The new permission level will be added to the root site
# of the site collection.
# </summary>
# <param name="siteCollection">The site collection of which the root site to be added permission.</param>
# <param name="permissionLevel">The name of the permission level.</param>
# <param name="permissionDetail">A string array that specifies permissions, such
# as "Open", "ViewListItems", "ViewPages" ,"UseRemoteAPIs", "UseClientIntegration".</param>
# <param name="force">Boolean value, true means that if the specified permission level exists,
# remove it and then create a new one; otherwise false.</param>
#-----------------------------------------------------------------------------------
function CreatePermissionLevel
{
    Param(
    [Microsoft.SharePoint.SPSite]$siteCollection,
    [string]$permissionLevel,
    [System.Array]$permissionDetail,
    [Bool]$force = $false
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($siteCollection -eq $null -or $siteCollection -eq "")
    {
        Throw "The parameter siteCollection cannot be empty."
    }
    if($permissionLevel -eq $null -or $permissionLevel -eq "")
    {
        Throw "The parameter permissionLevel cannot be empty."
    }
    if($permissionDetail -eq $null -or $permissionDetail -eq "")
    {
        Throw "The parameter permissionDetail cannot be empty."
    }

    #----------------------------------------------------------------------------
    # Retrieve required web and definitions.
    #----------------------------------------------------------------------------
    try
    {
        $web = $siteCollection.RootWeb
        $definition = $web.RoleDefinitions

        #----------------------------------------------------------------------------
        # Make sure if there is existing permission level.
        #----------------------------------------------------------------------------
        $isExist = $false
        foreach($existLevel in $definition)
        {
            if($permissionLevel -eq $existLevel.Name)
            {
                $isExist = $true
                break
            }
        }

        #----------------------------------------------------------------------------
        # Create a permission level.
        #----------------------------------------------------------------------------
        $roleDefinition = GetPermissionObject $permissionLevel $permissionDetail
        if($isExist -eq $true)
        {
            if($force -eq $true)
            {
                # Remove existing permission level.
                $web.RoleDefinitions.Delete($permissionLevel)
                Output "[CreatePermissionLevel] The permission level named $permissionLevel already exists, and it has been removed." "Yellow"
                
                # Add the new permission level to the web.
                $web.RoleDefinitions.Add($roleDefinition)
                Output "[CreatePermissionLevel] The specific permission level is created successfully." "Green"
            }
            else
            {
                Output "[CreatePermissionLevel] The permission level already exists on the site." "Yellow"
            }
        }
        else
        {
            # Add the new permission level to the web.
            $web.RoleDefinitions.Add($roleDefinition)
            Output "[CreatePermissionLevel] The specific permission level has been created." "Green"
        }
    }
    finally
    {
        if($web -ne $null)
        {
            $web.Dispose()
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Grant user with specific permission level. The user will be added to a
# web site of the site collection if not exists.
# </summary>
# <param name="web">The web to add the user to.</param>
# <param name="permissionLevel">The name of the permission level.</param>
# <param name="domainName">The domain name of the user granted permissions.</param>
# <param name="userName">The name of the user granted permissions.</param>
#-----------------------------------------------------------------------------------
function GrantUserPermission
{
    Param(
    [Microsoft.SharePoint.SPWeb]$web,
    [string] $permissionLevel,
    [string] $domainName,
    [string] $userName
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if($permissionLevel -eq $null -or $permissionLevel -eq "")
    {
        Throw "The parameter permissionLevel cannot be empty."
    }
    if($domainName -eq $null -or $domainName -eq "")
    {
        Throw "The parameter domainName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Check if the permission level exists.
    #----------------------------------------------------------------------------
    $defaultPermissionLevel = @("Full Control","Design","Contribute","Read","View Only","Limited Access","Edit")    
    switch($permissionLevel)
    {        
        "Full Control"
        {
            $permissionList = "FullMask"
            break
        }
        "Design"
        {
            $permissionList = @("ViewListItems", "AddListItems", "EditListItems", "DeleteListItems", "ApproveItems", "OpenItems", "ViewVersions","DeleteVersions", "CancelCheckout",
            "ManagePersonalViews", "ManageLists", "ViewFormPages", "Open", "ViewPages", "AddAndCustomizePages","ApplyThemeAndBorder", "ApplyStyleSheets", "CreateSSCSite", "BrowseDirectories",
            "BrowseUserInfo", "AddDelPrivateWebParts","UpdatePersonalWebParts", "UseClientIntegration", "UseRemoteAPIs", "CreateAlerts", "EditMyUserInfo")
            break
        }
        "Contribute"
        {
            $permissionList = @("ViewListItems", "AddListItems", "EditListItems", "DeleteListItems", "OpenItems", "ViewVersions","DeleteVersions", "ManagePersonalViews", "ViewFormPages", 
            "Open", "ViewPages", "CreateSSCSite", "BrowseDirectories", "BrowseUserInfo", "AddDelPrivateWebParts", "UpdatePersonalWebParts","UseClientIntegration", "UseRemoteAPIs",
            "CreateAlerts", "EditMyUserInfo")
            break
        }
        "Read"
        {
            $permissionList = @("ViewListItems", "OpenItems", "ViewVersions", "ViewFormPages", "Open", "ViewPages", "CreateSSCSite","BrowseUserInfo", "UseClientIntegration", "UseRemoteAPIs", "CreateAlerts")
            break
        }
        "View Only"
        {
            $permissionList = @("ViewListItems", "ViewVersions", "ViewFormPages", "Open", "ViewPages", "CreateSSCSite", "BrowseUserInfo","UseClientIntegration", "UseRemoteAPIs", "CreateAlerts")
            break
        }
        "Limited Access"
        {
            $permissionList = @("ViewFormPages", "Open", "BrowseUserInfo", "UseClientIntegration", "UseRemoteAPIs")
            break
        }
        "Edit"
        {
            $permissionList = @("ViewListItems", "AddListItems", "EditListItems", "DeleteListItems", "OpenItems", "ViewVersions",
            "DeleteVersions", "ManagePersonalViews", "ManageLists", "ViewFormPages", "Open", "ViewPages", "CreateSSCSite",
            "BrowseDirectories", "BrowseUserInfo", "AddDelPrivateWebParts", "UpdatePersonalWebParts",
            "UseClientIntegration", "UseRemoteAPIs", "CreateAlerts", "EditMyUserInfo")
            break
        }
    
    }
    $permissionLevelExsit = $web.RoleDefinitions | ?{$_.Name -eq "$permissionLevel"}
    if(!$permissionLevelExsit)
    {
        if($defaultPermissionLevel -Contains("$permissionLevel"))
        {
            CreatePermissionLevel $web.site $permissionLevel $permissionList $false
        }
        else
        {
            Throw "The permission level `"$permissionLevel`" does not exist on the web " + $web.Url
        }        
    }
    
    #----------------------------------------------------------------------------
    # Check if the specific user exists. If not, add the user.
    #----------------------------------------------------------------------------
    $loginName = $domainName + "\" + $userName
    $webApplicationUri = new-object System.Uri("http://localhost")
    $webApplication = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($webApplicationUri)
    $userAlreadyExist = $false
    foreach($user in $web.SiteUsers)
    {
        if($user.LoginName -eq $loginName)
        {
            $userAlreadyExist = $true
            break
        }
    }
    if(!$userAlreadyExist)
    {
        $web.SiteUsers.Add($loginName, $null, $userName, $null)
    }
    
    #----------------------------------------------------------------------------
    # Grant a user with specific permission level.
    #----------------------------------------------------------------------------
    $roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($loginName, "$userName@$domainName.com", $userName, "")
    $collRoleDefinitionBindings = new-object Microsoft.SharePoint.SPRoleDefinitionBindingCollection
    $collRoleDefinitionBindings = $roleAssignment.RoleDefinitionBindings
    $collRoleDefinitionBindings.Add($web.RoleDefinitions[$permissionLevel])
    $web.RoleAssignments.Add($roleAssignment)
    Output "[GrantUserPermission] The user $loginName has been granted with the permission: $permissionLevel" "Green"
    
    if($webApplication.UseClaimsAuthentication)
    {
        $loginName = "i:0#.w|" + $domainName + "\" + $userName
        $userAlreadyExist = $false
        foreach($user in $web.SiteUsers)
        {
            if($user.LoginName -eq $loginName)
            {
                $userAlreadyExist = $true
                break
            }
        }
        if(!$userAlreadyExist)
        {
            $web.SiteUsers.Add($loginName, $null, $userName, $null)
        }
        
        #----------------------------------------------------------------------------
        # Grant a user with specific permission level.
        #----------------------------------------------------------------------------
        $roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($loginName, "$userName@$domainName.com", $userName, "")
        $collRoleDefinitionBindings = new-object Microsoft.SharePoint.SPRoleDefinitionBindingCollection
        $collRoleDefinitionBindings = $roleAssignment.RoleDefinitionBindings
        $collRoleDefinitionBindings.Add($web.RoleDefinitions[$permissionLevel])
        $web.RoleAssignments.Add($roleAssignment)
        Output "[GrantUserPermission] The user $loginName has been granted with the permission: $permissionLevel" "Green"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Grant group with specific permission level. 
# </summary>
# <param name="web">The web to add the group to.</param>
# <param name="permissionLevel">The name of the permission level.</param>
# <param name="groupName">The name of the group.</param>
#-----------------------------------------------------------------------------------
function GrantGroupPermission
{
    Param(
    [Microsoft.SharePoint.SPWeb] $web,
    [string] $permissionLevel,
    [string] $groupName
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if($permissionLevel -eq $null -or $permissionLevel -eq "")
    {
        Throw "The parameter permissionLevel cannot be empty."
    }
    if($groupName -eq $null -or $groupName -eq "")
    {
        Throw "The parameter groupName cannot be empty."
    }
     
    #----------------------------------------------------------------------------
    # Grant a user with specific permission level.
    #----------------------------------------------------------------------------
    $group = $web.SiteGroups[$groupName]
    $role = $web.RoleDefinitions[$permissionLevel]
    $roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($group)
    $collRoleDefinitionBindings = new-object Microsoft.SharePoint.SPRoleDefinitionBindingCollection
    $roleAssignment.RoleDefinitionBindings.Add($role)
    $web.RoleAssignments.Add($roleAssignment)
    
    Output "[GrantGroupPermission] The group $groupName has been granted with the permission: $permissionLevel" "Green"
}

#-----------------------------------------------------------------------------------
# <summary>
# Create list item on the specified site.
# </summary>
# <param name="web">The web site where the list will be created.</param>
# <param name="listName">The name of the list item to be created.</param>
# <param name="type">The type of a list definition or a list template, the values are 
# defined in Microsoft.SharePoint.SPListTemplateType.</param>
#-----------------------------------------------------------------------------------
function CreateListItem
{
    Param(
    [Microsoft.SharePoint.SPWeb]$web,
    [String]$listName,
    [Int]$type
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if($listName -eq $null -or $listName -eq "")
    {
        Throw "The parameter listName cannot be empty."
    }
    if($type -eq $null -or $type -eq "")
    {
        Throw "The parameter type cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Create a list item under the specified web.
    #----------------------------------------------------------------------------
    $listItem = $web.Lists[$listName]
    
    if ($listItem -eq $null -or $listItem -eq "")
    {
        $web.Lists.Add($listName, "", $type) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
        Output ("[CreateListItem] The $listName has been created under the web """ + $web.Url + """") "Green" 
    }
    else
    {
        output "[CreateListItem] The list item already exists." "Yellow" 
    }  
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a site collection with specified name on the SharePoint Server.
# </summary>
# <param name="siteCollectionName">The name of the site collection.</param>
# <param name="computerName">The computer name of the SharePoint Server.</param>
# <param name="ownerName">The name of the owner of the site collection.</param>
# <param name="ownerMail">The email address of the owner.</param>
# <param name="webTemplate">A String that specifies the site definition or site template for the site object.</param>
# <param name="nlcid">An unsigned 32-bit integer that specifies the language code identifier (LCID) for the site object.</param>
# <returns>A Microsoft.SharePoint.SPSite object of the site collection.</returns>
#-----------------------------------------------------------------------------------
function CreateSiteCollection
{
    Param(
    [String]$siteCollectionName,
    [String]$computerName,
    [String]$ownerName,
    [String]$ownerMail,
    [String]$webTemplate,
    [Uint32]$nlcid
    )
     
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($ownerName -eq $null -or $ownerName -eq "")
    {
        Throw "The parameter ownerName cannot be empty."
    }
    if($ownerMail -eq $null -or $ownerMail -eq "")
    {
        Throw "The parameter ownerMail cannot be empty."
    }
     
    if($siteCollectionName -eq $null -or $siteCollectionName -eq "")
    {
        $defaultSiteCollection = new-object Microsoft.SharePoint.SPSite ("http://$computerName/$siteCollectionName")
        Output "[CreateSiteCollection] Use the default site collection http://$computerName, as no site collection is created." "Yellow"
        return $defaultSiteCollection
    }
    else
    {
        $siteCollectionServerRelativeURL = "sites/$siteCollectionName"
        $webApplicationUri = new-object System.Uri("http://$computerName")
        $webApplication = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($webApplicationUri)
        $siteCollections = $webApplication.Sites
        $siteCollection = $siteCollections[$siteCollectionServerRelativeURL]
        if ($siteCollection -ne $null)
        {
            Output "The site collection $siteCollectionName already exists. Delete it first and then create a new one." "Yellow"
            $siteCollections.Delete($siteCollectionServerRelativeURL)
        }
        if($webTemplate -eq $null -or $webTemplate -eq "")
        {
            $siteCollection = $siteCollections.add($siteCollectionServerRelativeURL, $ownerName, $ownerMail)
        }
        else
        {
            $siteCollection = $siteCollections.add($siteCollectionServerRelativeURL, $siteCollectionName, $null, $nlcid, $webTemplate, $ownerName, $ownerName, $ownerMail)
        }
        Output "[CreateSiteCollection] The site collection $siteCollectionName has been created, the URL is:" "Green"
        Output ("[CreateSiteCollection] " + $siteCollection.Url) "Green"
        return $siteCollection
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Enable server to allow scripts in custom pages to run on server side.
# </summary>
# <param name="siteUrl">The absolute url of the site.</param>
#-----------------------------------------------------------------------------------
function EnableServerSideScriptForCustomPages
{
    Param(
    [Microsoft.SharePoint.SPSite]$siteUrl
    )
     
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }  
       
    try
    {   
        $configFilePath = $siteUrl.WebApplication.IisSettings[[ Microsoft.SharePoint.Administration.SPUrlZone]::Default].Path.ToString()
        $configFile = $configFilePath + "/web.config"
        if(Test-Path $configFile)
        {
            [xml]$configContent = Get-Content "$configFile"
            $node = $configContent.GetElementsByTagName("PageParserPaths") 
            $path = $node.item(0)
            $childNodes = $path.GetElementsByTagName("PageParserPath")
            $serverSideScriptEnabled = $false
            foreach($childNode in $childNodes)
            {
                if($childNode.getAttribute("VirtualPath") -eq "/*" -and $childNode.getAttribute("CompilationMode") -eq "Always" -and 
                $childNode.getAttribute("AllowServerSideScript") -eq "true" -and $childNode.getAttribute("IncludeSubFolders") -eq "true")
                {
                    $serverSideScriptEnabled = $true
                    break
                }
            }
            if(!$serverSideScriptEnabled)
            {
                $enableNode = $configcontent.createelement("PageParserPath")
                $enableNode.SetAttribute("VirtualPath","/*")
                $enableNode.SetAttribute("CompilationMode","Always")
                $enableNode.SetAttribute("AllowServerSideScript","true")
                $enableNode.SetAttribute("IncludeSubFolders","true")
                $path.AppendChild($enableNode) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
                $configContent.save($configFile)
            }
            Output "[EnableServerSideScriptForCustomPages] The server is enabled to run the user custom script." "Green"
        }
        else
        {
            Throw "[EnableServerSideScriptForCustomPages] Can't find the file: $configFile."
        }
    }
    finally
    {
        if($site -ne $null)
        {
            $site.Dispose()
        }
    }
    
}


#-----------------------------------------------------------------------------------
# <summary>
# Add a group to the specific site collection.
# </summary>
# <param name="siteCollection">The site collection in which to add the group.</param>
# <param name="ownerName">The owner name of the group. The name is in the format DOMAIN\User_Alias.</param>
# <param name="groupName">The name of the group.</param>
#-----------------------------------------------------------------------------------
function AddGroupForSiteCollection
{
    Param(
    [Microsoft.SharePoint.SPSite]$siteCollection,
    [String]$ownerName,
    [String]$groupName
    )
     
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($siteCollection -eq $null -or $siteCollection -eq "")
    {
        Throw "The parameter siteCollection cannot be empty."
    }
    if($groupName -eq $null -or $groupName -eq "")
    {
        Throw "The parameter groupName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Check if the specific user exists. If not, add the user. 
    #----------------------------------------------------------------------------
    $loginName = $ownerName
    $webApplicationUri = new-object System.Uri("http://localhost")
    $webApplication = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($webApplicationUri)
    if($webApplication.UseClaimsAuthentication)
    {
        $loginName = "i:0#.w|" + $ownerName
    }
    $rootWeb = $siteCollection.RootWeb
    try
    {  
        $userAlreadyExist = $false
        foreach($user in $rootWeb.SiteUsers)
        {
            if($user.LoginName -eq $loginName)
            {
                $userAlreadyExist = $true
                break
            }
        }
        if(!$userAlreadyExist)
        {
            $rootWeb.SiteUsers.Add($loginName, $null, $ownerName, $null)
        }
        
        $groups = $rootWeb.SiteGroups
        $user = $rootWeb.SiteUsers[$loginName]
        if($user -eq $null -or $user -eq "")
        {
            Throw "The user $ownerName is not a member of the site collection."
        }
        else
        {
            if($groups[$groupName] -ne $null)
            {
                Output "[AddGroupForSiteCollection] The group $groupName already exists. Delete it first and then create a new one." "Yellow"
                $groups.Remove($groupName)
            }
            $groups.Add($groupName,$user,$user,$groupName)
            Output "[AddGroupForSiteCollection] The group $groupName has been created." "Green"
        }
    }
    finally
    {
        if($rootWeb -ne $null)
        {
            $rootWeb.Dispose()
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a list workflow association and add it to the specified list. 
# </summary>
# <param name="siteUrl">The absolute url of the site.</param>
# <param name="listName">The title of the list.</param>
# <param name="workflowName">The name of the workflow to be created.</param>
# <param name="workflowTemplate">The name of the workflow template to be used.</param>
# <param name="workflowTaskList">The list where the workflow tasks will be created.</param>
# <param name="workflowHistoryList">The list where the workflow history events will be logged.</param>
# <param name="overwrite">If overwrite is false, this script will throw an exception when  
# a workflow with the same name already exists. If overwrite is true, and the workflow already 
# exists, the old workflow will be removed. A new one will be created and added to the list. 
# By default, this param is set to true.</param>
#-----------------------------------------------------------------------------------
function AddListWorkFlow
{
    Param ( 
      [Microsoft.SharePoint.SPSite]$siteUrl,
      [string]$listName,
      [string]$workflowName,
      [string]$workflowTemplate,
      [string]$workflowTaskList,
      [string]$workflowHistoryList,
      [bool]$overwrite = $true
      )
      
      #----------------------------------------------------------------------------
      # Validate parameter.
      #----------------------------------------------------------------------------
      if ($siteUrl -eq $null -or $siteUrl -eq "")
      {
          Throw "The parameter siteUrl cannot be empty."
      }
      if ($listName -eq $null -or $listName -eq "")
      {
          Throw "The parameter listName cannot be empty."
      }
      if ($workflowName -eq $null -or $workflowName -eq "")
      {
          Throw "The parameter workflowName cannot be empty."
      }
      if ($workflowTemplate -eq $null -or $workflowTemplate -eq "")
      {
         Throw "The parameter workflowTemplate cannot be empty."
      }
      # Parameter '$workflowTaskList' and '$workflowHistoryList' may be null depending on the 
      # requirements of the workflow template.
      
      #----------------------------------------------------------------------------
      # Get the specified list.
      #----------------------------------------------------------------------------
      $web =$siteUrl.RootWeb
      $list = $web.Lists[$listName]
      if ($list -eq $null -or $list -eq "")
      {
          Throw "Cannot find the list '$listName' in the '$siteUrl' site."
      }
                  
      #----------------------------------------------------------------------------
      # Get the specified workflow template.
      #----------------------------------------------------------------------------
      $wfTemplate = $web.WorkflowTemplates | where {$_.Name -eq $workflowTemplate}
      if ($wfTemplate -eq $null -or $wfTemplate -eq "")
      {
          Throw "Cannot find the specified workflow template '$workflowTemplate' in $siteUrl."
      }

      #----------------------------------------------------------------------------
      # Get the workflow tasks list if specified.
      #----------------------------------------------------------------------------
      $wfTasks = $null
      if ($workflowTaskList -ne $null -and $workflowTaskList -ne "")
      {
          $wfTasks = $web.Lists[$workflowTaskList]
          if ($wfTasks -eq $null -or $wfTasks -eq "")
          {
              Throw "Cannot find the specified workflow tasks list '$workflowTaskList' in $siteUrl."
          }
      }

      #----------------------------------------------------------------------------
      # Get the workflow history list if specified.
      #----------------------------------------------------------------------------
      $wfHistory = $null
      if ($workflowHistoryList -ne $null -and $workflowHistoryList -ne "")
      {
          $wfHistory = $web.Lists[$workflowHistoryList]
          
          if ($wfHistory -eq $null -or $wfHistory -eq "")
          {
              Throw "Cannot find the specified workflow history list '$workflowHistoryList' in $siteUrl."
          }
      }

      #----------------------------------------------------------------------------
      # Create the workflow.
      #----------------------------------------------------------------------------
      $wfAssociation = [Microsoft.SharePoint.Workflow.SPWorkflowAssociation]::CreateListAssociation($wfTemplate, $workflowName, $wfTasks, $wfHistory)
      $wfAssociation.AutoStartCreate = $true
      #----------------------------------------------------------------------------
      # Check if a workflow with the same name already exists in the list.
      #----------------------------------------------------------------------------
      $wfOriginalAssociation = $list.WorkflowAssociations | where {$_.Name -eq $workflowName}
      if ($wfOriginalAssociation -ne $null)
      {
          if ($overwrite)
          {
              $list.RemoveWorkflowAssociation($wfOriginalAssociation) | Out-Null
              Output "An existing workflow with the name '$workflowName' is removed." "Yellow"
          }
          else
          {
              Throw "Failed to add a workflow. The workflow '$workflowName' already exists on the '$listName' list."
          }
      }
      #----------------------------------------------------------------------------
      # Add the workflow to the list.
      #----------------------------------------------------------------------------
      $list.AddWorkflowAssociation($wfAssociation) | Out-Null
      Output "A workflow named '$workflowName' is created for the list '$listName' in $siteUrl." "Green"
}

#-----------------------------------------------------------------------------------
# <summary>
# Modify the value of the specific node in the specific XML file.
# </summary>
# <param name="sourceFileName">The name of the XML file containing the node.</param>
# <param name="specAttributeValue">The value content of specific attribute in the node.</param>
# <param name="modifyAttributeValue">The value content is for the attribute which will be modified in the node.</param>
# <param name="nodeName">The name of the node.</param>
# <param name="specAttributeName">The name of the specific attribute.</param>
# <param name="modifyAttributeName">The name is for the attribute which will be modified in the node.</param>
#-----------------------------------------------------------------------------------
function ModifyXMLFileNode
{
    Param(
    [string]$sourceFileName, 
    [string]$specAttributeValue, 
    [string]$modifyAttributeValue,
    [string]$nodeName,
    [string]$specAttributeName,
    [string]$modifyAttributeName
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if ($sourceFileName -eq $null -or $sourceFileName -eq "")
    {
        Throw "The parameter sourceFileName is required."
    }
    if ($specAttributeValue -eq $null -or $specAttributeValue -eq "")
    {
        Throw "The parameter specAttributeValue is required."
    }
    if ($modifyAttributeValue -eq $null -or $modifyAttributeValue -eq "")
    {
        Throw "The parameter modifyAttributeValue is required."
    }

    #----------------------------------------------------------------------------
    # Modify the content of the node.
    #----------------------------------------------------------------------------
    $isFileAvailable = $false
    $isNodeFound = $false

    $isFileAvailable = Test-Path $sourceFileName
    if($isFileAvailable -eq $true)
    {    
        [xml]$configContent = Get-Content $sourceFileName
        $PropertyNodes = $configContent.GetElementsByTagName($nodeName)
        foreach($node in $PropertyNodes)
        {
            if($node.GetAttribute($specAttributeName) -eq $specAttributeValue)
            {
                $node.SetAttribute($modifyAttributeName,$modifyAttributeValue)
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
            Throw "Failed while changing the file $sourceFileName : Could not find node with the attribute $specAttributeValue." 
        }
    }
    else
    {
        Throw "Failed while changing the file $sourceFileName : it does not exist."
    }

    #----------------------------------------------------------------------------
    # Verify the result after changing the file $sourceFileName.
    #----------------------------------------------------------------------------
    if($isFileAvailable -eq $true -and $isNodeFound)
    {
        [xml]$configContent = Get-Content $sourceFileName
        $PropertyNodes = $configContent.GetElementsByTagName($nodeName)
        foreach($node in $PropertyNodes)
        {
            if($node.GetAttribute($specAttributeName) -eq $specAttributeValue)
            {
                if($node.GetAttribute($modifyAttributeName) -eq $modifyAttributeValue)
                {
                    Output "Configuration success: Set the value $specAttributeValue to $modifyAttributeValue" "Green"
                    return
                }
            }
        }
        
        Throw "Failed after changing the file $sourceFileName : The actual value of the node is not same as the updated content value."
    }
}


#-----------------------------------------------------------------------------------
# <summary>
# Modify the user's email address.
# </summary>
# <param name="userName">The user name whose email address will be modified.</param>
# <param name="email">The email address.</param>
# <param name="siteUrl">The site url that the user belongs to.</param>
#-----------------------------------------------------------------------------------
function ConfigSPUserEmail
{
    Param(
    [string]$userName,
    [string]$email,
    [Microsoft.SharePoint.SPSite]$siteUrl
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if ($email -eq $null -or $email -eq "")
    {
        Throw "The parameter email cannot be empty."
    }
    if ($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Update the email address.
    #----------------------------------------------------------------------------
    $web = $siteUrl.OpenWeb()
    if ($siteUrl.WebApplication.UseClaimsAuthentication) 
    { 
        $claimName = "i:0#.w|" + $userName
        $user = $web.AllUsers[$claimName]
    } 
    else 
    { 
        $user = $web.AllUsers[$userName]
    }
    
    if ($user -ne $null) 
    { 
        $user.Email = $email
        $user.Update()
        Output "Update the email address of user $userName successfully." "Yellow"
    }    
    else
    {
        Throw "The user $userName does not exist."
    }
}

#----------------------------------------------------------------------------
# <summary>
# Add a hold in the specific site.
# </summary>
# <param name="siteUrl">The URL of a site</param>
# <param name="holdName">The name of a hold</param>
# <param name="overWrite">Boolean value, true means overwrite the hold with the same name; otherwise false.</param>
#----------------------------------------------------------------------------
function AddHolds
{
    Param(
    [string]$siteUrl,
    [string]$holdName,
    [bool]$overWrite = $true
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if($holdName -eq $null -or $holdName -eq "")
    {
        Throw "The parameter holdName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Check if the hold exists.
    #----------------------------------------------------------------------------
    $spsite = New-Object Microsoft.SharePoint.SPSite($siteUrl)
    $spWeb = $spsite.OpenWeb()
    $spSiteUrl = $spWeb.ParentWeb.Url
    $spList = $spWeb.Lists["Holds"]
    foreach ($tempItem in $spList.items)
    {
        if ($tempItem["Title"] -eq $holdName)
        {
            $removeId = $tempItem["ID"]
            if($overwrite)
            {
                [Microsoft.Office.RecordsManagement.Holds.Hold]::RemoveHold($removeId, $spWeb)
                $spList.Update()
                Output "The existing hold $holdName has been deleted."
            }
            else
            {
                Throw "The holds $holdName already exists."
            }           
        }
    }

    #----------------------------------------------------------------------------
    # Create a new Hold.
    #----------------------------------------------------------------------------
    $newItem = $spList.Items.Add()
    $newItem["Title"] = $holdName
    $newItem.Update()
    $spList.Update()
    $spWeb.Dispose()
    Output "The hold $holdName has been created." 
}

#----------------------------------------------------------------------------
# <summary>
# Create a document set on a site.
# </summary>
# <param name="siteUrl">The absolute url of the site where the document set should be created.</param>
# <param name="libraryName">The name of the library where the document set should be created.</param>
# <param name="documentSetName">The document set name.</param>
# <param name="overWrite">Boolean value, true means overwrite the document set with the same name; otherwise false.</param>
#----------------------------------------------------------------------------
function CreateDocumentSet
{

    Param ( 
    [string]$siteUrl,
    [string]$libraryName,
    [string]$documentSetName,
    [bool]$overWrite = $true
    )

    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl is required."
    }
    if ($libraryName -eq $null -or $libraryName -eq "")
    {
        Throw "The parameter libraryName is required."
    }
    if ($documentSetName -eq $null -or $documentSetName -eq "")
    {
        Throw "The parameter documentSetName is required."
    }

    #----------------------------------------------------------------------------
    # Check whether the document set already exists.
    #----------------------------------------------------------------------------
    $spsite = New-Object Microsoft.SharePoint.SPSite($siteUrl)
    $spWeb = $spsite.OpenWeb()
    $spList = $spWeb.Lists[$libraryName]
    $itemCounts = $spList.ItemCount

    $contentType = $spList.ContentTypes["Document Set"]
    $contentTypeId = $contentType.Id.ToString()

    for($index = 0 ; $index -lt $itemCounts ; $index++)
    {
        $tempItem = $spList.Items[$index]
        $tempContentTypeId = $tempItem.ContentTypeId.ToString()
        if ($documentSetName -eq $tempItem.Name -and $tempContentTypeId -eq $contentTypeId)
        {
            if ($overWrite -eq $true)
            {
                $tempItem.Delete()
                $spList.Update()
                Output "The old document set `"$documentSetName`" has been deleted."
            }
            else
            {
                Throw "$documentSetName already exists."
            }
        }    
    }
    
    #----------------------------------------------------------------------------
    # Create the content type.
    #----------------------------------------------------------------------------
    $hashTable = new-object System.Collections.Hashtable
    $documentSet = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSet]::Create($spList.RootFolder,$documentSetName,$spList.ContentTypes["Document Set"].id,$hashTable)
    $spList.Update()

    Output  "A new document set $documentSetName is created in $siteUrl/$libraryName."

}

#----------------------------------------------------------------------------
# <summary>
# Add a specified content type to an existing list.
# </summary>
# <param name="siteUrl">The url of the specific site.</param>
# <param name="listName">The name of the list.</param>
# <param name="contentTypeName">The name of the content type.</param>
# <param name="overWrite">Boolean value, true means overwrite the content type in the list with the same name; otherwise false.</param>
#----------------------------------------------------------------------------
function AddContentTypeToList
{
    Param ( 
    [string]$siteUrl,
    [string]$listName,
    [string]$contentTypeName,
    [bool]$overWrite = $true
    )

    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if ($listName -eq $null -or $listName -eq "")
    {
        Throw "The parameter listName cannot be empty."
    }
    if ($contentTypeName -eq $null -or $contentTypeName -eq "")
    {
        Throw "The parameter contentTypeName cannot be empty."
    }   
    
    #----------------------------------------------------------------------------
    # Add a specified content type to an exiting list.
    #----------------------------------------------------------------------------
    $spsite = New-Object Microsoft.SharePoint.SPSite($siteUrl)
    $spWeb = $spsite.OpenWeb()

    $spContentType=$spWeb.AvailableContentTypes["$contentTypeName"]
    if($spContentType -eq $null -or $spContentType -eq "")
    {
        Throw "Check whether the $contentTypeName exists in the `"$siteUrl`"."
    }

    $spList=$spWeb.Lists[$listName]
    if ($spList -eq $null -or $spList -eq "")
    {
        Throw "Cannot find the $listName, please check whether the list $listName exists in the $siteUrl."
    }
        
    #Allow management of content types of the list.
    $spList.ContentTypesEnabled=$true
    $spList.Update()
        
    if (!$spList.IsContentTypeAllowed($spContentType))
    {
        Throw "The $contentTypeName is not allowed on the $listName."
    }
    elseif($spList.ContentTypes[$contentTypeName] -ne $null)
    {   
        if($overWrite)
        {
            #Directly use $spContentType. If ID as a parameter of Delete method is out of the range of valid values, it will throw an error to prompt the argument.
            #Therefore use $spList.ContentTypes[$contentTypeName]. ID is as a parameter.
            $spList.ContentTypes.Delete($spList.ContentTypes[$contentTypeName].ID)
            Output "The content type $contentTypeName is already in use on the $listName, so delete it."
            
            $spList.ContentTypes.Add($spContentType) | Out-Null
            Output "The new content type $contentTypeName has been added on the list $listName on the site $siteUrl."
        }
        else
        {
            Throw "The $contentTypeName is already in use on the $listName list"
        }
    }
    else
    {
        $spList.ContentTypes.Add($spContentType) | Out-Null
        Output "The content type $contentTypeName has been added on the list $listName on the site $siteUrl."
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a content organizer rules on a site.
# </summary>
# <param name="siteUrl">The absolute url of the site where the content organizer rules 
# should be created.</param>
# <param name="targetLibraryName">The name of the library where the content organizer rules 
# should be created.</param>
# <param name="ruleName">The name of the rule.</param>
# <param name="contentTypeName">The title of the content type to create.</param>
# <param name="conditionsDetailsArray">the value that specifies the rule's condition.</param>
# <param name="overWrite">Boolean value, true means overwrite the Content Organizer Rules with the same name; otherwise false.</param>
# <param name="rulePriority">The priority of the rule. By default, this param is set to '5'.</param>
# <param name="ruleStatus">The status of the rule. By default, this param is set to true.</param>
#-----------------------------------------------------------------------------------
function CreateContentOrganizerRules
{

    Param ( 
    [string]$siteUrl,
    [string]$targetLibraryName,
    [string]$ruleName,
    [string]$contentTypeName,
    [System.Array]$conditionsDetailsArray,
    [bool]$overWrite = $true,
    [string]$rulePriority = 5,
    [string]$ruleStatus = $true
    )

    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if ($targetLibraryName -eq $null -or $targetLibraryName -eq "")
    {
        Throw "The parameter TargetLibraryName cannot be empty."
    }
    if ($ruleName -eq $null -or $ruleName -eq "")
    {
        Throw "The parameter ruleName cannot be empty."
    }
    if ($contentTypeName -eq $null -or $contentTypeName -eq "")
    {
        Throw "The parameter ContentTypeName cannot be empty."
    }
    if ($conditionsDetailsArray -eq $null -or $conditionsDetailsArray -eq "")
    {
        Throw "The parameter conditionsDetailsArray cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Check whether the Content Organizer Rules Folder exists or not.
    #----------------------------------------------------------------------------
    $contentOrganizerRulesList = "Content Organizer Rules"
    $spsite = New-Object Microsoft.SharePoint.SPSite("$siteUrl")
    $spWeb = $spsite.OpenWeb()
    $searchedResult = $spWeb.Lists | Where {$_.Title -eq $contentOrganizerRulesList}

    if ($searchedResult -ne "" -or $searchedResult -ne $null)
    {
       $spList = $searchedResult
    }
    else
    {
        Throw "The List $contentOrganizerRulesList does not exist. The Rules couldn't be added."
    }
    #----------------------------------------------------------------------------
    # Define parameter.
    #----------------------------------------------------------------------------
    # Specify where to place content that matches this rule.
    $tempTargetLocation = $siteUrl + "/" + $targetLibraryName
    $tempCount = $tempTargetLocation.IndexOf("//")
    $tempTargetLocation = $tempTargetLocation.Substring($tempCount + 2)
    $tempCount = $tempTargetLocation.IndexOf("/")
    $targetLocation = $tempTargetLocation.Substring($tempCount)

    # Content Organizer Rule's condition detail
    $columnNumbers = $conditionsDetailsArray.Length
    $contentTypeDetails = New-Object "String[][]" $columnNumbers,3
    $conditionDetails = New-Object "String[][]" $columnNumbers,3

    #----------------------------------------------------------------------------
    # Make sure Content Organizer Rule's condition detail.
    #----------------------------------------------------------------------------
    $routingConditionProperties = $null
    for($index = 0 ; $index -lt $columnNumbers ; $index++)
    {
        $item = $conditionsDetailsArray[$index]
        $item = $item.Split(",")
        $conditionDetails[$index][0]= $item[0]
        $conditionDetails[$index][1]= $item[1]
        $conditionDetails[$index][2]= $item[2]
        $routingConditionProperties += $conditionDetails[$index][0] + ","
    }

    # Content Type details.
    for($index = 0 ; $index -lt $columnNumbers ; $index++)
    {
        $tempTitle = $conditionDetails[$index][0]
       
        foreach ($tempItem in $spList.Fields)
        {
            if ($tempItem.Title -eq $tempTitle)
            {
                Break
            }
        }
        $contentTypeDetails[$index][0] = $tempItem.Id
        $contentTypeDetails[$index][1] = $tempItem.InternalName
        $contentTypeDetails[$index][2] = $tempItem.Title
    }

    $conditionXmlArray = $null
    for($count = 0 ; $count -lt $conditionDetails.Length ; $count++)
    {
        $conditionFieldTitle = $contentTypeDetails[$count][0]+"|"+$contentTypeDetails[$count][1] +"|"+ $contentTypeDetails[$count][2]
        $conditionOperator = $conditionDetails[$count][1]
        $conditionFieldValue =$conditionDetails[$count][2]
        $conditionXmlDetail = [string] '<Condition Column="'+ $conditionFieldTitle +'" Operator="' + $conditionOperator + '" Value="' + $conditionFieldValue+ '" />'
        $conditionXmlArray += $conditionXmlDetail
    }

    $conditionXml = [string] "<Conditions>" + $conditionXmlArray+ "</Conditions>"

    #----------------------------------------------------------------------------
    # Add new rules.
    #----------------------------------------------------------------------------
    foreach ($item in $spList.items)
    {
        if ($ruleName -eq $item.Title)
        {
            if($overWrite -eq $false)
            {
                Throw "Failed to create the rule. The rule $ruleName already exists."
            }
            else
            {
                $oldItem = $spList.Items | Where {$_.Title -eq $ruleName}
                $oldItem.Delete()
                $spList.Update()
            }
        }
    }

    $newRule = $spList.Items.Add()
    $newRule["Title"] = $ruleName
    $newRule["RoutingRuleName"] = $ruleName
    $newRule["RoutingContentType"] = $contentTypeName
    $newRule["RoutingPriority"] = $rulePriority 
    $newRule["RoutingEnabled"] = $ruleStatus
    $newRule["RoutingConditions"] = $conditionXml
    $newRule["RoutingTargetLibrary"] = $targetLibraryName
    $routingConditionProperties = $routingConditionProperties.Trim(",")
    $newRule["RoutingConditionProperties"] = [string] $routingConditionProperties
    $newRule["RoutingRuleExternal"] = $false
    $newRule["RoutingTargetPath"]= $targetLocation
    $newRule.Update()
    $spList.Update()
    $spWeb.Update()

    Output "The rule $ruleName has been added in `"$siteUrl`"."
}

#-----------------------------------------------------------------------------------
# <summary>
# Activate or deactivate web feature on the specified site.
# </summary>
# <param name="siteUrl">The absolute url of the site where the SharePoint feature is installed.</param>
# <param name="featureName">The feature name will be activated.</param>
# <param name="isActive">Boolean value, true means Activated the installed SharePoint feature; otherwise false.</param>
#-----------------------------------------------------------------------------------
function SetWebFeature
{
    Param(
    [string]$siteUrl,
    [string]$featureName,
    [bool]$isActive = $true
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if($featureName -eq $null -or $featureName -eq "")
    {
        Throw "The parameter featureName cannot be empty."
    }
    $outPutParam = "Deactivate"
    if($isActive)
    {
         $outPutParam = "Activate"
    }
    
    try
    {    
        if($isActive)
        {
            Enable-SPFeature -Identity $featureName -Url $siteUrl -ErrorAction Stop
            Output "Successfully activated the installed SharePoint feature: $featureName on the site: $siteUrl." "Green"
        }
        else
        {
            Disable-SPFeature -Identity $featureName -Url $siteUrl -Confirm:$false -Force:$true -ErrorAction Stop
            Output "Deactivate the installed SharePoint feature $featureName on the site $siteUrl." "Green"

        }
    }
    catch [System.Management.Automation.ActionPreferenceStopException]
    {
        if(!($_.Exception -is [System.Data.DuplicateNameException]))
        {
            Throw "Failed to $outPutParam the feature $featureName"   
        }
        else
        {
            Output "$outPutParam the feature $featureName successfully." "Yellow" 
        }
    }
        
}

#-----------------------------------------------------------------------------------
# <summary>
# Add a field in the specified list.
# </summary>
# <param name="web">The web site where the list belongs to.</param>
# <param name="listName">The list where the filed will be added.</param>
# <param name="fieldName">The name of the filed item to be added.</param>
# <param name="type">The type of a Filed definition or a filed template, the values are 
# defined in Microsoft.SharePoint.SPFieldType.</param>
#<param name="isChoice">Boolean value, true means the filed type is Choice; otherwise false.</param>
#<param name="choiceList">The name of the filed item to be added.</param>
#<param name="defaultValue">The default value of the filed.</param>
#<param name="readOnly">this value is a flag,null means keeping default value for readonly property."true" means setting readonly properties to true,other means false</param>
#-----------------------------------------------------------------------------------
function AddFieldInList
{
    Param(
    [Microsoft.SharePoint.SPWeb]$web,
    [String]$listName,
    [String]$fieldName,
    [Int]$type,
    [bool]$isChoice = $false,
    [System.Array]$choiceList,
    [String]$defaultValue = "",
    [String]$readOnly = ""
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($web -eq $null)
    {
        Throw "The parameter web cannot be empty."
    }
    if($listName -eq $null -or $listName -eq "")
    {
        Throw "The parameter listName cannot be empty."
    }
    if($fieldName -eq $null -or $fieldName -eq "")
    {
        Throw "The parameter fieldName cannot be empty."
    }
    if($type -eq $null -or $type -eq "")
    {
        Throw "The parameter type cannot be empty."
    }
    if($isChoice)
    {
        if($choiceList -eq $null -or $choiceList -eq "")
        {
            Throw "The parameter choiceList cannot be empty."
        }
    }
    
    #----------------------------------------------------------------------------
    # Add a field to the specified list.
    #----------------------------------------------------------------------------
    $listItem = $web.Lists[$listName]
    $isExsit = $listItem.Fields["$fieldName"]
    if($isExsit)
    {
        $listItem.Fields["$fieldName"].Delete()
        $listItem.Update()  
    }  
    $listItem.Fields.Add($fieldName,$type,$false)
    if($isChoice)
    {
        $choiceList | %{$listItem.Fields[$fieldName].Choices.Add($_)}
        $defaultValue = $choiceList[0]
    }
    if($defaultValue)
    {
        $listItem.Fields[$fieldName].DefaultValue = $defaultValue
        $listItem.Fields[$fieldName].Update()
    }
    if($readOnly)
    {  
        $listItem.Fields[$fieldName].ReadOnlyField = ($readOnly -eq "true")        
    }  
    
    $listItem.Update()
    Output "The Fields $fieldName has been created in the list $listName." "Green"     
}

#------------------------------------------------------------------------
# <summary>
# Add an item in specific list.
# </summary>
# <param name="web">The web site where the list item will be added.</param>
# <param name="listName">The name of the list.</param>
# <param name="$itemTitle">The title of the item.</param>
#------------------------------------------------------------------------
function AddListItem
{
    param(
    [Microsoft.SharePoint.SPWeb]$spweb,
    [string]$listName,
    [string]$itemTitle
    )
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($spweb -eq $null -or $spweb -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if ($listName -eq $null -or $listName -eq "")
    {
        Throw "The parameter listName cannot be empty."
    }
    if ($itemTitle -eq $null -or $itemTitle -eq "")
    {
        Throw "The parameter itemTitle cannot be empty."
    }
    
    $spList = $spWeb.Lists["$listName"]
    if($spList -eq $null -or $spList -eq "")
    {
        Throw "The list `"$listName`" doesn't exist on the site, please create it first."
    }    

    $item = $spList.items.add()
    $item["Title"] = $itemTitle     
    $item.update()
    $spList.Update()
    $spWeb.Dispose()

}

#----------------------------------------------------------------------------
# <summary>
# Set the server's authentication mode.
# </summary>
# <param name="webConfigPath">The computer name of the server.</param>
# <param name="SharePointVersion">The sharePoint server version.
# Note: The value is obtained by calling function "GetSharePointVersion".</param>
# <param name="domain">The user domain name.</param>
# <param name="connectionUsername">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="connectionPassword">The password of the user name.</param>
# <param name="connectionString">The connection string from Active Directory Domain Services. The format is like LDAP:\\domain\OU=Groups,DC=domain,DC=domain suffix.</param>
# <param name="groupConnectionString">Current user container in active directory. The format is like OU=Groups,DC=domain,DC=domain suffix.</param>
# <param name="authenticationMode">the name of authentication Mode.</param>
# <param name="isOnlySetAuthenMode">A Boolean value, true means only set the mode of authentication, otherwise false.</param>
#----------------------------------------------------------------------------
function SetServerAuthenticationMode
{
    Param(
    [String]$webConfigPath,
    [String]$SharePointVersion,
    [String]$domain,
    [String]$connectionUsername,
    [String]$connectionPassword,
    [String]$connectionString,
    [String]$groupConnectionString,
    [String]$authenticationMode,
    [String]$isOnlySetAuthenMode = $true
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
       if($webConfigPath -eq $null -or $webConfigPath -eq "")
    {
        Throw "The parameter webConfigPath cannot be empty."
    }
    if($SharePointVersion -eq $null -or $SharePointVersion -eq "")
    {
        Throw "The parameter SharePointVersion cannot be empty."
    }
    if($domain -eq $null -or $domain -eq "")
    {
        Throw "The parameter domain cannot be empty."
    }
    if($connectionUsername -eq $null -or $connectionUsername -eq "")
    {
        Throw "The parameter connectionUsername cannot be empty."
    }
    if($connectionPassword -eq $null -or $connectionPassword -eq "")
    {
        Throw "The parameter connectionPassword cannot be empty."
    }
    if($connectionString -eq $null -or $connectionString -eq "")
    {
        Throw "The parameter connectionString cannot be empty."
    }
    if($groupConnectionString -eq $null -or $groupConnectionString -eq "")
    {
        Throw "The parameter groupConnectionString cannot be empty."
    }
    if($authenticationMode -eq $null -or $authenticationMode -eq "")
    {
        Throw "The parameter authenticationMode cannot be empty."
    }
    if(!([System.IO.File]::Exists($webConfigPath)))
    {
        Throw "the webconfig file $webConfigPath does not exist."
    }
    
    #Define provider in web.config.

    $mumberShip = "AspNetActiveDirectoryMembershipProvider"
    $SystemWeb = @"
    <configuration>
     <system.web>
      <membership defaultProvider="$mumberShip">
       <providers>
         <add name="$mumberShip" type="System.Web.Security.ActiveDirectoryMembershipProvider,System.Web, Version=2.0.0.0, Culture=neutral,PublicKeyToken=b03f5f7f11d50a3a" connectionStringName="ADServiceString" connectionUsername="$connectionUsername" connectionPassword="$connectionPassword" attributeMapUsername="sAMAccountName" />
       </providers>
      </membership>
     </system.web>
    </configuration>
"@       
    
    if($SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
    {
        $mumberShip = "LdapMember"
        $roleManager = "LdapRole"
        $SystemWeb = @"
        <configuration>
         <system.web>
          <membership>
           <providers>
            <add name="$mumberShip" type="Microsoft.Office.Server.Security.LdapMembershipProvider, Microsoft.Office.Server, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" server="$domain" port="389" useSSL="false" userDNAttribute="distinguishedName" userNameAttribute="sAMAccountName" userContainer="$connectionString" userObjectClass="person" userFilter="(ObjectClass=person)" scope="Subtree" otherRequiredUserAttributes="sn,givenname,cn" />
           </providers>
          </membership>
          <roleManager enabled="true" cacheRolesInCookie="false">
           <providers>
            <add name="$roleManager" type="Microsoft.Office.Server.Security.LdapRoleProvider, Microsoft.Office.Server, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" server="$domain" port="389" useSSL="false" groupContainer="$connectionString" groupNameAttribute="cn" groupNameAlternateSearchAttribute="samAccountName" groupMemberAttribute="member" userNameAttribute="sAMAccountName" dnAttribute="distinguishedName" groupFilter="(ObjectClass=group)" userFilter="(ObjectClass=person)" scope="Subtree" />
           </providers>
          </roleManager>
         </system.web>
        </configuration>
"@
    }

    [xml]$configContent = New-Object XML 
    $configContent.LoadXml($SystemWeb)
    $stsRoot = $configContent.DocumentElement
    $stsProviders = $stsRoot."system.web".membership.providers.add
    $stsProvider = $stsProviders | where {$_.name -eq $mumberShip}
     
    #Get xmlDocument object which is used to modify the web.config file and save the changes.
    $xml=[xml](Get-Content $webConfigPath)
    $root=$xml.get_DocumentElement()
    
    if($isOnlySetAuthenMode -eq $true -and $root."system.web".authentication -ne $null)
    {
        $root."system.web".authentication.mode = "$authenticationMode"
    }
    else
    {
    
        if((($SharePointVersion -ne $SharePointServer2013[0]) -and ($SharePointVersion -eq $SharePointServer2016[0])) -and $root.connectionStrings -eq $null)
        {
            $formsElement = $xml.CreateElement("connectionStrings")
            $root.appendchild($formsElement)
            $adNameElement = $xml.CreateElement("add")
            $adNameElement.SetAttribute("name","ADServiceString")
            $adNameElement.SetAttribute("connectionString","$groupConnectionString")
            $formsElement.appendchild($adNameElement) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
        }
        if(($SharePointVersion -eq $WindowsSharePointServices3[0] -or $SharePointVersion -eq $SharePointServer2007[0]) -and $root."system.web".authentication -ne $null)
        {
            $root."system.web".authentication.mode = "$authenticationMode"
        }
        if(!($root."system.web".count))
        {
            if($root."system.web".membership -eq $null)
            {
                $root.AppendChild($xml.ImportNode($stsRoot."system.web", $true))
            }
            elseif(($root."system.web".membership.providers.add | where {$_.name -eq $mumberShip}) -eq $null)
            {         
                $root."system.web".membership.providers.AppendChild($xml.ImportNode($stsProvider, $true)) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
            }
        }
        
        if($SharePointVersion -eq $SharePointServer2013[0] -or $SharePointVersion -eq $SharePointServer2016[0])
        {
            $stsRoleProviders = $stsRoot."system.web".roleManager.providers.add
            $stsRoleProvider = $stsRoleProviders | where {$_.name -eq $roleManager} 
            if($root."system.web".roleManager -eq $null)
            {
                $root."system.web".AppendChild($xml.ImportNode($stsRoot."system.web".roleManager, $true)) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
            }
            elseif(($root."system.web".roleManager.providers.add | where {$_.name -eq $roleManager}) -eq $null)
            {         
                $root."system.web".roleManager.providers.AppendChild($xml.ImportNode($stsRoleProvider, $true)) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
                $root."system.web".roleManager.SetAttribute("cacheRolesInCookie", $false)
            } 
        } 
        
    }
    $xml.Save($webConfigPath)
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the connection string from Active Directory Domain Services
# </summary>
# <returns>[string]The script returns the connection string from Active Directory Domain Services. </returns>
#-----------------------------------------------------------------------------------
function GetFQDN
{
    $userFQDN = whoami /fqdn
    $groupN = "DC="
    if($userFQDN -like "*OU=*")
    {
        $groupN = "OU="
    }
    $userFQDN_substring = $userFQDN.Substring($userFQDN.IndexOf("$groupN"))
    return $userFQDN_substring
}

#----------------------------------------------------------------------------
# <summary>
# Create WebApplication.
# </summary>
# <param name="computerName">The computer name of the machine where SharePoint Server is installed.</param>
# <param name="poolAccount">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="portNum">Specifies the port on which this Web application can be accessed.</param>
# <param name="webAppName">Specifies the name of the new Web application.</param>
# <param name="applicationPoolName">Specifies the name of an application pool to use.</param>
# <param name="password">The password of poolAccount.</param>
# <param name="SharePointVersion">The sharePoint server version.
# Note: The value is obtained by calling function "GetSharePointVersion".</param>
# <param name="isSetProvider">A Boolean value, true means add a specific claim provider to the defined Web application, otherwise false.</param>
# <param name="isAllowAnonymousAccess">A Boolean value, true means allows anonymous access to the Web application, otherwise false.</param>
# <param name="isOverWrite">Boolean value, true means overwrite the web application; otherwise false.</param>
#----------------------------------------------------------------------------
function CreateWebApplication
{
    param(
    [string]$computerName,
    [string]$poolAccount,
    [string]$portNum,
    [string]$webAppName,
    [string]$applicationPoolName,
    [string]$password,
    [string]$SharePointVersion,    
    [bool]$isSetProvider = $false,
    [bool]$isAllowAnonymousAccess = $false,
    [bool]$isOverWrite = $false
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
       if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($poolAccount -eq $null -or $poolAccount -eq "")
    {
        Throw "The parameter poolAccount cannot be empty."
    }
    if($portNum -eq $null -or $portNum -eq "")
    {
        Throw "The parameter portNum cannot be empty."
    }
    if($webAppName -eq $null -or $webAppName -eq "")
    {
        Throw "The parameter webAppName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($applicationPoolName -eq $null -or $applicationPoolName -eq "")
    {
        Throw "The parameter applicationPoolName cannot be empty."
    }
    if($SharePointVersion -eq $null -or $SharePointVersion -eq "")
    {
        Throw "The parameter SharePointVersion cannot be empty."
    }

    $memberShipProvider = "AspNetActiveDirectoryMembershipProvider"
    $roleProviderName = "temp"
    if($SharePointVersion -ge $SharePointServer2013[0])
    {        
        $memberShipProvider = "LdapMember"
        $roleProviderName = "LdapRole"
    }
        
    $snapin = Get-PSSnapin | Where-Object -FilterScript {$_.Name -eq "Microsoft.SharePoint.PowerShell"}
    if($snapin -eq $null -or $snapin -eq "")
    {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
        
    #----------------------------------------------------------------------------
    # Add user to managed account.
    #----------------------------------------------------------------------------
    $managedAccount = Get-SPManagedAccount | Where {$_.UserName -eq $poolAccount} 
    if ($managedAccount -eq $null -or $managedAccount -eq "")
    {
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PsCredential ($poolAccount,$securePassword)
        New-SPManagedAccount -Credential $Credential
        $managedAccount = Get-SPManagedAccount $userAccount
        Output "The user $poolAccount has been registered to the managed account." "Green"
    }
        
    #----------------------------------------------------------------------------
    # Create new web application.
    #----------------------------------------------------------------------------
    $targetWeb = Get-SPWebApplication -IncludeCentralAdministration | Where {$_.DisplayName -eq $webAppName }
    $isCreateWebApp = $true

    if($targetWeb -ne $null)
    {
        if($isOverWrite)
        {
            Output "The webApplication $WebAppName already exists. Delete it first and then create a new one." "Yellow"
            $webSiteAlternateUrl = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("http://${computerName}:${portNum}").AlternateUrls | where {$_.UrlZone -eq "internet"}
            if($webSiteAlternateUrl -ne $null -or $webSiteAlternateUrl -ne "")
            {   
                $number = $webSiteAlternateUrl.Uri.Port
                cmd /c "netsh http delete sslcert ipport = 0.0.0.0:$number"
            }    
            Remove-SPWebApplication $WebAppName -Confirm:$false -DeleteIISSite -removeContentDatabase            
        }
        else
        {
            $isCreateWebApp = $false
            Output "The webApplication $WebAppName already exists" "Yellow"
        }
    }
    if($isCreateWebApp)
    {
        if($isSetProvider)
        {   
            Output "Start to create a webApplication named $WebAppName..." "White"
            $ap = New-SPAuthenticationProvider -ASPNETMembershipProvider "$memberShipProvider" -ASPNETRoleProviderName "$roleProviderName"
            if($SharePointVersion -ne $SharePointServer2013[0])
            {
                $ap.RoleProvider = $null
            }
            New-SPWebApplication -Name $WebAppName -ApplicationPool $applicationPoolName -ApplicationPoolAccount $managedAccount -Port $portNum -AuthenticationProvider $ap | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
        }
        else
        {
            if(($SharePointVersion -eq $SharePointFoundation2010[0] -or $SharePointVersion -eq $SharePointServer2010[0])-and $isAllowAnonymousAccess)
            {
                 New-SPWebApplication -Name $WebAppName -ApplicationPool $applicationPoolName -AllowAnonymousAccess -ApplicationPoolAccount $managedAccount -Port $portNum -AuthenticationProvider (New-SPAuthenticationProvider) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
            }        
            else
            {
                New-SPWebApplication -Name $WebAppName -ApplicationPool $applicationPoolName -ApplicationPoolAccount $managedAccount -Port $portNum -AuthenticationProvider (New-SPAuthenticationProvider) | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
            }
        }
                 
        Output "The webApplication $WebAppName has been created." "Green"
    }

}

#----------------------------------------------------------------------------
# <summary>
# Create WebApplication.
# </summary>
# <param name="computerName">The computer name of the machine where SharePoint Server is installed.</param>
# <param name="portNum">Specifies the port on which this Web application can be accessed.</param>
# <param name="webAppName">Specifies the name of the new Web application.</param>
# <param name="webFilePath">The path of web config file.</param>
# <param name="isOverWrite">Boolean value, true means overwrite the web application; otherwise false.</param>
#----------------------------------------------------------------------------
function CreateWebApplicationOn2007
{
    param(
    [string]$computerName,
    [string]$portNum,
    [string]$webAppName,
    [string]$webFilePath,
    [bool]$isOverWrite = $false
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
       if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computer cannot be empty."
    }
    if($portNum -eq $null -or $portNum -eq "")
    {
        Throw "The parameter portNum cannot be empty."
    }
    if($webAppName -eq $null -or $webAppName -eq "")
    {
        Throw "The parameter webAppName cannot be empty."
    }
    if($webFilePath -eq $null -or $webFilePath -eq "")
    {
        Throw "The parameter webFilePath cannot be empty."
    }
    
    $uri = "http://${computerName}:${portNum}"
    $isCreateWebApp = $true
    $webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($uri)
    if($webApp -ne $null -or $webApp.Name -eq $webAppName)
    {
        if($isOverWrite)
        {
            Output "The webApplication $webAppName already exists. Delete it first and then create a new one." "Yellow"
            $webSiteAlternateUrl = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup("$uri").AlternateUrls | where {$_.UrlZone -eq "internet"}
            if($webSiteAlternateUrl -ne $null -or $webSiteAlternateUrl -ne "")
            {   
                $number = $webSiteAlternateUrl.Uri.Port
                cmd /c "netsh http delete sslcert ipport = 0.0.0.0:$number"
            }    
            $webApp.Delete()
            $webApp.Unprovision()
        }
        else
        {   
            $isCreateWebApp = $false
            Output "The webApplication $webAppName already exists." "Yellow"
        }
    }
    if($isCreateWebApp)
    {
        $spfarm = [Microsoft.SharePoint.Administration.SPfarm]::Local
        if($spfarm -ne $null -and $spfarm -ne "")
        {
            $appbuilder = new-object Microsoft.SharePoint.Administration.SPWebApplicationBuilder($spfarm)
        }
        if($appbuilder -ne $null -and $appbuilder -ne "")
        {
            $appbuilder.Port = $portNum
               $isFolderExsited = Get-ChildItem $webFilePath | ?{$_.Name -eq $portNum}
            if($isFolderExsited -ne "" -and $isFolderExsited -ne $null)
            {   
                $guid = [GUID]::NewGuid()
                $portNum = $portNum+$guid            
            }
            $webFilePath = $webFilePath + "\$portNum"
            $appbuilder.RootDirectory = New-object System.IO.DirectoryInfo ("$webFilePath")
            $appbuilder.ServerComment = $webAppName
        }
        
        while($i++ -lt 2)
        {
            Try
            {
                $webapplication = $appbuilder.Create()
                Break
            }
            Catch [System.NullReferenceException]
            {
                Continue
            }
        }
        
        $webapplication.Update()
        $webapplication.Provision()
        Output "The webApplication $WebAppName has been created." "Green"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Creates a file. 
# </summary>
# <param name="fileName">The name of the file to create.</param>
# <param name="size">The size of the file to create.</param>
# <param name="filePath">The path of the file to create.</param>
#-----------------------------------------------------------------------------------
function CreateFile
{
    Param(
    [string]$fileName,
    [double]$size,
    [string]$filePath
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($fileName -eq $null -or $fileName -eq "")
    {
        Throw "The parameter fileName cannot be empty."
    }
    if($size -eq $null -or $size -eq "")
    {
        Throw "The parameter size cannot be empty."
    }
    if($filePath -eq $null -or $filePath -eq "")
    {
        Throw "The parameter filePath cannot be empty."
    }
    
    $file = [System.IO.File]::Create("$filePath\$fileName")
    $file.SetLength($Size)
    $file.Close()
    
    Output ("The file '$fileName' has been created under """ + $filePath + """") "Green"

}
#-----------------------------------------------------------------------------------
# <summary>
# Modify the value of the specific node in the specific XML file.
# </summary>
# <param name="sourceFileName">The name of the XML file containing the node.</param>
# <param name="specAttributeValue">The value content of specific attribute in the node.</param>
# <param name="modifyAttributeValue">The value content is for the attribute which will be modified in the node.</param>
# <param name="nodeName">The name of the node.</param>
# <param name="specAttributeName">The name of the specific attribute.</param>
# <param name="modifyAttributeName">The name is for the attribute which will be modified in the node.</param>
# <param name="singleNodeName">The name of the single node.</param>
#-----------------------------------------------------------------------------------
function ModifyXMLFileSingleNode
{
    Param(
    [string]$sourceFileName, 
    [string]$specAttributeValue, 
    [string]$modifyAttributeValue,
    [string]$nodeName,
    [string]$specAttributeName,
    [string]$modifyAttributeName,
    [string]$singleNodeName
    )

    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if ($sourceFileName -eq $null -or $sourceFileName -eq "")
    {
        Throw "The parameter sourceFileName is required."
    }
    if ($specAttributeValue -eq $null -or $specAttributeValue -eq "")
    {
        Throw "The parameter specAttributeValue is required."
    }
    if ($modifyAttributeValue -eq $null -or $modifyAttributeValue -eq "")
    {
        Throw "The parameter modifyAttributeValue is required."
    }
    if ($singleNodeName -eq $null -or $singleNodeName -eq "")
    {
        Throw "The parameter singleNodeName is required."
    }

    #----------------------------------------------------------------------------
    # Modify the content of the node.
    #----------------------------------------------------------------------------
    $isFileAvailable = $false
    $isNodeFound = $false

    $isFileAvailable = Test-Path $sourceFileName
    if($isFileAvailable -eq $true)
    {    
        [xml]$configContent = Get-Content $sourceFileName
        $PropertyNodes = $configContent.GetElementsByTagName($nodeName)
        foreach($node in $PropertyNodes)
        {
            if($node.GetAttribute($specAttributeName) -eq $specAttributeValue)
            {
                $node.SelectSingleNode($singleNodeName).SetAttribute($modifyAttributeName,$modifyAttributeValue)
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
            Throw "Failed while changing the file $sourceFileName : Could not find the node with the attribute $specAttributeValue." 
        }
    }
    else
    {
        Throw "Failed while changing the file $sourceFileName : it does not exist!"
    }

    #----------------------------------------------------------------------------
    # Verify the result after changing the file $sourceFileName.
    #----------------------------------------------------------------------------
    if($isFileAvailable -eq $true -and $isNodeFound)
    {
        [xml]$configContent = Get-Content $sourceFileName
        $PropertyNodes = $configContent.GetElementsByTagName($nodeName)
        foreach($node in $PropertyNodes)
        {
            if($node.GetAttribute($specAttributeName) -eq $specAttributeValue)
            {
                if($node.SelectSingleNode($singleNodeName).GetAttribute($modifyAttributeName) -eq $modifyAttributeValue)
                {
                    Output "Configuration success: Set the value $specAttributeValue to $modifyAttributeValue" "Green"
                    return
                }
            }
        }
        
        Throw "Failed after changing the file $sourceFileName : The actual value of the node is not same as the updated content value."
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Enable the major version for list. 
# </summary>
# <param name="spWeb">The web where the list is located.</param>
# <param name="listName">The list where the major version will be enabled.</param>
#-----------------------------------------------------------------------------------
function EnableMajorVersion
{
    param(
    [Microsoft.SharePoint.SPWeb]$spWeb,
    [string]$listName
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($spWeb -eq $null -or $spWeb -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if ($listName -eq $null -or $listName -eq "")
    {
        Throw "The parameter listName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Enable Major Version.
    #----------------------------------------------------------------------------
    
    $spList=$spWeb.Lists[$listName]
    if ($spList -eq $null -or $spList -eq "")
    {
        Throw "Cannot find the $listName list, please check whether the list $listName exists on the $spWeb."
    }
    $spList.EnableVersioning = $true
    $spList.MajorVersionLimit = 10
    $spList.Update()
}

#-----------------------------------------------------------------------------------
# <summary>
# Adds a new item to the record routing . 
# </summary>
# <param name="spWeb">The web where the new item will be added in record routing.</param>
# <param name="itemName">The list where the major version will be enabled.</param>
# <param name="itemLocation">The location where the item is located.</param>
# <param name="itemMappings">The list of other file types for which the rule applies.</param>
# <param name="isDefault">Boolean value, true means this instance is to be configured as the default rule;otherwise,false.</param>
# <param name="itemDescription">the description that explains the rule.</param>
# <param name="itemRouter">The custom Router for the record routing that is added</param>
#-----------------------------------------------------------------------------------
function NewItemToRecordRouting
{
    Param ( 
    [Microsoft.SharePoint.SPWeb]$web,
    [string]$itemName,
    [string]$itemLocation,    
    [System.Array]$itemMappings,
    [bool]$isDefault = $true,
    [string]$itemDescription = "",
    [string]$itemRouter = ""    
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if ($itemName -eq $null -or $itemName -eq "")
    {
        Throw "The parameter itemName cannot be empty."
    }
    if ($itemLocation -eq $null -or $itemLocation -eq "")
    {
        Throw "The parameter itemLocation cannot be empty."
    }
    if ($itemMappings -eq $null -or $itemMappings -eq "")
    {
        Throw "The parameter itemMappings cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Adds a new rule to the record routing.
    #----------------------------------------------------------------------------
    $recSeries = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.RecordSeriesCollection -ArgumentList $web
    $isExist = $false
    $counts = $recSeries.Count;
    for($count=0;$count -lt $counts;$count++)
    {  
        if($recSeries[$count].Name -eq $itemName)
        {
            Output "The item $itemName already exists in record routing of $itemLocation." "Yellow"
            $isExist = $true            
            break
        }    
    }    
    if(!$isExist)
    {    
        $success = $true
        try
        {
            $recSeries.Add($itemName,$itemLocation,$itemDescription,$itemMappings,$itemRouter,$isDefault)
        }
        catch [Exception]
        {
            $recSeries = New-Object Microsoft.Office.RecordsManagement.RecordsRepository.RecordSeriesCollection -ArgumentList $web
            
            if($recSeries.Count -eq $counts)
            {
                $success = $false
                throw "The item" +$itemName+" cannot be created in the record routing of "+$itemLocation+"."                                            
            }
        }
        finally
        {
            $web.update()
            $recSeries.dispose()
            if($success)
            {
                Output "The item $itemName has been created to record the routing of $itemLocation." "Green"
            }
        }
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a record center web on sharepoint 2007. 
# </summary>
# <param name="$siteUrl">The site url where the web is located.</param>
# <param name="webUrl">A string that contains the new website URL relative to the root website in the site.</param>
# <remarks>
# If the template is "OFFILE#1",the method Add would always throw exception on 2007,maybe it is a product bug after investigation,current solution is to catch this exception if we can get the record center
# </remarks>
#-----------------------------------------------------------------------------------
function CreateRecordCenterOn2007
{
    Param ( 
    [string]$siteUrl,
    [string]$webUrl
    )
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if ($webUrl -eq $null -or $webUrl -eq "")
    {
        Throw "The parameter webUrl cannot be empty."
    }
    
    $parentSite = New-Object Microsoft.SharePoint.SPSite("$siteUrl")
    $web = $parentSite.OpenWeb($webUrl)
    if ($web.Exists)
    {
        OutPut "The $webUrl already exists. Delete it first and then create a new one." "Yellow"
        DeleteWeb $web
        $web.Dispose()
    }
    try
    {
        $parentSite.AllWebs.Add("$webUrl", $webUrl, "", $null, "OFFILE#1", $false, $false)
    }
    catch [Exception]
    {
        $newsite = New-Object Microsoft.SharePoint.SPSite("$siteUrl\$webUrl")
        $newWeb = $newsite.openWeb().webTemplate
        if($newsite -and ($newWeb -eq "OFFILE"))
        {
             Output ("The web ""$webUrl"" has been created under the web $siteUrl") "Green"
        }
        else
        {
            throw $_.Exception.Message
        }        
    }

}

#-----------------------------------------------------------------------------------
# <summary>
# Add a user to specific sharepoint group.
# </summary>
# <param name="siteURL">specific URL of a site</param>
# <param name="groupName">specific group name.</param>
# <param name="domainName">domain name.</param>
# <param name="userName">user name.</param>
#-----------------------------------------------------------------------------------
function AddUserToSharePointGroup
{
    param(
    [Microsoft.SharePoint.SPWeb]$web,
    [string]$groupName,
    [string]$domainName,
    [string]$userName
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($web -eq $null -or $web -eq "")
    {
        Throw "The parameter web cannot be empty."
    }
    if ($groupName -eq $null -or $groupName -eq "")
    {
        Throw "The parameter groupName cannot be empty."
    }
    if ($domainName -eq $null -or $domainName -eq "")
    {
        Throw "The parameter domainName cannot be empty."
    }
    if ($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Add user to SharePoint group 
    #----------------------------------------------------------------------------
    $spUser = $web.EnsureUser("$domainName\$userName")
    if($spUser -ne $null)
    {
        $group = $web.Groups["$groupName"]
        if($group -ne $null)
        {
            $group.AddUser($spUser) 
        }
        else
        {
            Throw "The group `"$groupName`" does not exist on the web " + $web.Url
        }
    }
    else
    {
        Throw "The user `"$userName`" does not exist on the web " + $web.Url
    }
}

#-----------------------------------------------------------------------------------------------
# <summary>
# Creates a new Secure Store application. 
# </summary>
# <param name="applicationName">Specifies the name of the new target application.</param>
# <param name="userNameType">Specifies the user name type of credential field to add to a target application.</param>
# <param name="passwordType">Specifies the password type of credential field to add to a target application.</param>
# <param name="applicationType">Specifies the type of target application.</param>
# <param name="userName">Specifies the name of the new claims principal.</param>
# <param name="password">specifies the password of the user.</param>
# <param name="domain">specifies the name of the domain.</param>
# <param name="serviceContext">The Web application that contains the specified site collection.</param>
# <param name="userCredentialItem">The "user" value of credential items which are created in Secure Store feature.</param>
# <param name="passwordCredentialItem">The "password" value of credential items which are created in Secure Store feature.</param>
#------------------------------------------------------------------------------------------------
function CreateSecureStoreServiceApplication
{
    param(
    [string]$applicationName,
    [string]$userNameType,
    [string]$passwordType,
    [string]$applicationType,
    [string]$userName,
    [string]$password,
    [string]$domain,
    [string]$serviceContext,
    [string]$userCredentialItem,
    [string]$passwordCredentialItem 
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if ($applicationName -eq $null -or $applicationName -eq "")
    {
        Throw "The parameter applicationName cannot be empty."
    }
    if ($userNameType -eq $null -or $userNameType -eq "")
    {
        Throw "The parameter userNameType cannot be empty."
    }
    if ($passwordType -eq $null -or $passwordType -eq "")
    {
        Throw "The parameter passwordType cannot be empty."
    }
    if ($applicationType -eq $null -or $applicationType -eq "")
    {
        Throw "The parameter applicationType cannot be empty."
    }
    if ($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if ($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if ($domain -eq $null -or $domain -eq "")
    {
        Throw "The parameter domain cannot be empty."
    }
    if ($serviceContext -eq $null -or $serviceContext -eq "")
    {
        Throw "The parameter serviceContext cannot be empty."
    }
    if ($userCredentialItem -eq $null -or $userCredentialItem -eq "")
    {
        Throw "The parameter userCredentialItem cannot be empty."
    }
    if ($passwordCredentialItem -eq $null -or $passwordCredentialItem -eq "")
    {
        Throw "The parameter passwordCredentialItem cannot be empty."
    }
    
        
    #----------------------------------------------------------------------------
    # Create the Secure Store Application.
    #----------------------------------------------------------------------------
    $poolAccount = $domain.split(".")[0] + "\" + $userName
    $secureStoreApplicationProxyName = "SecureStoreServiceApplicationProxy"
    $secureStore = Get-SPServiceApplicationProxy | where { $_.GetType().Name -eq $secureStoreApplicationProxyName }
    if(!$secureStore)
    {
        Output "Creating secure store service application proxy."
        Get-SPServiceApplication | ?{$_.GetType().Equals([Microsoft.Office.SecureStoreService.Server.SecureStoreServiceApplication])}|New-SPSecureStoreServiceApplicationProxy -Name $secureStoreApplicationProxyName -DefaultProxyGroup
        $secureStore = Get-SPServiceApplicationProxy | where { $_.GetType().Name -eq $secureStoreApplicationProxyName }
    }
    Update-SPSecureStoreMasterKey -ServiceApplicationProxy $secureStore.Id -Passphrase $password
    Start-Sleep -Seconds 60   
    
    $i = 0
    while($i++ -le 3)
    {
        try
        {
            $isTargetApplicationItemExisted = Get-SPSecureStoreApplication -ServiceContext $serviceContext -all | %{$_.TargetApplication.ApplicationID -eq $applicationName}
            Break
        }
        catch [Exception]
        {            
            sleep 60
            Continue
        }
    }
    if($isTargetApplicationItemExisted -is [array])
    {
        $isTargetApplicationItemExisted = $isTargetApplicationItemExisted.contains($true)
    }
    
    if($isTargetApplicationItemExisted)
    {   
        Output "The secure store application $applicationName already exists. Delete it first and then create a new one." "Yellow"
        Get-SPSecureStoreApplication -ServiceContext $serviceContext -Name $applicationName | Remove-SPSecureStoreApplication -Confirm:$false
    }
    Output "Creating a target application item in secure store with the name $applicationName ..." "White"
    $userNameField = New-SPSecureStoreApplicationField -name "UserName" -type $userNameType -masked:$false
    $passwordField = New-SPSecureStoreApplicationField -name "Password" -type $passwordType -masked:$true 
    $fields = $userNameField, $passwordField 
    $targetApp = New-SPSecureStoreTargetApplication -Name $applicationName -FriendlyName $applicationName -ContactEmail "$userName@$domain" -ApplicationType $applicationType 
        
    #Set the group claim and admin principals.
    $targetAppAdminAccount = New-SPClaimsPrincipal -Identity $poolAccount -IdentityType WindowsSamAccountName 

    $i = 0
    while($i++ -le 3)
    {
        try
        {
            $ssApp = New-SPSecureStoreApplication -ServiceContext $serviceContext -TargetApplication $targetApp -CredentialsOwnerGroup $targetAppAdminAccount -Administrator $targetAppAdminAccount -Fields $fields 
            Break
        }
        catch [Exception]
        {
            if($_.Exception.Message -like "*Secure Store Service is under maintenance*")
            {
                if($i -eq 2)
                {
                    Output "The secure store service is under maintenance.Please wait for a moment." "Yellow"
                }
                sleep 60
                Continue
            }
            else
            {
                Throw $_.Exception.Message
            }
        }
    }
    # Convert values to secure strings.
    $secureUserName = ConvertTo-SecureString $userCredentialItem -asplaintext -force
    $securePassword = ConvertTo-SecureString $passwordCredentialItem -asplaintext -force
  
    $credentialValues = $secureUserName, $securePassword
    # Fill in the values for the fields in the target application.
    if($applicationType -eq "Group")
    {
        Update-SPSecureStoreGroupCredentialMapping -Identity $ssApp -Values $credentialValues
    }
    else
    {
        Update-SPSecureStoreCredentialMapping -Identity $ssApp -Values $credentialValues -Principal $targetAppAdminAccount
    }
    
    Output "Create a target application item in the secure store with the name $applicationName successfully." "Green"
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Create a new login user in SQL Server. 
# </summary>
# <param name="loginUser">A valid account which could be a domain user or a local machine user, the account format is as follows: "DomainName\UserName" or "Computername\UserName".</param>
#-----------------------------------------------------------------------------------
function CreateLoginUserOnSQL
{
    param(
    [string]$loginUser
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if ($loginUser -eq $null -or $loginUser -eq $null)
    {
        Throw "The parameter loginUser cannot be empty."
    }
    
    $conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection -ArgumentList $env:ComputerName
    $conn.ServerInstance = "$env:ComputerName"
    $conn.StatementTimeout = 0
    $conn.Connect()
    $smo = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $conn
    $isExsiteloginName = $smo.Logins | ? {$_.Name -eq $loginUser}
    if($isExsiteloginName)
    {
        Output "The login user $loginUser already exists in the SQL Server." "Yellow"
    }
    else
    {
        $SqlUser = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Login -ArgumentList $smo,$loginUser
        $SqlUser.LoginType = 'WindowsUser'
        $sqlUser.PasswordPolicyEnforced = $false
        $sqlUser.PasswordExpirationEnabled = $false
        $SqlUser.Create()
        $SqlUser.AddToRole('sysadmin')
        $smo = New-Object Microsoft.SqlServer.Management.Smo.Server -ArgumentList $conn
        $isCreated = $smo.Logins | ? {$_.Name -eq $loginUser}
        if($isCreated)
        {
             Output "Create the login user $loginUser in the SQL Server successfully." "Green"
        }
    }
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Change The Authentication Classic Mode to Claim based.
# </summary>
# <param name="webApplicationUrl">The url of the web application.</param>
#-----------------------------------------------------------------------------------
function ChangeAuthenticationModeToClaimBased
{
    Param(
    [String]$webApplicationUrl 
    )
    
    #----------------------------------------------------------------------------
    # Validate parameter.
    #----------------------------------------------------------------------------
    if($webApplicationUrl -eq $null -or $webApplicationUrl -eq "")
    {
        Throw "The parameter webApplicationUrl cannot be empty."
    }
    
    #----------------------------------------------------------------------------
    # Main function.
    #----------------------------------------------------------------------------
    $uri = new-object System.Uri($webApplicationUrl)
    $webApp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($uri)    
    $useClaims = $webApp.UseClaimsAuthentication
    
    if (!$useClaims)
    {
        $webApp.UseClaimsAuthentication = $true
        $webApp.Update()
        Output "The web application '$webApplicationUrl' has been modified to use claims-based authentication." "Green"
    }
}

#-----------------------------------------------------------------------------------
# <summary>
# Check the SUT server installation mode.
# </summary>
# <param name="SharePointVersion">The SUT version.
# Note:its value is gotten by calling function ""GetSharePointVersion"</param>
# <returns>
# A boolean value, true if the server installation mode is Farm, otherwise false.
# </returns>
#-----------------------------------------------------------------------------------
function CheckServerInstallationMode
{
    param(
    [String]$SharePointVersion
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    if($SharePointVersion -eq $null -or $SharePointVersion -eq "")
    {
        Throw "The parameter sutVersion cannot be empty."
    }
    $sutShortVersion = "14.0"
    if($SharePointVersion -eq $SharePointFoundation2013[0] -or $SharePointVersion -eq $SharePointServer2013[0])
    {
        $sutShortVersion = "15.0"
    }
    elseif(($SharePointVersion -eq $SharePointServer2016[0]) -or ( $SharePointVersion -eq $SharePointServer2019[0]) -or ( $SharePointVersion -eq $SharePointServerSubscriptionEdition[0]))
    {
        $sutShortVersion = "16.0"
    }    
    $ServerModeChildItem = get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\$sutShortVersion\WSS"
    $isStandaloneInstallation = $true
    if($ServerModeChildItem.ServerRole -ieq "SINGLESERVER")
    {
        $isStandaloneInstallation = $false
    }
    return $isStandaloneInstallation  
 }
#-----------------------------------------------------------------------------------
# <summary>
# Get the language code identifier (LCID) installed on the Web server in the farm and the LCID with which the server was originally installed.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="siteUrl">The site collection url of the Web server.</param>
# <returns>An array: the first element is the default language code identifier (LCID) of the server under test (SUT),
# the second element is one installed, non-default, language code identifier (LCID) of test server under test (SUT).</returns>
#-----------------------------------------------------------------------------------
function GetSharePointLCIDs
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password,
    [String]$siteUrl
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $ret = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    param(
    [string]$siteUrl
    )

        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
        $site = new-object Microsoft.SharePoint.SPSite($siteUrl)
        return [Microsoft.SharePoint.SPRegionalSettings]::GlobalServerLanguage.LCID,$site.RootWeb.RegionalSettings.InstalledLanguages[0].LCID

    }-argumentlist $siteUrl
    
    if($ret -ne $null)
    {
        if($ret[0] -eq $null -or $ret[0] -eq "")
        {
            Output "Enter the default language code identifier (LCID) of the server under test (SUT):" "Cyan"         
            $defaultLCID = CheckForEmptyUserInput "The default language code identifier (LCID)" "defaultLCID"
            $ret[0] = $defaultLCID
            Output ("Your entered language code identifier (LCID): " + $ret[0]) "White"
        }
        if($ret[1] -eq $null -or $ret[1] -eq "")
        {
            Output "Enter one installed and non-default language code identifier (LCID) of the SUT, or leave it blank if you have only one LCID installed:" "Cyan"
            $nonDefaultLCID = GetUserInput "nonDefaultLCID"
            $ret[1] = $nonDefaultLCID
            if($ret[1] -eq "" -or $ret[1] -eq " ")
            {
                $ret[1] =  $ret[0]                
            }
            else
            {
                Output ("Your entered language code identifier (LCID): " + $ret[1]) "White"
            }
        }
    }
    else
    {
        Output "Enter the default language code identifier (LCID) of the SUT:" "Cyan"
        $defaultLCID = CheckForEmptyUserInput "The default language code identifier (LCID)" "defaultLCID"

        Output ("Your entered language code identifier (LCID): " + $defaultLCID) "White"
        
        Output "Enter one installed and non-default language code identifier (LCID) of the SUT, or leave it blank if you have only one LCID installed:" "Cyan"
        $nonDefaultLCID = GetUserInput "nonDefaultLCID"
        if($installedLCID -eq "" -or $installedLCID -eq " ")
        {
            $installedLCID = $defaultLCID
        }
        else
        {
            Output ("Your entered language code identifier (LCID): " + $installedLCID) "White"
        }
        
        $ret = $defaultLCID,$installedLCID
    }
    return $ret
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the port number of SharePoint central administration site.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <returns>[int]The script returns the port number of central administration site,
# the return value is integer type.</returns>
#-----------------------------------------------------------------------------------
function GetSharePointAdminSitePort
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $sharePointAdminPortForHTTP = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
   
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
        $adminUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.sites[0].url.Split(":")
        [int]$number = $adminUrl[2]
        return $number
    }
    if($sharePointAdminPortForHTTP -eq $null -or $sharePointAdminPortForHTTP -eq "")
    {
        Output "Unable to get the SharePointAdminPortForHTTP automatically. Enter the SharePointAdminPortForHTTP on the server:" "Cyan"
        Output "For example, the http port number used by the administration web service on the protocol server, you can enter 6396." "Cyan"
        $sharePointAdminPortForHTTP = CheckForEmptyUserInput "SharePoint Admin Site Port for HTTP" "sharePointAdminPortForHTTP"

        Output ("Your entered SharePointAdminSitePortForHTTP: " + $sharePointAdminPortForHTTP) "White"
    }      
 
    return $sharePointAdminPortForHTTP
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the https port number of SharePoint central administration site.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <returns>[int]The script returns the https port number of central administration site,
# the return value is integer type.</returns>
#-----------------------------------------------------------------------------------
function GetHttpsSharePointAdminSitePort
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $sharePointAdminPortForHTTPS = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
   
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
        $adminAlternateUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.AlternateUrls | where {$_.UrlZone -eq "internet"}
        [int]$number = $adminAlternateUrl.Uri.Port
        return $number
    }
    if($sharePointAdminPortForHTTPS -eq $null -or $sharePointAdminPortForHTTPS -eq "")
    {
        Output "Unable to get the HTTPS SharePointAdminSitePort automatically. Enter the HTTPS port number on the server:" "Cyan"
        Output "For example, for the HTTPS port number used by the administration web service on the protocol server, you can enter 9443." "Cyan"
           
        $sharePointAdminPortForHTTPS = CheckForEmptyUserInput "SharePoint Admin Site Port for HTTPS" "sharePointAdminPortForHTTPS"
        Output ("Your entered SharePointAdminSitePortForHTTPS: " + $sharePointAdminPortForHTTPS) "White"
    }      
 
    return $sharePointAdminPortForHTTPS
    
}
#-----------------------------------------------------------------------------------
# <summary>
# Get the port number of SUT web-site.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="sutWebSite">The url of sut web-site.</param>
# <returns>[int]The script returns the port number of SUT web-site,
# the return value is integer type.</returns>
#-----------------------------------------------------------------------------------
function GetSUTWebSitePort
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password,
    [String]$sutWebSite
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($sutWebSite -eq $null -or $sutWebSite -eq "")
    {
        Throw "The parameter sutWebSite cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $sutWebSitePortForHTTP = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    Param(
    [string]$sutWebSite
    )   
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
        [int]$number = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($sutWebSite).IisSettings.item(0).Path.Name
        return $number
    }-ArgumentList $sutWebSite
    
    if($sutWebSitePortForHTTP -eq $null -or $sutWebSitePortForHTTP -eq "")
    {
        Output "Unable to get the HTTP port number of the SUT website automatically. Enter the port number of the SUT website on the server:" "Cyan"
        Output "For example, for the HTTP port number used by a web service on the protocol server, you can enter 80." "Cyan"
        $sutWebSitePortForHTTP = CheckForEmptyUserInput "The HTTP Port number of SUT web-site" "sutWebSitePortForHTTP"

        Output ("Your entered SUTWebSitePortForHTTP: " + $sutWebSitePortForHTTP) "White"
    }  
 
    return $sutWebSitePortForHTTP
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the port number of SUT web-site.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="sutWebSite">The url of sut web-site.</param>
# <returns>[int]The script returns the https port number of SUT web-site,
# the return value is integer type.</returns>
#-----------------------------------------------------------------------------------
function GetHttpsSUTWebSitePort
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password,
    [String]$sutWebSite
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($sutWebSite -eq $null -or $sutWebSite -eq "")
    {
        Throw "The parameter sutWebSite cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $sutWebSitePortForHTTPS = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    Param(
    [string]$sutWebSite
    )   
        [void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
        $webSiteAlternateUrl = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($sutWebSite).AlternateUrls | where {$_.UrlZone -eq "internet"}
        [int]$number = $webSiteAlternateUrl.Uri.Port
        return $number
        
    }-ArgumentList $sutWebSite
    
    if($sutWebSitePortForHTTPS -eq $null -or $sutWebSitePortForHTTPS -eq "")
    {
        Output "Unable to get the HTTPS port number of the SUT website automatically. Enter the HTTPS port number of the SUT website on the server:" "Cyan"
        Output "For example, for the HTTPS port number used by a web service on the protocol server, you can enter 443." "Cyan"
        $sutWebSitePortForHTTPS = CheckForEmptyUserInput "The HTTPS Port number of SUT web-site" "sutWebSitePortForHTTPS"

        Output ("Your entered SUTWebSitePortForHTTPS: " + $sutWebSitePortForHTTPS) "White"
    }  
 
    return $sutWebSitePortForHTTPS
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the SharePoint Server Version. 
# </summary>
# <param name="computerName">The computer name which the SharePoint server is installed on.</param>
# <param name="userName">The user name of the SutComputerName, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <returns>
# An array with three values, represents the version information of the SharePoint Server.
# Return "WindowsSharePointServices3, Microsoft Windows SharePoint Services 3.0, SP3" if SharePoint version is "Windows SharePoint Services 3.0".
# Return "SharePointServer2007, Microsoft Office SharePoint Server 2007, SP3" if SharePoint version is "Microsoft Office SharePoint Server 2007".
# Return "SharePointFoundation2010, Microsoft SharePoint Foundation 2010, SP2" if SharePoint version is "Microsoft SharePoint Foundation 2010".
# Return "SharePointServer2010, Microsoft SharePoint Server 2010, SP2"" if SharePoint version is "Microsoft SharePoint Server 2010".
# Return "SharePointFoundation2013, Microsoft SharePoint Foundation 2013, SP1"" if SharePoint version is "Microsoft SharePoint Foundation 2013".
# Return "SharePointServer2013, Microsoft SharePoint Server 2013, SP1"" if SharePoint version is "Microsoft SharePoint Server 2013".
# Otherwise, return "Unknown Version".
# </returns>
#-----------------------------------------------------------------------------------
function GetSharePointServerVersion
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    
    #--------------------------------------------------------------------------------------------------
    # SharePoint name: the first element is short name of SharePoint, the second one is display name in registry, the third one is Service Pack name.
    #--------------------------------------------------------------------------------------------------
    $script:WindowsSharePointServices3OnSUT     = "WindowsSharePointServices3","Microsoft Windows SharePoint Services 3.0","SP3"
    $script:SharePointServer2007OnSUT           = "SharePointServer2007","Microsoft Office SharePoint Server 2007 ","SP3"
    $script:SharePointFoundation2010OnSUT       = "SharePointFoundation2010","Microsoft SharePoint Foundation 2010","SP2"
    $script:SharePointServer2010OnSUT           = "SharePointServer2010","Microsoft SharePoint Server 2010","SP2"
    $script:SharePointFoundation2013OnSUT       = "SharePointFoundation2013","Microsoft SharePoint Foundation 2013","SP1"
    $script:SharePointServer2013OnSUT           = "SharePointServer2013","Microsoft SharePoint Server 2013","SP1"
    $script:SharePointServer2016OnSUT           = "SharePointServer2016","Microsoft SharePoint Server 2016"
    $script:SharePointServer2019OnSUT           = "SharePointServer2019","Microsoft SharePoint Server 2019"
    $script:SharePointServerSubscriptionEditionOnSUT           = "SharePointServerSubscriptionEdition","Microsoft SharePoint Server Subscription Edition"
    $SharePointVersion                          = "Unknown Version"
    
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

   $sutVersion = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    param(
    $script:WindowsSharePointServices3OnSUT,$script:SharePointServer2007OnSUT,$script:SharePointFoundation2010OnSUT,$script:SharePointServer2010OnSUT,$script:SharePointFoundation2013OnSUT,$script:SharePointServer2013OnSUT,$script:SharePointServer2016OnSUT,$script:SharePointServer2019OnSUT,$script:SharePointServerSubscriptionEditionOnSUT
    )

        $keys = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $items = $keys | foreach-object {Get-ItemProperty $_.PsPath}    
        foreach ($item in $items)
        {
            if($item.DisplayName -eq $script:WindowsSharePointServices3OnSUT[1])
            {
                $SharePointVersion = $script:WindowsSharePointServices3OnSUT
                foreach ($item in $items)
                {
                    if($item.DisplayName -eq ($script:SharePointServer2007OnSUT[1]))
                    {
                        $SharePointVersion = $script:SharePointServer2007OnSUT[0], $script:SharePointServer2007OnSUT[1].TrimEnd(),$script:SharePointServer2007OnSUT[2]
                        break
                    }
                }
                break
            }
            elseif($item.DisplayName -eq $script:SharePointFoundation2010OnSUT[1])
            {
                $SharePointVersion = $script:SharePointFoundation2010OnSUT
                break
            }        
            elseif($item.DisplayName -eq $script:SharePointServer2010OnSUT[1])
            {
                $SharePointVersion = $script:SharePointServer2010OnSUT
                break
            }        
            elseif($item.DisplayName -eq $script:SharePointFoundation2013OnSUT[1])
            {
                $SharePointVersion = $script:SharePointFoundation2013OnSUT
                break
            }        
            elseif($item.DisplayName -eq $script:SharePointServer2013OnSUT[1])
            {
                $SharePointVersion = $script:SharePointServer2013OnSUT[0], $script:SharePointServer2013OnSUT[1].TrimEnd(),$script:SharePointServer2013OnSUT[2]
                break
            }
            elseif($item.DisplayName -eq $script:SharePointServer2016OnSUT[1])
            {
                $SharePointVersion = $script:SharePointServer2016OnSUT[0], $script:SharePointServer2016OnSUT[1]
                break
            }
            elseif($item.DisplayName.StartsWith($script:SharePointServer2019OnSUT[1]))
            {
                $SharePointVersion = $script:SharePointServer2019OnSUT[0], $script:SharePointServer2019OnSUT[1]
                break
            }
            elseif($item.DisplayName.StartsWith($script:SharePointServerSubscriptionEditionOnSUT[1]))
            {
                $SharePointVersion = $script:SharePointServerSubscriptionEditionOnSUT[0], $script:SharePointServerSubscriptionEditionOnSUT[1]
                break
            }
        }
        return $SharePointVersion
    }-ArgumentList $script:WindowsSharePointServices3OnSUT,$script:SharePointServer2007OnSUT,$script:SharePointFoundation2010OnSUT,$script:SharePointServer2010OnSUT,$script:SharePointFoundation2013OnSUT,$script:SharePointServer2013OnSUT,$script:SharePointServer2016OnSUT,$script:SharePointServer2019OnSUT,$script:SharePointServerSubscriptionEditionOnSUT
    
    return $sutVersion
}
#-----------------------------------------------------------------------------------
# <summary>
# Get the SharePoint Server Version manually. 
# </summary>
# <returns>
# An array with three values, represents the version information of the SharePoint Server.
# Return "SharePointFoundation2010, Microsoft SharePoint Foundation 2010, SP2" if SharePoint version is "Microsoft SharePoint Foundation 2010".
# Return "SharePointServer2010, Microsoft SharePoint Server 2010, SP2" if SharePoint version is "Microsoft SharePoint Server 2010".
# Return "SharePointFoundation2013, Microsoft SharePoint Foundation 2013, SP1" if SharePoint version is "Microsoft SharePoint Foundation 2013".
# Return "SharePointServer2013, Microsoft SharePoint Server 2013, SP1" if SharePoint version is "Microsoft SharePoint Server 2013".
# Return "SharePointServer2013, Microsoft SharePoint Server 2016" if SharePoint version is "Microsoft SharePoint Server 2016".
# Otherwise, return "Unknown Version".
# </returns>
#-----------------------------------------------------------------------------------
function GetSharePointVersionManually
{
    
    Output "Unable to get the SharePoint version automatically. Select the SharePoint version: " "Cyan"
    Output "1: Microsoft SharePoint Foundation 2010 SP2" "Cyan"
    Output "2: Microsoft SharePoint Server 2010 SP2" "Cyan"
    Output "3: Microsoft SharePoint Foundation 2013 SP1" "Cyan"
    Output "4: Microsoft SharePoint Server 2013 SP1" "Cyan"
    Output "5: Microsoft SharePoint Server 2016" "Cyan"
    Output "6: Microsoft SharePoint Server 2019" "Cyan"
    Output "7: Microsoft SharePoint Server Subscription Edition" "Cyan"
    $isManualSelectVersion = $true
    $sutVersionChoices = @('1: Microsoft SharePoint Foundation 2010 SP2','2: Microsoft SharePoint Server 2010 SP2','3: Microsoft SharePoint Foundation 2013 SP1','4: Microsoft SharePoint Server 2013 SP1','5: Microsoft SharePoint Server 2016','6: Microsoft SharePoint Server 2019','7: Microsoft SharePoint Server Subscription Edition')
    $sutVersion = ReadUserChoice $sutVersionChoices "sutVersion"
    Switch($sutVersion)
    {
        "1" {$sutVersion = $script:SharePointFoundation2010OnSUT[0]; break }
        "2" {$sutVersion = $script:SharePointServer2010OnSUT[0]; break }
        "3" {$sutVersion = $script:SharePointFoundation2013OnSUT[0]; break }
        "4" {$sutVersion = $script:SharePointServer2013OnSUT[0]; break }
        "5" {$sutVersion = $script:SharePointServer2016OnSUT[0]; break } 
        "6" {$sutVersion = $script:SharePointServer2019OnSUT[0]; break }
        "7" {$sutVersion = $script:SharePointServerSubscriptionEditionOnSUT[0]; break }       
    }
    
    return $sutVersion

}
#-----------------------------------------------------------------------------------
# <summary>
# Get the GUID that uniquely identifies the file in the content database.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="siteCollectionUrl">Specify the url of a site collection.</param>
# <param name="folderUrl">A string that specifies the URL of the folder.</param>
# <param name="fileName">A string that specifies the name of file.</param>
# <returns>[string]The script returns the GUID that uniquely identifies the file in the content database. </returns>
#-----------------------------------------------------------------------------------
function GetFileId
{
    param(
    [string]$computerName,
    [string]$userName,
    [string]$password,
    [string]$siteCollectionUrl,
    [string]$folderUrl,
    [string]$fileName

    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($siteCollectionUrl -eq $null -or $siteCollectionUrl -eq "")
    {
        Throw "The parameter siteCollectionUrl cannot be empty."
    }
    if($folderUrl -eq $null -or $folderUrl -eq "")
    {
        Throw "The parameter folderUrl cannot be empty."
    }
    if($fileName -eq $null -or $fileName -eq "")
    {
        Throw "The parameter fileName cannot be empty."
    }
    
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $uniqueId = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{        
        Param(
        [string]$siteCollectionUrl,
        [string]$folderUrl,
        [string]$fileName
        )
   
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

        $spSites = new-object Microsoft.SharePoint.SPSite($siteCollectionUrl)
        $spWeb =  $spSites.OpenWeb()
        $file = $spWeb.GetFile("$siteCollectionUrl\$folderUrl\$fileName")
        $uniqueId = $file.UniqueId.ToString("B")
        return $uniqueId
    }-ArgumentList $siteCollectionUrl,$folderUrl,$fileName
    
    return $uniqueId
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the connection string from Active Directory Domain Services.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must keep the format of DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="domain">The user domain name.</param>
# <returns>[string]The script returns the connection string from Active Directory Domain Services.</returns>
#-----------------------------------------------------------------------------------
function GetADConnectionString
{
    param(
    [string]$computerName,
    [string]$userName,
    [string]$password,
    [string]$domain
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($domain -eq $null -or $domain -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $userFQDN = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{   
        $userFQDN = whoami /fqdn
        return $userFQDN
    }
    if($userFQDN -eq "" -or $userFQDN -eq $null)
    {
        Output "Unable to get the connection string from Active Directory Domain Services. Enter the connection string:" "Cyan"
        $userFQDN = CheckForEmptyUserInput "The connection string from Active Directory Domain Services" "userFQDN"

        Output ("Your entered the connection string: " + $userFQDN) "White"
    }
    
    $groupN = "DC="
    if($userFQDN -like "*OU=*")
    {
        $groupN = "OU="
    }
    $userFQDN_substring = $userFQDN.Substring($userFQDN.IndexOf("$groupN"))
    $adConnectingString = "LDAP://" + $domain + "/" + $userFQDN_substring
    return $adConnectingString
    
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the hold id and url.
# </summary>
# <param name="computerName">The computer name of the server.</param>
# <param name="userName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="password">The password of the user name.</param>
# <param name="siteUrl">The site url of the Web server.</param>
# <param name="holdName">The hold name under siteUrl.</param>
# <returns>An array: the first element is the hold id,the second element is the hold url.</returns>
#-----------------------------------------------------------------------------------
function GetHoldInfo
{
    param(
    [String]$computerName,
    [String]$userName,
    [String]$password,
    [String]$siteUrl,
    [String]$holdName
    )
    
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($computerName -eq $null -or $computerName -eq "")
    {
        Throw "The parameter computerName cannot be empty."
    }
    if($userName -eq $null -or $userName -eq "")
    {
        Throw "The parameter userName cannot be empty."
    }
    if($password -eq $null -or $password -eq "")
    {
        Throw "The parameter password cannot be empty."
    }
    if($siteUrl -eq $null -or $siteUrl -eq "")
    {
        Throw "The parameter siteUrl cannot be empty."
    }
    if($holdName -eq $null -or $holdName -eq "")
    {
        Throw "The parameter holdName cannot be empty."
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($userName,$securePassword)

    $ret = invoke-command -computer $computerName -Credential $credential -ErrorAction SilentlyContinue -scriptblock{
    param(
    [string]$siteUrl,
    [string]$holdName
    )

        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
        $spsite = New-Object Microsoft.SharePoint.SPSite($siteUrl)
        $spWeb = $spsite.OpenWeb()
        $spList = $spWeb.Lists["Holds"]
        $counts = $spList.ItemCount
        $holdInfo = $null
        for($i = 0;$i -lt $counts;$i++)
        {           
            if($spList.items[$i].Name -eq $holdName)
            {
                $id = $spList.items[$i].id
                $url = $spList.items[$i].url
                $holdInfo = $id,$url
                break
                
            }        
        }
        return $holdInfo

    }-argumentlist $siteUrl,$holdName
    
    if($ret -ne $null)
    {
        if($ret[0] -eq $null -or $ret[0] -eq "")
        {
            Output "Enter the hold $holdName ID:" "Cyan"         
            $holdId = CheckForEmptyUserInput "The hold $holdName ID" "holdId"
            $ret[0] = $holdId
            Output ("Your entered hold $holdName ID: " + $ret[0]) "White"
        }
        if($ret[1] -eq $null -or $ret[1] -eq "")
        {
            Output "Enter the hold $holdName url:" "Cyan"
            $holdUrl = CheckForEmptyUserInput "The hold $holdName Url" "holdUrl"
            $ret[1] = $holdUrl            
            Output ("Your entered hold $holdName Url: " + $ret[1]) "White"
        }
    }
    else
    {
         Output "Enter the hold $holdName ID:" "Cyan"         
         $holdId = CheckForEmptyUserInput "The hold $holdName ID" "holdId"
         Output ("Your entered hold $holdName ID: " + $holdId) "White"
        
         Output "Enter the hold $holdName url:" "Cyan"
         $holdUrl = CheckForEmptyUserInput "The hold $holdName Url" "holdUrl"
         Output ("Your entered hold $holdName Url: " + $holdUrl) "White"
        
        $ret = $holdId,$holdUrl
    }
    return $ret
}

#-----------------------------------------------------------------------------------
# <summary>
# Start a url in IE.
# </summary>
# <param name="trustedSite">The site will be added to the "Trusted Sites" zone.</param>
# <param name="url">The url will be opened in IE.</param>
#-----------------------------------------------------------------------------------
function StartUrlInIE
{
    Param(
    [string] $trustedSite,
    [string] $url
    )
    #----------------------------------------------------------------------------
    # Parameter validation.
    #----------------------------------------------------------------------------
    if($trustedSite -eq $null -or $trustedSite -eq "")
    {
        Throw "The parameter trustedSite cannot be empty."
    }
    if($url -eq $null -or $url -eq "")
    {
        Throw "The parameter url cannot be empty."
    }
       
    $regeditPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\$trustedSite"
    if (!(Test-Path -Path $regeditPath))
    {
        New-Item -Path $regeditPath | Out-File -FilePath $logFile -Append -Encoding ASCII -Width 100
    }
    Set-ItemProperty -Path $regeditPath -Name http -Value 1 -Type DWord
    Set-ItemProperty -Path $regeditPath -Name https -Value 1 -Type DWord
    $ie=New-Object -com InternetExplorer.Application
    $ie.Visible=$true
    $ie.Navigate($url)
}

#-----------------------------------------------------------------------------------
# <summary>
# Get the DisplayName of a SharePoint user.
# </summary>
# <param name="sutComputerName">The computer name of the server.</param>
# <param name="siteCollectionName">The name of the sitecollection.</param>
# <param name="userName">The name of a SharePoint user.</param>
# <param name="credentialUserName">The user name of the server, must be in the format DOMAIN\User_Alias.</param>
# <param name="credentialPassword">The password of the domain user name.</param>
# <returns>the DisplayName of a SharePoint user.</returns> 
#-----------------------------------------------------------------------------------
function GetUserDisplayName
{
    param(
    [string]$sutComputerName,
    [string]$siteCollectionName,
    [string]$userName,
    [string]$credentialUserName,
    [string]$credentialPassword 
    )

    #----------------------------------------------------------------------------
    # Parameter validation
    #----------------------------------------------------------------------------
    ValidateParameter 'sutComputerName' $sutComputerName	
    ValidateParameter 'siteCollectionName' $siteCollectionName	
    ValidateParameter 'userName' $userName	
    ValidateParameter 'credentialUserName' $credentialUserName	
    ValidateParameter 'credentialPassword' $credentialPassword	

    $securePassword = ConvertTo-SecureString $credentialPassword -AsPlainText -Force
    $credential = new-object Management.Automation.PSCredential($credentialUserName,$securePassword)
    $userDisplayName = Invoke-Command -ComputerName $sutComputerName -Credential $credential -ErrorAction SilentlyContinue -ScriptBlock {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
        $webApplicationUrl = "http://$($args[0])"
        $webSite = "http://$($args[0])/sites/$($args[1])"
        $webApp = Get-SPWebApplication $webApplicationUrl
        $userName = "$env:USERDOMAIN\$($args[2])"
        if($webApp.UseClaimsAuthentication)
        {
            $userName = "i:0#.w|" + $userName
        } 
        $user = Get-SPUser -Web $webSite -Identity  $userName -ErrorAction SilentlyContinue
        if($user -ne $null)
        {
            return $user.DisplayName
        }
    } -ArgumentList $sutComputerName,$siteCollectionName,$userName

    if($userDisplayName -eq "" -or $userDisplayName -eq $null)
    {
        #-----------------------------------------------------
        # Get $userName's DisplayName manually
        #-----------------------------------------------------
        Output "Unable to get $userName's DisplayName automatically." "Yellow"
        Output "Enter the display name of the user $userName on $siteCollectionName" "Cyan"    
        $userDisplayName = CheckForEmptyUserInput "the DisplayName of user $userName" $userName
        Output "Your entered $userDisplayName" "White"
    }
    return $userDisplayName
}