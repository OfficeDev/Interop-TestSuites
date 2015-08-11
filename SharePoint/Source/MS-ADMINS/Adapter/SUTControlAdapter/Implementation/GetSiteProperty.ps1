$script:ErrorActionPreference = "Stop"

#PTF Properties
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

#Convert password from plain test to secure strings
$securePassword = convertTo-SecureString $password -AsPlainText -Force

#Initializes a new instance of the Sytstem.Net.NetworkCredential class with the specified user name and password read from ptfconfig file
$credential=new-object Management.Automation.PSCredential(($domain+"\"+$userName), $securePassword)

#Execute the following script remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
      #External Parameters:
      param( [string] $url , [string]$proName)

      $script:ErrorActionPreference = "Stop"

      #Load Assemblies Microsoft.SharePoint.dll PSPowerShell.dll
      [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
      #Get SharePoint List
      $spSite = new-object Microsoft.SharePoint.SPSite($url)
     
	  #Get Site Property 

      $retValue = ""
      try
	  {
			if (($proName -eq "Title") -or ($proName -eq "Description") -or ($proName -eq "WebTemplate"))
			{
				$retValue = $spSite.rootweb.($proName)
            }
			elseif($proName -eq "OwnerName")
			{
				$retValue = $spSite.Owner.Name
		    }
			elseif($proName -eq "OwnerEmail")
			{
				$retValue = $spSite.Owner.Email
			}
			else
			{
				$retValue = $spSite.($proName)
			}
		}
		catch
		{
			throw $error[0]
		}
        finally
		{
			$spSite.Dispose()
			$spSite.Close()
        }

          return $retValue
        
} -argumentlist $url, $proName 
  
$ret