$script:ErrorActionPreference = "Stop"
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$sutVersion = .\Get-ConfigurationPropertyValue.ps1 SutVersion

if($sutVersion -ieq "SharePointServer2010" -or $sutVersion -ieq "SharePointServer2013")
{
	$ret = invoke-command -computer $computerName -Credential $credential -scriptblock{
		param(
		 [Bool]$setDisabled,
		 [string]$computerName
		)
		$script:ErrorActionPreference = "Stop"

	    [void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
	    $userProfileService = [Microsoft.SharePoint.Administration.SPFarm]::Local.Services | where {$_.TypeName -eq "User Profile Service"}
	    [Microsoft.SharePoint.Administration.SPServiceInstanceDependencyCollection] $serviceInstance
	    if($userProfileService -eq "" -or $userProfileService -eq $null -or $userProfileService.Instances -eq "" -or $userProfileService.Instances -eq $null)
	    {
	        write-host "The User Profile Service is not installed on SUT"
			return $true
	    }
		else
		{	
	        try
			{
				foreach($serviceInstance in $userProfileService.Instances)
				{
					if($serviceInstance.Server.Address -eq $computerName)
					{
						if($setDisabled)
						{   
							$serviceInstance.Unprovision()
						}
						else
						{
							$serviceInstance.Provision()
						}	
						$serviceInstance.Update()
						return $true	
					}
				}
				return $false
			}
			catch
			{
			    return $false
			}
		}
	}-ArgumentList $setDisabled, $computerName

	return $ret
}
else
{
    return $true
}