$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$SutComputerName=.\Get-ConfigurationPropertyValue.ps1 SutComputerName
$securepassword=ConvertTo-SecureString $password -AsPlainText -Force
$credential=new-object Management.Automation.PSCredential(($domain+"\"+$username),$securepassword)

$ret=invoke-command -computer $SutComputerName  -credential $credential -ScriptBlock{
#----------------------------------------------------------------------------------------------------
param(
[string] $inputClaimProviderNames
)

$script:ErrorActionPreference = "Stop"
$claimTypeArray = new-object system.collections.arraylist($null)
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
$basicClaimTypes = [Microsoft.SharePoint.Administration.Claims.SPClaimProviderManager]::BasicClaimTypes

foreach($basicClaimType in $basicClaimTypes)
{
    $claimTypeArray.Add($basicClaimType)
}

Add-PSSnapin Microsoft.SharePoint.Powershell
$claimProviders = Get-SPClaimProvider

foreach ($inputClaimProviderName in $inputClaimProviderNames.split(","))
{

    foreach($claimProvider in $claimProviders)
    {
        if($inputClaimProviderName -eq $claimProvider.ClaimProvider.Name)
        {
          $claimTypes = $claimProvider.ClaimProvider.ClaimTypes()
          break
        }
    }

    foreach($claimType in $claimTypes)
    {
        if(!$claimTypeArray.contains($claimType))
        {
           $claimTypeArray.add($claimType)
        }
    }
}

$resultArray = ""
foreach($claimType in $claimTypeArray)
{
	if($claimType -ne $null)
	{
		$resultArray+=$claimType.ToString()+","
	}
}

$result = $resultArray.trim(",")
return $result
}-argumentlist $inputClaimProviderNames
return $ret