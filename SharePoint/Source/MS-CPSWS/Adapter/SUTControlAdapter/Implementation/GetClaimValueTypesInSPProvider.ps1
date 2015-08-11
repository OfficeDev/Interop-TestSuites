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
$claimValueTypeArray = new-object system.collections.arraylist($null)
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
$basicClaimValueTypes = [Microsoft.SharePoint.Administration.Claims.SPClaimProviderManager]::BasicClaimValueTypes
foreach($basicClaimValueType in $basicClaimValueTypes)
{
    $claimValueTypeArray.Add($basicClaimValueType)
}

Add-PSSnapin Microsoft.SharePoint.Powershell
$claimProviders = Get-SPClaimProvider

foreach ($inputClaimProviderName in $inputClaimProviderNames.split(","))
{
    foreach($claimProvider in $claimProviders)
    {
        if($inputClaimProviderName -eq $claimProvider.ClaimProvider.Name)
        {
          $claimValueTypes = $claimProvider.ClaimProvider.ClaimValueTypes()
          break
        }
    }

        foreach($claimValueType in $claimValueTypes)
        {
           if(!$claimValueTypeArray.contains($claimValueType))
           {
           $claimValueTypeArray.add($claimValueType)
           }
        }
}

$resultArray= ""
foreach($claimValueType in $claimValueTypeArray)
{
	if($claimValueType -ne $null)
	{
		$resultArray+=$claimValueType.ToString()+","
	}
}

$result = $resultArray.trim(",")
return $result
}-argumentlist $inputClaimProviderNames
return $ret