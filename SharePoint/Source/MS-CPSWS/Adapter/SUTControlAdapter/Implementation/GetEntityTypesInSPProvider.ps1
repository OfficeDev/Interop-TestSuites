#-----------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-----------------------------------------------------------------------

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
[string]$inputClaimProviderNames
)

$script:ErrorActionPreference = "Stop"
  try
  {
    Add-PSSnapin Microsoft.SharePoint.Powershell
    $claimProviders = Get-SPClaimProvider
    $entityTypeArray = new-object system.collections.arraylist($null)
    foreach ($inputClaimProviderName in $inputClaimProviderNames.split(","))
    {
        foreach($claimProvider in $claimProviders)
        {
            if($inputClaimProviderName -eq $claimProvider.ClaimProvider.Name)
            {
                $entityTypes = $claimProvider.ClaimProvider.EntityTypes()

                if($entityTypes -ne $null)
                {
                    foreach($entityType in $entityTypes)
                    {
                        if(!$entityTypeArray.contains($entityType))
                        {
                            $entityTypeArray.add($entityType)
                        }
                    }
                }
            }
        }
    }

    $resultArray = ""
    foreach($entityType in $entityTypeArray)
    {
        $resultArray+=$entityType.ToString()+","
    }

    $result = $resultArray.trim(",")
    return $result
  }
  catch
  {
    throw $error[0]
  }
}-argumentlist $inputClaimProviderNames
return $ret