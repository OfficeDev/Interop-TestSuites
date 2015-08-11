$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName

$requestUrl=.\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$workFlowAssociationName,
       [string]$requestUrl,
       [string]$targetListName
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
      $spWeb =  $spSites.RootWeb
      $targetList = $spWeb.Lists[$targetListName]
      $associationItems = $targetList.WorkflowAssociations 

      if($associationItems -eq $null)
      {
         $ret = $null
         return  $ret
      }
      
      $cultureInfoInstance = new-object System.Globalization.CultureInfo("en-US")
      $specifiedAssociationItem = $associationItems.GetAssociationByName($workFlowAssociationName, $cultureInfoInstance)
      
      if($specifiedAssociationItem -eq $null)
      {
         $ret = $null
         return  $ret
      }

    $ret = $specifiedAssociationItem.Id.Tostring()
  }
  catch
  {
    throw $error[0]
  }
  finally
  {
      if($spSites -ne $null)
      {
          $spSites.Dispose()
      }
  }
  return $ret
} -argumentlist $workFlowAssociationName, $requestUrl, $targetListName

return $ret