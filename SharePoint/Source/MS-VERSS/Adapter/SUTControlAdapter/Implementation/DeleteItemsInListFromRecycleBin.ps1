$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName

$requestUrl = .\Get-ConfigurationPropertyValue.ps1 RequestUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$requestUrl,
       [string]$listName
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  $ret=$false
  try
  {
      $spSites = new-object Microsoft.SharePoint.SPSite $requestUrl
      $itemColl=$spSites.RecycleBin
      for($i=0;$i -le $itemColl.Count-1;$i++)
      {
          if($itemColl.Item($i).DirName.Contains($listName))
          {
              $idsarray = [System.Guid]$itemColl.Item($i).ID
              $itemColl.Delete($idsarray)
          }
          if($itemColl.Item($i).Title -eq $listName)
          {
             $idsarray = [System.Guid]$itemColl.Item($i).ID
             $itemColl.Delete($idsarray)
          }
      }
      $ret=$true
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
} -argumentlist $requestUrl, $listName

return $ret