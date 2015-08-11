$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$computerName = .\Get-ConfigurationPropertyValue.ps1 SUTComputerName

$requestUrl= .\Get-ConfigurationPropertyValue.ps1 TargetServiceUrl

$securePassword = $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword)

#invoke function remotely
$ret = invoke-command -computer $computerName -Credential $credential -scriptblock {
  param(
       [string]$currentTaskListName,
       [string]$requestUrl,
       [string]$taskIds
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  try
  {   
      if($taskIds -eq $null)
      { 
         $ret = $true
         return $ret
      }
      
      $spSites = new-object Microsoft.SharePoint.SPSite "$requestUrl"
      $spWeb =  $spSites.RootWeb
      $specifiedTaskList = $spWeb.Lists[$currentTaskListName]
      if($specifiedTaskList -eq $null)
      {
         $ret = $false
         return  $ret
      }

        $taskIdArray = $taskIds.Split(',')
        $ErrorCounter = 0
        $ret = $false
        foreach($taskIditem in $taskIdArray)
        {
            try
            {
                $taskIdValue = [int]::Parse($taskIditem)
                $tasklistiem = $specifiedTaskList.GetItemById($taskIdValue)
                if($tasklistiem -ne $null )
                {
                    $tasklistiem.Delete()
                    $specifiedTaskList.Update()
                }
                else
                {
                    $ErrorCounter ++
                }
            }
            catch
            {
                $ErrorCounter ++
            }
        }

    $ret = $ErrorCounter -eq 0
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

} -argumentlist $currentTaskListName, $requestUrl, $taskIds

return $ret