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
       [bool]$isEnabled
  )
  $script:ErrorActionPreference = "Stop"
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
  $ret=$false
  try
  {
      $webapp = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($requestUrl)
      $webapp.RecycleBinEnabled = $isEnabled
      $webapp.Update()
      $ret=$true
  }
  catch
  {
      throw $error[0]
  }
  return $ret
} -argumentlist $requestUrl, $isEnabled

return $ret