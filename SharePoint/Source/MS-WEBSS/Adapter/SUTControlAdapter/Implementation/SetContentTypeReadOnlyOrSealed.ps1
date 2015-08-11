$script:ErrorActionPreference = "Stop"

$password= .\Get-ConfigurationPropertyValue.ps1 Password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$domain=.\Get-ConfigurationPropertyValue.ps1 Domain
$userName=.\Get-ConfigurationPropertyValue.ps1 UserName
$credential = new-object Management.Automation.PSCredential(($domain+"\"+$userName),$securePassword) 
$computerName = .\Get-ConfigurationPropertyValue.ps1 SutComputerName
$mainUrl = .\Get-ConfigurationPropertyValue.ps1 SubSiteUrl

$ret = Invoke-Command -ComputerName $computerName -Credential $credential -ScriptBlock{
  param(
       [string]$webSiteName,
       [string]$contentTypeName,
       [bool]$isReadOnly,
       [bool]$isSealed,
       [string]$mainUrl
  )
  $script:ErrorActionPreference = "Stop"
  [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | out-null
  try
  {
    $spSite = new-object Microsoft.SharePoint.SPSite($mainUrl)
    $spWebs =  [Microsoft.SharePoint.SPWeb]$spSite.OpenWeb("")
    $contentTypes = [Microsoft.SharePoint.SPContentTypeCollection]$spWebs.ContentTypes;
    [int]$Index=$contentTypes.Count
    for( $intIndex = 0; $intIndex -lt $Index; $intIndex++)
    {
        $name = $contentTypes[$intIndex].Name;
        if($name -eq $contentTypeName)
        {
            if($isReadOnly)
            {
                $contentTypes[$intIndex].ReadOnly = [bool]$isReadOnly
                $contentTypes[$intIndex].Update()
            }
            if($isSealed)
            {
                $contentTypes[$intIndex].Sealed = [bool]$isSealed
                $contentTypes[$intIndex].Update()
            }
            return;
        }
    }
  }
  catch
  {
      throw $error[0]
  }
  finally
  {
      if($spSite -ne $null)
      {
          $spSite.Dispose()
      }
  }
} -argumentlist $webSiteName,$contentTypeName,$isReadOnly,$isSealed,$mainUrl