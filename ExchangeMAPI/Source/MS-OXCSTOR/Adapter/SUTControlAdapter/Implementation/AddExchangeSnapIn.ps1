param(
    $serverVersion
)

$ExchangeShellSnapIn2010 = "Microsoft.Exchange.Management.PowerShell.E2010"
$ExchangeShellSnapIn2007 = "Microsoft.Exchange.Management.PowerShell.Admin"
$ExchangeShellSnapIn = $ExchangeShellSnapIn2010

$script:ErrorActionPreference = "Stop"

#----------------------------------------------------------------------------
# If the Exchange server is 2007, change ExchangeShellSnapIn to 2007
#----------------------------------------------------------------------------
If($serverVersion -ne $null -and $serverVersion -ne "" -and $serverVersion -eq "ExchangeServer2007")
{
    $ExchangeShellSnapIn = $ExchangeShellSnapIn2007 
}

#----------------------------------------------------------------------------
# Add the snap in
#----------------------------------------------------------------------------
Add-PSSnapin $ExchangeShellSnapIn