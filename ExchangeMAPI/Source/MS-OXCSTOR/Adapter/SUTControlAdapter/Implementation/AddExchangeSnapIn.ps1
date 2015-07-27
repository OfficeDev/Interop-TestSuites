#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

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