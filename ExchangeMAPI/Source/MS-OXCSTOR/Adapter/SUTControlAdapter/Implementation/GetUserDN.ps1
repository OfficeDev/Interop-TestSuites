#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

$script:ErrorActionPreference = "Stop"

$credentialUserName = "$PtfPropDomain\$PtfPropUserName"
$credentialPassword = $PtfPropUserPassword

#----------------------------------------------------------------------------
# Parameter validation
#----------------------------------------------------------------------------
if ($computerName -eq $null -or $computerName -eq "")
{
    Throw "Parameter computerName is required."
}
if ($userName -eq $null -or $userName -eq "")
{
    Throw "Parameter userName is required."
}
if ($credentialUserName -eq $null -or $credentialUserName -eq "")
{
    Throw "Parameter credentialUserName is required."
}
if ($credentialPassword -eq $null -or $credentialPassword -eq "")
{
    Throw "Parameter credentialPassword is required."
}

$securePassword = ConvertTo-SecureString $credentialPassword -AsPlainText -Force
$credential = new-object Management.Automation.PSCredential($credentialUserName,$securePassword)

Invoke-Command -ComputerName $computerName -Credential $credential -ErrorAction SilentlyContinue -ScriptBlock {

	# Create A New ADSI Call
    $dnsDomain = $args[1].Split("\")[0]
    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$dnsDomain",$args[1],$args[2])
    # Create a New DirectorySearcher Object
    $searcher = new-object System.DirectoryServices.DirectorySearcher($root)
    # Set the filter to search for a specific CNAME
    $temp = $args[0]
    $searcher.filter = "(&(objectClass=user) (CN=$temp))"
    # Set results in $adFind variable
    $adFind = $searcher.findall()

	 return $adFind[0].Properties.legacyexchangedn
} -Args $userName,$credentialUserName,$credentialPassword