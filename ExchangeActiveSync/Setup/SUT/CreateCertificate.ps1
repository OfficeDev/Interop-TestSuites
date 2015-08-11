#----------------------------------------------------------------------------
# <param name="mailboxUserName">The name of the mailbox user.</param>
# <param name="userPassword">The password of the mailbox user.</param>
# <param name="pfxFileName">The name of the personal encryption certificate.</param>
#----------------------------------------------------------------------------
param(
[string]$mailboxUserName,
[string]$userPassword,
[string]$pfxFileName
)

$certFolderPath = & {Split-Path $MyInvocation.scriptName}
$policyFile="$certFolderPath\cert.inf"
$requestFile="$certFolderPath\requestFile.req"
$certFile="$certFolderPath\certFile.cer"
$pfxFile ="$certFolderPath\$pfxFileName"

#Create a personal certificate for $mailboxUserName
certreq -new -f -q $policyFile $requestFile
certreq -submit -f -q $requestFile $certFile
certreq -accept $certFile
Import-Module ActiveDirectory
$userInfo = Get-ADUser $mailboxUserName -Properties "Certificates"
$userCertificates = $userInfo.Certificates | foreach {New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $_}
if(($userCertificates -eq $null) -and ($userCertificates -eq ""))
{
    throw "Failed to create the personal certificate for mailbox user $mailboxUserName."
}
else
{
    #Export the personal encryption certificate
    if($userCertificates -is [array])
    {
        certutil -user -f -p $userPassword -exportPFX my $userCertificates[0].Thumbprint $pfxFile 
    }
    else
    {
        certutil -user -f -p $userPassword -exportPFX my $userCertificates.Thumbprint $pfxFile	
    }   
}

#----------------------------------------------------------------------------
# Ending script
#----------------------------------------------------------------------------
exit 0