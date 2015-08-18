$script:ErrorActionPreference = "Stop"
[string]$computerName = $PtfPropSutComputerName
[string]$usr = $PtfPropAdminUserName
[string]$pwd = $PtfPropAdminUserPassWord
[string]$domainName = $PtfPropDomain

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true} 

$credentials= New-Object System.Net.NetworkCredential $usr,$pwd,$domainName

# Remember to set exchange server address
$exchangeServerAddress = $PtfPropEwsUrl.Replace("[SutComputerName]", $computerName).Replace("[Domain]", $domainName)

$soapRequest = @'
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <CreateItem MessageDisposition="SendAndSaveCopy" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <Items>
        <t:Message>
          <t:ItemClass>IPM.Note</t:ItemClass>
          <t:Subject>This is an interval event test mail, let's go!</t:Subject>
          <t:Body BodyType="Text">The body part is not important, these words are totally useless!</t:Body>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>{0}</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:IsRead>false</t:IsRead>
        </t:Message>
      </Items>
    </CreateItem>
  </soap:Body>
</soap:Envelope>

'@

# Create the request 
$webRequest = [System.Net.WebRequest]::Create($exchangeServerAddress)
$webRequest.ContentType = "text/xml"
$webRequest.Headers.Add("Translate", "F")
$webRequest.Method = "Post"
$webRequest.Credentials = $credentials
# Timeout property indicates the length of time, in milliseconds, until the request time out and throws a webException
# Set it to 200000 aims to avoid the request time out
$webRequest.Timeout = 200000

# Setup the soap request to send to the server
$soapRequest = $soapRequest -f ($usr + "@" + $domainName)
$content = [System.Text.Encoding]::UTF8.GetBytes($soapRequest)
$webRequest.ContentLength = $content.Length
$requestStream = $webRequest.GetRequestStream()
$requestStream.Write($content, 0, $content.Length)
$requestStream.Close()

# Get the xml response from the server
$webResponse = $webRequest.GetResponse()
$responseStream = $webResponse.GetResponseStream()
$responseXml = [xml](new-object System.IO.StreamReader $responseStream).ReadToEnd()
$responseStream.Close()
$webResponse.Close()
$responseCode=$responseXml.get_InnerText()
if($responseCode -eq "NoError")
{
	return $true
}
else
{
    return $false
}