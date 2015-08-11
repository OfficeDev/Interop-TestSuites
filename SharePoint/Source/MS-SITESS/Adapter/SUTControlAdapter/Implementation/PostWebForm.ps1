$script:ErrorActionPreference = "Stop"
$domain = .\Get-ConfigurationPropertyValue.ps1 Domain
$userName = .\Get-ConfigurationPropertyValue.ps1 UserName
$password = .\Get-ConfigurationPropertyValue.ps1 Password
$str = .\Get-ConfigurationPropertyValue.ps1 RegularExpression
$form = $null;
$url = $webPageUrl;

$wc = new-object System.Net.WebClient;
$wc.Credentials = new-object System.Net.NetworkCredential($userName, $password, $domain);
$page = $wc.DownloadString($url);
if (![string]::IsNullOrEmpty($page))
{
    $form = new-object System.Collections.Specialized.NameValueCollection;
    $reg = new-object System.Text.RegularExpressions.Regex($str);
    foreach ($match in $reg.Matches($page))
    {
        $form.Add($match.Groups["name"].Value, $match.Groups["value"].Value);
    }
}

if (![string]::IsNullOrEmpty($digest))
{
	$form["__REQUESTDIGEST"] = $digest;
}

$s = ""
try{
	$response = $wc.UploadValues($url, $form);
	$s = [System.Text.Encoding]::UTF8.GetString($response);
}
catch [Net.WebException]
{
	$s = $_.tostring()
}
return $s