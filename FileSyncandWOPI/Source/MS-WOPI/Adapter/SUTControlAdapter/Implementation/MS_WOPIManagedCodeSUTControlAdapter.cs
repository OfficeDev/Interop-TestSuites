namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Specialized;
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Security;
    using System.Security.Principal;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Web;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the IMS-WOPIManagedCodeSUTControlAdapter interface.
    /// </summary>
    public class MS_WOPIManagedCodeSUTControlAdapter : ManagedAdapterBase, IMS_WOPIManagedCodeSUTControlAdapter
    {
        /// <summary>
        /// A TransportProtocol type value represents the transport used by this test suite.
        /// </summary>
        private TransportProtocol currentTransport;

        #region properties

        /// <summary>
        /// Gets or sets the pattern of a URL which is used to request a WOPI view for a file.
        /// </summary>
        protected string PatternOfRequestWOPIViewFile { get; set; }

        /// <summary>
        /// Gets or sets the pattern of a URL which is used to request a WOPI view for a folder.
        /// </summary>
        protected string PatternOfRequestWOPIViewFolder { get; set; }

        #endregion

        #region Managed code sut contol adapter methods implementation

        /// <summary>
        /// The Overridden Initialize method, it includes the initialization logic of this adapter.
        /// </summary>
        /// <param name="testSite">The ITestSite member of ManagedAdapterBase</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (TransportProtocol.HTTPS == transport)
            {
                Common.AcceptServerCertificate();
            }

            // In this moment, the thread user should not be impersonated.
            if (null != WindowsIdentity.GetCurrent(true))
            {
                string errorMsg = string.Format("Current thread[{0}]: The thread user should not be impersonated in adapter initialization.", Thread.CurrentThread.ManagedThreadId);
                throw new InvalidOperationException(errorMsg);
            }

            this.PatternOfRequestWOPIViewFile = Common.GetConfigurationPropertyValue("RequestWOPIViewFileUrlPattern", this.Site);
            this.PatternOfRequestWOPIViewFolder = Common.GetConfigurationPropertyValue("RequestWOPIViewFolderUrlPattern", this.Site);
            this.currentTransport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
        }

        /// <summary>
        /// A method used to get the WOPI resource URL by specified user credentials.
        /// </summary>
        /// <param name="absoluteUrlOfResource">A parameter represents the absolute URL of normal resource which will be used to get a WOPI resource URL.</param>
        /// <param name="rootResourceType">A parameter indicating the WOPI root resource URL type will be return.</param>
        /// <param name="userName">A parameter represents the name of user whose associated token will be return in the WOPI resource URL.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>A return value represents the WOPI resource URL, which can be used in MS-WOPI operations.</returns>
        public string GetWOPIRootResourceUrl(string absoluteUrlOfResource, WOPIRootResourceUrlType rootResourceType, string userName, string password, string domain)
        {
            #region Verify the parameter

            if (string.IsNullOrEmpty(absoluteUrlOfResource))
            {
                throw new ArgumentNullException("absoluteUrlOfResource");
            }

            if (string.IsNullOrEmpty(userName))
            {
                throw new ArgumentNullException("userName");
            }

            if (string.IsNullOrEmpty(password))
            {
                throw new ArgumentNullException("password");
            }

            if (string.IsNullOrEmpty(domain))
            {
                throw new ArgumentNullException("domain");
            }

            if (WOPIRootResourceUrlType.FileLevel != rootResourceType && WOPIRootResourceUrlType.FolderLevel != rootResourceType)
            {
                throw new NotSupportedException(string.Format(@"The test suite only supports [{0}] and [{1}] two WOPI root resource URL formats.", WOPIRootResourceUrlType.FileLevel, WOPIRootResourceUrlType.FolderLevel));
            }

            #endregion

            string wopiResourceUrl = string.Empty;

            if (!WOPIResourceUrlCache.TryGetWOPIResourceUrl(userName, domain, absoluteUrlOfResource, out wopiResourceUrl))
            {
                // Get the response by specified user credential
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "Try to get the [{0}] type WOPI URL for resource:[{1}]\r\n User:[{2}]\\[{3}]",
                    rootResourceType,
                    absoluteUrlOfResource,
                    domain,
                    userName);

                string triggerWOPIViewUrl = this.GenerateRequestWOPIViewUrl(absoluteUrlOfResource, rootResourceType);
                string htmlReponseOfViewFile = this.TriggerRequestViewByUrl(triggerWOPIViewUrl, userName, password, domain);

                // Get the token
                string token = this.GetTokenByResponseOfRequestView(htmlReponseOfViewFile);
                token = HttpUtility.UrlEncode(token);

                // Get the WOPISrc value
                string wopiSrc = this.GetWOPISrcValueByResponseOfRequestViewFile(htmlReponseOfViewFile);

                // Construct the absolute request URL.
                wopiResourceUrl = this.GetAbsoluteRequestUrlWithToken(wopiSrc, token);

                // Verify the WOPI root resource URL.
                if (string.IsNullOrEmpty(wopiResourceUrl))
                {
                    string errorMsg = string.Format("Could not get the WOPI resource URL for a [{0}] type resource[{1}]", rootResourceType, absoluteUrlOfResource);
                    throw new InvalidOperationException(errorMsg);
                }

                Uri wopiUrl;
                if (!Uri.TryCreate(wopiResourceUrl, UriKind.Absolute, out wopiUrl))
                {
                    string errorMsg = string.Format("The WOPI resource URL [{0}] is not valid absolute URL.", wopiResourceUrl);
                    throw new InvalidOperationException(errorMsg);
                }

                string currentTransportValue = TokenAndRequestUrlHelper.CurrentHttpTransport.ToString();
                if (!currentTransportValue.Equals(wopiUrl.Scheme, StringComparison.OrdinalIgnoreCase))
                {
                    wopiResourceUrl = string.Format(
                                                    @"{0}://{1}{2}{3}",
                                                    currentTransportValue,
                                                    wopiUrl.Host,
                                                    wopiUrl.LocalPath,
                                                    wopiUrl.Query);
                }

                WOPIResourceUrlCache.AddWOPIResourceUrl(userName, domain, absoluteUrlOfResource, wopiResourceUrl);
            }

            return wopiResourceUrl;
        }

        #endregion

        #region protect method

        /// <summary>
        /// A method is used to get the token value from a html response which is returned by WOPI server when it receive a WOPI view request.
        /// </summary>
        /// <param name="responseOfRequestViewResource">A parameter represents the html response which should contains token and WOPIsrc URL parameter.</param>
        /// <returns>A return value represents the token value</returns>
        protected virtual string GetTokenByResponseOfRequestView(string responseOfRequestViewResource)
        {
            if (string.IsNullOrEmpty(responseOfRequestViewResource))
            {
                throw new ArgumentNullException("responseOfRequestViewResource");
            }

            Regex accessTokenRegex = new Regex(@"<form.*action=""(?<action>.*?)"".*(\n.*)*?<input.*name=""access_token"".*value=""(?<token>.*?)"".*/>");
            Match match = accessTokenRegex.Match(responseOfRequestViewResource);
            string token = match.Groups["token"].Value;

            if (string.IsNullOrEmpty(token))
            {
                string errorMsg = string.Format(@"Could not get the token value from the response of request a WOPI file view:\r\n{0}", responseOfRequestViewResource);
                throw new InvalidOperationException(errorMsg);
            }

            return token;
        }

        /// <summary>
        /// A method is used to get the WOPISrc URL parameter value from a html response which is returned by WOPI server when it receive a WOPI view request.
        /// </summary>
        /// <param name="responseOfRequestViewFile">A parameter represents the html response which should contains token and WOPIsrc URL parameter.</param>
        /// <returns>A return value represents the WOPISrc value</returns>
        protected virtual string GetWOPISrcValueByResponseOfRequestViewFile(string responseOfRequestViewFile)
        {
            if (string.IsNullOrEmpty(responseOfRequestViewFile))
            {
                throw new ArgumentNullException("responseOfRequestViewFile");
            }

            Regex accessTokenRegex = new Regex(@"<form.*action=""(?<action>.*?)"".*(\n.*)*?<input.*name=""access_token"".*value=""(?<token>.*?)"".*/>");
            Match match = accessTokenRegex.Match(responseOfRequestViewFile);
            string actionOfForm = match.Groups["action"].Value;

            if (string.IsNullOrEmpty(actionOfForm))
            {
                string errorMsg = string.Format(@"Could not get the action value of the Form from the response of request a WOPI file view:\r\n{0}", responseOfRequestViewFile);
                throw new InvalidOperationException(errorMsg);
            }

            NameValueCollection queries = HttpUtility.ParseQueryString(actionOfForm);
            string wopiSrc = queries["WOPISrc"];
            wopiSrc = HttpUtility.UrlDecode(wopiSrc);

            if (string.IsNullOrEmpty(wopiSrc))
            {
                string errorMsg = string.Format(@"Could not get the WOPISrc value from the action value [{0}]", actionOfForm);
                throw new InvalidOperationException(errorMsg);
            }

            return wopiSrc;
        }

        /// <summary>
        /// A method is used to generate request URL for viewing a normal resource in WOPI mode.
        /// </summary>
        /// <param name="absoluteUrlOfResource">A parameter represents the absolute URL of a normal resource.</param>
        /// <param name="rootResourceType">A parameter indicating the WOPI root resource URL type.</param>
        /// <returns>A return value represents the request URL for viewing a normal resource in WOPI mode. </returns>
        protected virtual string GenerateRequestWOPIViewUrl(string absoluteUrlOfResource, WOPIRootResourceUrlType rootResourceType)
        {
            if (string.IsNullOrEmpty(absoluteUrlOfResource))
            {
                throw new ArgumentNullException("absoluteUrlOfResource");
            }

            string placeHolderInPattern = string.Empty;
            string patternValue = string.Empty;

            switch (rootResourceType)
            {
                case WOPIRootResourceUrlType.FileLevel:
                    {
                        placeHolderInPattern = @"[filepath]";
                        patternValue = this.PatternOfRequestWOPIViewFile;
                        break;
                    }

                case WOPIRootResourceUrlType.FolderLevel:
                    {
                        placeHolderInPattern = @"[folderpath]";
                        patternValue = this.PatternOfRequestWOPIViewFolder;
                        break;
                    }

                default:
                    {
                        throw new InvalidOperationException("The test suite only supports folder and file two WOPI root resource URL formats.");
                    }
            }

            // Verify whether the expected placeHolder exists in pattern.  
            if (patternValue.IndexOf(placeHolderInPattern, StringComparison.OrdinalIgnoreCase) < 0)
            {
                string errorMsg = string.Format(
                                                "Could not find the [{0}] place holder, this place holder indicating the file path to view. Default format is [TargetSiteCollectionUrl]_layouts/15/WopiFrame.aspx?sourcedoc={0}",
                                                placeHolderInPattern);
                throw new InvalidOperationException(errorMsg);
            }

            // Ensure the http transport in "absoluteUrlOfResource" parameter should match the current transport used by this test suite.
            Uri targetResourceNormalPath;
            if (!Uri.TryCreate(absoluteUrlOfResource, UriKind.Absolute, out targetResourceNormalPath))
            {
                string errorMsg = string.Format(
                                                "The target resource URL should be a valid absolute URL. Current[{0}]",
                                                absoluteUrlOfResource);
                throw new InvalidOperationException(errorMsg);
            }

            absoluteUrlOfResource = absoluteUrlOfResource.ToLower(CultureInfo.CurrentCulture);
            if (!targetResourceNormalPath.Scheme.Equals(this.currentTransport.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                string needReplaceTransportvalue = this.currentTransport.Equals(TransportProtocol.HTTP) ? "https" : "http";
                needReplaceTransportvalue = string.Format(@"{0}://", needReplaceTransportvalue);
                string newReplaceValue = string.Format(@"{0}://", this.currentTransport);
                absoluteUrlOfResource = absoluteUrlOfResource.Replace(needReplaceTransportvalue, newReplaceValue);
            }

            patternValue = patternValue.ToLower(CultureInfo.CurrentCulture).Replace(placeHolderInPattern, "{0}");
            absoluteUrlOfResource = HttpUtility.UrlEncode(absoluteUrlOfResource);
            string requestWopiViewUrlValue = string.Format(patternValue, absoluteUrlOfResource);
            return requestWopiViewUrlValue;
        }

        /// <summary>
        /// A method is used to get the absolute request URL with the token.
        /// </summary>
        /// <param name="wopiResourceUrl">The value of the WOPI request URL.</param>
        /// <param name="tokenValue">The value of the token.</param>
        /// <returns>The absolute request URL.</returns>
        protected string GetAbsoluteRequestUrlWithToken(string wopiResourceUrl, string tokenValue)
        {
            if (string.IsNullOrEmpty(wopiResourceUrl))
            {
                throw new ArgumentNullException("wopiResourceUrl");
            }

            if (string.IsNullOrEmpty(tokenValue))
            {
                throw new ArgumentNullException("tokenValue");
            }

            string formatedRequestUrl = string.Format(
                                                @"{0}?access_token={1}",
                                                wopiResourceUrl,
                                                tokenValue);
            return formatedRequestUrl;
        }

        /// <summary>
        /// A method is used to send a WOPI view request to the WOPI server and get a html response.
        /// </summary>
        /// <param name="requestViewResourceUrl">A parameter represents the request URL for viewing a resource by using WOPI mode.</param>
        /// <param name="userName">A parameter represents the name of user whose associated token will be return in the HTML content.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>A return value represents the html response from the WOPI server.</returns>
        protected string TriggerRequestViewByUrl(string requestViewResourceUrl, string userName, string password, string domain)
        {
            if (string.IsNullOrEmpty(requestViewResourceUrl))
            {
                throw new ArgumentException("requestViewFileUri");
            }

            // Secure string only support 65536 length.
            if (password.Length > 65536)
            {
                throw new ArgumentException("The Password length is larger than 65536, the test suite only support password less than 65536.");
            }

            string htmlContent = string.Empty;

            // Create the HTTP request with browser's setting.
            Uri requestTargetLocation = new Uri(requestViewResourceUrl);
            HttpWebRequest request = HttpWebRequest.Create(requestTargetLocation) as HttpWebRequest;
            request.Method = "GET";
            request.UserAgent = @"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";

            SecureString securePassword = new SecureString();
            foreach (char passwordCharItem in password)
            {
                securePassword.AppendChar(passwordCharItem);
            }

            NetworkCredential credentialInstance = new NetworkCredential(userName, securePassword, domain);
            CredentialCache credentialCache = new CredentialCache();
            credentialCache.Add(requestTargetLocation, "NTLM", credentialInstance);
            request.Credentials = credentialCache;

            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                htmlContent = this.ReadHtmlContentFromResponse(response);
            }

            return htmlContent;
        }

        /// <summary>
        /// A method is used to read html contents from the http response. If could not read the html contents, it will throw an exception.
        /// </summary>
        /// <param name="response">A parameter represents http response which should contain html contents.</param>
        /// <returns>A return value represents the read html contents from http response.</returns>
        protected string ReadHtmlContentFromResponse(HttpWebResponse response)
        {
            HelperBase.CheckInputParameterNullOrEmpty<HttpWebResponse>(response, "response", "ReadHtmlContentsFromResponse");

            string htmlContent = string.Empty;
            Stream stream = null;
            try
            {
                stream = response.GetResponseStream();
                Encoding encoding = Encoding.GetEncoding(response.CharacterSet);
                using (StreamReader reader = new StreamReader(stream, encoding))
                {
                    htmlContent = reader.ReadToEnd();
                }
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }

            if (string.IsNullOrEmpty(htmlContent))
            {
                string errorMsg = string.Format(@"Could not get the request view file's response from the WOPI server.");

                throw new InvalidOperationException(errorMsg);
            }

            return htmlContent;
        }

        #endregion
    }
}