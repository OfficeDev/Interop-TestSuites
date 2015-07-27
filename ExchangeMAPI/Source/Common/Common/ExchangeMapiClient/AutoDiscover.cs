//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Public structure which contains auto discover properties
    /// </summary>
    public struct AutoDiscoverProperties
    {
        /// <summary>
        /// Server name for logon private mailbox
        /// </summary>
        public string PrivateMailboxServer;

        /// <summary>
        /// Proxy name for logon private mailbox
        /// </summary>
        public string PrivateMailboxProxy;

        /// <summary>
        /// Server name for logon public folder
        /// </summary>
        public string PublicMailboxServer;

        /// <summary>
        /// Proxy name for logon public folder
        /// </summary>
        public string PublicMailboxProxy;

        /// <summary>
        /// The URL that a client can use to connect with a private Mailbox through MAPI over HTTP
        /// </summary>
        public string PrivateMailStoreUrl;

        /// <summary>
        /// The URL that a client can use to connect with a public Mailbox through MAPI over HTTP
        /// </summary>
        public string PublicMailStoreUrl;
        
        /// <summary>
        /// The URL that a client can use to connect with a NSPI server through MAPI over HTTP.
        /// </summary>
        public string AddressBookUrl;
    }

    /// <summary>
    /// Static class which contains methods related to auto discover
    /// </summary>
    public static class AutoDiscover
    {
        /// <summary>
        /// Get auto discover properties for server name and proxy name
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="server">Server to connect.</param>
        /// <param name="userName">User name used to logon.</param>
        /// <param name="domain">Domain name.</param>
        /// <param name="requestURL">The server url address to receive the request from clien.</param>
        /// <param name="transport">The current transport used in the test suite.</param>
        /// <returns>Returns the structure contains auto discover properties.</returns>
        public static AutoDiscoverProperties GetAutoDiscoverProperties(
            ITestSite site,
            string server,
            string userName,
            string domain,
            string requestURL,
            string transport)
        {
            HttpStatusCode httpStatusCode = HttpStatusCode.Unused;
            XmlDocument doc = new XmlDocument();
            XmlNodeList elemList = null;
            string requestXML = string.Empty;
            string responseXML = string.Empty;

            AutoDiscoverProperties autoDiscoverProperties = new AutoDiscoverProperties
            {
                PrivateMailboxServer = null,
                PrivateMailboxProxy = null,

                PublicMailboxServer = null,
                PublicMailboxProxy = null,

                PublicMailStoreUrl = null,
                PrivateMailStoreUrl = null,

                AddressBookUrl = null
            };

            // Get auto discover properties for private mailbox
            requestXML = "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006\">" +
                "<Request><EMailAddress>" + userName + "@" + domain + "</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>";

            if (transport == "mapi_http")
            {
                httpStatusCode = SendHttpPostRequest(site, userName, domain, requestXML, requestURL, out responseXML, true);
                site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, httpStatusCode, "Http status code should be 200 (OK), the error code is {0}", httpStatusCode);
                doc.LoadXml(responseXML);

                elemList = doc.GetElementsByTagName("MailStore");
                site.Assert.IsTrue(elemList != null && elemList.Count > 0, "MailStore element should exist in response.");
                string mailStoreUrl = GetMAPIInternalURLProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(mailStoreUrl), "The mailstore internal URL should be gotten under MailStore element");
                autoDiscoverProperties.PrivateMailStoreUrl = mailStoreUrl; // Add private mailbox server

                elemList = doc.GetElementsByTagName("AddressBook");
                site.Assert.IsTrue(elemList != null && elemList.Count > 0, "AddressBook element should exist in response.");
                string addressBookUri = GetMAPIInternalURLProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(addressBookUri), "The AddressBook internal URL should be gotten under AddressBook element");
                autoDiscoverProperties.AddressBookUrl = addressBookUri;
            }
            else
            {
                httpStatusCode = SendHttpPostRequest(site, userName, domain, requestXML, requestURL, out responseXML, false);
                site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, httpStatusCode, "Http status code should be 200 (OK), the error code is {0}", httpStatusCode);
                doc.LoadXml(responseXML);

                elemList = doc.GetElementsByTagName("PublicFolderInformation");
                if (elemList == null || elemList.Count == 0)
                {
                    // Not found PublicFolderInformation means the SUT is not Microsoft Exchange Server 2013, in this case, proxy will not be changed and using original server name.
                    // For Microsoft Exchange Server 2010 the private mailbox server should be same as public folder server
                    autoDiscoverProperties.PrivateMailboxServer = server; // Add private mailbox server
                    autoDiscoverProperties.PublicMailboxServer = server; // Add public mailbox server

                    return autoDiscoverProperties;
                }

                // That the PublicFolderInformation is found means that the SUT is Microsoft Exchange Server 2013, in this case, get server name and proxy information for both private mailbox and public folder.
                elemList = doc.GetElementsByTagName("Protocol");
                site.Assert.IsTrue(elemList != null && elemList.Count > 0, "Protocol element should exist in response.");

                string privateMailboxServer = GetServerNameProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(privateMailboxServer), "The private mailbox server name should be gotten under Protocol element");

                string privateMailboxProxy = GetServerProxyProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(privateMailboxProxy), "The private mailbox server proxy should be gotten under Protocol element");

                autoDiscoverProperties.PrivateMailboxServer = privateMailboxServer; // Add private mailbox server
                autoDiscoverProperties.PrivateMailboxProxy = privateMailboxProxy; // Add private mailbox proxy
            }

            // Get auto discover properties for public mailbox
            string smtpAddress = "PublicFolderMailbox_" + server + "@" + domain;
            requestXML = "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006\">" +
                "<Request><EMailAddress>" + smtpAddress + "</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>";
            responseXML = string.Empty;

            if (transport == "mapi_http")
            {
                httpStatusCode = SendHttpPostRequest(site, userName, domain, requestXML, requestURL, out responseXML, true);
                site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, httpStatusCode, "Http status code should be 200 (OK), the error code is {0}", httpStatusCode);
                doc.LoadXml(responseXML);

                elemList = doc.GetElementsByTagName("MailStore");
                site.Assert.IsTrue(elemList != null && elemList.Count > 0, "MailStore element should exist in response.");

                string publicMailStoreUrl = GetMAPIInternalURLProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(publicMailStoreUrl), "The public folder mailbox server name should be gotten under MailStore element");

                autoDiscoverProperties.PublicMailStoreUrl = publicMailStoreUrl; // Add public mailbox server
            }
            else
            {
                httpStatusCode = SendHttpPostRequest(site, userName, domain, requestXML, requestURL, out responseXML, false);
                site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, httpStatusCode, "Http status code should be 200 (OK), the error code is {0}", httpStatusCode);
                doc.LoadXml(responseXML);
                elemList = doc.GetElementsByTagName("Protocol");
                site.Assert.IsTrue(elemList != null && elemList.Count > 0, "There should be an element called Protocol.");

                string publicMailboxServer = GetServerNameProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(publicMailboxServer), "The public folder mailbox server name should be gotten under Protocol element");

                string publicMailboxProxy = GetServerProxyProperty(elemList);
                site.Assert.IsFalse(string.IsNullOrEmpty(publicMailboxProxy), "The public folder mailbox server proxy should be gotten under Protocol element");

                autoDiscoverProperties.PublicMailboxServer = publicMailboxServer; // Add public mailbox server
                autoDiscoverProperties.PublicMailboxProxy = publicMailboxProxy; // Add public mailbox proxy
            }

            return autoDiscoverProperties;
        }

        /// <summary>
        /// Run the Http post method
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="userName">User name used to logon</param>
        /// <param name="domain">Domain name</param>
        /// <param name="requestXml">the request xml</param>
        /// <param name="url">the request send to</param>
        /// <param name="responseXml">The response xml</param>
        /// <param name="getMAPIURL">True indicates add headers to get MAPIHTTP url; default value is false</param>
        /// <returns>Return HttpStatusCode</returns>
        private static HttpStatusCode SendHttpPostRequest(ITestSite site, string userName, string domain, string requestXml, string url, out string responseXml, bool getMAPIURL)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback =
            new System.Net.Security.RemoteCertificateValidationCallback(Common.ValidateServerCertificate);
            HttpStatusCode httpStatusCode = HttpStatusCode.Created;
            responseXml = null;

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.CookieContainer = new CookieContainer();
                request.Method = "POST";
                request.Accept = "*/*";
                request.ContentType = "text/xml";
                request.Credentials = CredentialCache.DefaultCredentials;
                request.AllowAutoRedirect = false;
                request.KeepAlive = false;

                // Add headers to get MAPIHTTP url                        
                if (getMAPIURL)
                {
                    request.Headers.Add("X-MapiHttpCapability", "2");
                    request.Headers.Add("X-AnchorMailbox", userName + "@" + domain);
                }

                byte[] buffer = Encoding.UTF8.GetBytes(requestXml);
                request.ContentLength = buffer.Length;
                Stream webRequestStream = request.GetRequestStream();
                webRequestStream.Write(buffer, 0, buffer.Length);
                webRequestStream.Flush();
                webRequestStream.Dispose();
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                responseXml = reader.ReadToEnd();
                httpStatusCode = response.StatusCode;
                reader.Close();
                response.Close();

                if (httpStatusCode != HttpStatusCode.OK)
                {
                    site.Assert.Fail("Can't connect the server.");
                }
            }
            catch (WebException e)
            {
                site.Assert.Fail("A WebException happened when connecting the server. The error message is: {0}", e.Message.ToString());
            }

            return httpStatusCode;
        }

        /// <summary>
        /// Get server name from xml node
        /// </summary>
        /// <param name="elemList">xml node list</param>
        /// <returns>Server name</returns>
        private static string GetServerNameProperty(XmlNodeList elemList)
        {
            foreach (XmlNode xmlNode in elemList)
            {
                if (xmlNode.ChildNodes.Count > 0)
                {
                    if (string.Compare(xmlNode["Type"].InnerText, "EXCH", true) == 0)
                    {
                        if (xmlNode["Server"] != null)
                        {
                            return xmlNode["Server"].InnerText;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get server proxy from xml node
        /// </summary>
        /// <param name="elemList">xml node list</param>
        /// <returns>Server proxy name</returns>
        private static string GetServerProxyProperty(XmlNodeList elemList)
        {
            foreach (XmlNode xmlNode in elemList)
            {
                if (xmlNode.ChildNodes.Count > 0)
                {
                    if (string.Compare(xmlNode["Type"].InnerText, "EXPR", true) == 0)
                    {
                        if (xmlNode["Server"] != null)
                        {
                            return "RpcProxy=" + xmlNode["Server"].InnerText;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get mailstore internal url from xml node
        /// </summary>
        /// <param name="elemList">xml node list</param>
        /// <returns>Internal url of mailstore</returns>
        private static string GetMAPIInternalURLProperty(XmlNodeList elemList)
        {
            if (elemList != null && elemList.Count > 0)
            {
                foreach (XmlNode xmlNode in elemList)
                {
                    if (xmlNode.ChildNodes.Count > 0)
                    {
                        if (xmlNode["InternalUrl"].InnerText != null)
                        {
                            return xmlNode["InternalUrl"].InnerText;
                        }
                    }
                }
            }

            return null;
        }
    }
}
