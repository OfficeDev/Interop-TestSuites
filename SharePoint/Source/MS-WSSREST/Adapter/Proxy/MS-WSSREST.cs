//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System.IO;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The proxy of MS-WSSREST protocol.
    /// </summary>
    public class MS_WSSREST
    {
        /// <summary>
        /// An object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        private ITestSite testsite;

        /// <summary>
        /// Initialize the instance of ITestSite.
        /// </summary>
        /// <param name="site">A object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public void Initialize(ITestSite site)
        {
            this.testsite = site;
        }

        /// <summary>
        /// Send http request with specified method and header to get the corresponding response.
        /// </summary>
        /// <param name="method">Http request method.</param>
        /// <param name="request">The header and content.</param>
        /// <returns>The response from server.</returns>
        public HttpWebResponse SendMessage(HttpMethod method, Request request)
        {
            HttpWebResponse response = null;
            string url = string.Format("{0}/{1}", Common.GetConfigurationPropertyValue("TargetServiceUrl", this.testsite), request.Parameter);
            HttpWebRequest webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Method = method.ToString();
            webRequest.Accept = request.Accept;
            webRequest.ContentType = request.ContentType;

            if (method == HttpMethod.POST && !string.IsNullOrEmpty(request.Slug))
            {
                webRequest.Headers.Add("slug", request.Slug);
            }

            if (method == HttpMethod.PUT || method == HttpMethod.MERGE)
            {
                webRequest.Headers.Add("If-Match", request.ETag);
            }

            webRequest.Credentials = new NetworkCredential(
                Common.GetConfigurationPropertyValue("UserName", this.testsite),
                Common.GetConfigurationPropertyValue("Password", this.testsite),
                Common.GetConfigurationPropertyValue("Domain", this.testsite));

            if (Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.testsite) == TransportProtocol.HTTPS)
            {
                Common.AcceptServerCertificate();
            }

            if (!string.IsNullOrEmpty(request.Content))
            {
                byte[] postBytes = Encoding.ASCII.GetBytes(request.Content);
                Stream postStream = webRequest.GetRequestStream();
                postStream.Write(postBytes, 0, postBytes.Length);
                postStream.Close();
            }

            response = webRequest.GetResponse() as HttpWebResponse;

            return response;
        }
    }
}
