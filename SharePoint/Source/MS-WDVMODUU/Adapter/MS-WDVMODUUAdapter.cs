namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System;
    using System.Collections.Specialized;
    using System.IO;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-WDVMODUU.
    /// </summary>
    public partial class MS_WDVMODUUAdapter : ManagedAdapterBase, IMS_WDVMODUUAdapter
    {
        #region Variables

        /// <summary>
        /// The name of the user which is used to invoke protocol methods.
        /// </summary>
        private string userName;

        /// <summary>
        /// The domain of the user which is used to invoke protocol methods.
        /// </summary>
        private string domain;

        /// <summary>
        /// The password of the user which is used to invoke protocol methods.
        /// </summary>
        private string password;

        /// <summary>
        /// The last raw HTTP Request string (include the HTTP headers and body content if exists) that send to the protocol server.
        /// </summary>
        private string lastRawRequest;

        /// <summary>
        /// The last raw HTTP Response string (include the HTTP headers and XML data if exists) from the protocol server.
        /// </summary>
        private string lastRawResponse;

        #endregion Variables

        #region Properties Implementation in IMS_WDVMODUUAdapter interface
        /// <summary>
        /// Gets the last raw HTTP Response string (include the HTTP headers and XML data if exists) from the protocol server.
        /// </summary>
        public string LastRawResponse
        {
            get
            {
                return this.lastRawResponse;
            }
        }

        /// <summary>
        /// Gets the last raw HTTP Request string (include the HTTP headers and body content if exists) that send to the protocol server.
        /// </summary>
        public string LastRawRequest
        {
            get
            {
                return this.lastRawRequest;
            }
        }
        #endregion  Properties Implementation in IMS_WDVMODUUAdapter interface

        #region Initialize TestSuite

        /// <summary>
        /// Override IAdapter's Initialize(), set default protocol short name of the testSite, and initialize variables in "MS_WDVMODUUAdapter" class.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set default protocol short name of the testSite
            testSite.DefaultProtocolDocShortName = "MS-WDVMODUU";

            // Merge the common configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, false);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);

            // Initialize variables in "MS_WDVMODUUAdapter" class.
            this.userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.password = Common.GetConfigurationPropertyValue("Password", testSite);
            this.lastRawRequest = string.Empty;
            this.lastRawResponse = string.Empty;
        }

        #endregion

        #region Methods Implementation in IMS_WDVMODUUAdapter interface

        #region PropFind method

        /// <summary>
        /// The method is used to send a HTTP request using PROPFIND method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="body">The body content in the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        public WDVMODUUResponse PropFind(string requestUri, string body, NameValueCollection headersCollection)
        {
            // Construct an "HttpWebRequest" object based on the input request URI.
            HttpWebRequest httpWebRequest = this.ConstructHttpWebRequest(requestUri);

            // Specify the method in the HTTP request.
            httpWebRequest.Method = "PROPFIND";

            // Set the HTTP headers in the HTTP request from inputs.
            foreach (string name in headersCollection.AllKeys)
            {
                if (name.Equals("User-Agent"))
                {
                    httpWebRequest.UserAgent = headersCollection[name];
                }
                else if (name.Equals("ContentType"))
                {
                    httpWebRequest.ContentType = headersCollection[name];
                }
                else
                {
                    httpWebRequest.Headers.Set(name, headersCollection[name]);
                }
            }

            // Encode the body using UTF-8.
            byte[] bytes = Encoding.UTF8.GetBytes((string)body);

            // Set the HTTP header of content length. 
            httpWebRequest.ContentLength = bytes.Length;

            // Get a reference to the request stream.
            Stream requestStream = httpWebRequest.GetRequestStream();

            // Write the request body to the request stream.
            requestStream.Write(bytes, 0, bytes.Length);

            // Close the Stream object to release the connection for further use.
            requestStream.Close();

            // Reserve the last HTTP Request Data.
            this.ReserveRequestData(httpWebRequest, body);

            // Send the PROPFIND method request and get the response from the protocol server.
            HttpWebResponse httpWebResponse = this.GetResponse(httpWebRequest);
            this.Site.Assert.IsNotNull(httpWebResponse, "The 'HttpWebResponse' object should not be null!");

            // Reserve the last HTTP Response Data.
            WDVMODUUResponse responseWDVMODUU = new WDVMODUUResponse();
            this.lastRawResponse = responseWDVMODUU.ReserveResponseData(httpWebResponse, this.Site);
            this.AssertWDVMODUUResponse(responseWDVMODUU);

            // Return the "WDVMODUUResponse" Object.
            return responseWDVMODUU;
        }

        #endregion PropFind method

        #region Put method

        /// <summary>
        /// The method is used to send a HTTP request using PUT method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="body">The body content in the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        public WDVMODUUResponse Put(string requestUri, byte[] body, NameValueCollection headersCollection)
        {
            // Construct an "HttpWebRequest" object based on the input request URI.
            HttpWebRequest httpWebRequest = this.ConstructHttpWebRequest(requestUri);

            // Specify the method in the HTTP request.
            httpWebRequest.Method = "PUT";

            // Set the HTTP headers in the HTTP request from inputs.
            foreach (string name in headersCollection.AllKeys)
            {
                if (name.Equals("User-Agent"))
                {
                    httpWebRequest.UserAgent = headersCollection[name];
                }
                else if (name.Equals("ContentType"))
                {
                    httpWebRequest.ContentType = headersCollection[name];
                }
                else
                {
                    httpWebRequest.Headers.Set(name, headersCollection[name]);
                }
            }

            // Set the content header length. 
            httpWebRequest.ContentLength = body.Length;

            // Get a reference to the request stream.
            Stream requestStream = httpWebRequest.GetRequestStream();

            // Write the request body to the request stream.
            requestStream.Write(body, 0, body.Length);

            // Close the Stream object to release the connection.
            requestStream.Close();

            // Reserve the last HTTP Request Data.
            this.ReserveRequestData(httpWebRequest, Encoding.UTF8.GetString(body));

            // Send the PUT method request and get the response from the protocol server.
            HttpWebResponse httpWebResponse = this.GetResponse(httpWebRequest);
            this.Site.Assert.IsNotNull(httpWebResponse, "The 'HttpWebResponse' object should not be null!");

            // Reserve the last HTTP Response Data.
            WDVMODUUResponse responseWDVMODUU = new WDVMODUUResponse();
            this.lastRawResponse = responseWDVMODUU.ReserveResponseData(httpWebResponse, this.Site);
            this.AssertWDVMODUUResponse(responseWDVMODUU);

            // Return the "WDVMODUUResponse" Object.
            return responseWDVMODUU;
        }

        #endregion Put method

        #region Get method

        /// <summary>
        /// The method is used to send a HTTP request using GET method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        public WDVMODUUResponse Get(string requestUri, NameValueCollection headersCollection)
        {
            // Construct an "HttpWebRequest" object based on the input request URI.
            HttpWebRequest httpWebRequest = this.ConstructHttpWebRequest(requestUri);

            // Specify the method in the HTTP request.
            httpWebRequest.Method = "GET";

            // Set the HTTP headers in the HTTP request from inputs.
            foreach (string name in headersCollection.AllKeys)
            {
                if (name.Equals("User-Agent"))
                {
                    httpWebRequest.UserAgent = headersCollection[name];
                }
                else if (name.Equals("ContentType"))
                {
                    httpWebRequest.ContentType = headersCollection[name];
                }
                else
                {
                    httpWebRequest.Headers.Set(name, headersCollection[name]);
                }
            }

            // Reserve the last HTTP Request Data.
            this.ReserveRequestData(httpWebRequest, null);

            // Send the GET method request and get the response from the protocol server.
            HttpWebResponse httpWebResponse = this.GetResponse(httpWebRequest);
            this.Site.Assert.IsNotNull(httpWebResponse, "The 'HttpWebResponse' object should not be null!");

            // Reserve the last HTTP Response Data.
            WDVMODUUResponse responseWDVMODUU = new WDVMODUUResponse();
            this.lastRawResponse = responseWDVMODUU.ReserveResponseData(httpWebResponse, this.Site);
            this.AssertWDVMODUUResponse(responseWDVMODUU);

            // Return the "WDVMODUUResponse" Object.
            return responseWDVMODUU;
        }

        #endregion Get method

        #region Delete method

        /// <summary>
        /// The method is used to send a HTTP request using DELETE method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        public WDVMODUUResponse Delete(string requestUri, NameValueCollection headersCollection)
        {
            // Construct an "HttpWebRequest" object based on the input request URI.
            HttpWebRequest httpWebRequest = this.ConstructHttpWebRequest(requestUri);

            // Specify the method in the HTTP request.
            httpWebRequest.Method = "DELETE";

            // Set the HTTP headers in the HTTP request from inputs.
            foreach (string name in headersCollection.AllKeys)
            {
                if (name.Equals("User-Agent"))
                {
                    httpWebRequest.UserAgent = headersCollection[name];
                }
                else if (name.Equals("ContentType"))
                {
                    httpWebRequest.ContentType = headersCollection[name];
                }
                else
                {
                    httpWebRequest.Headers.Set(name, headersCollection[name]);
                }
            }

            // Reserve the last HTTP Request Data.
            this.ReserveRequestData(httpWebRequest, null);

            // Send the DELETE method request and get the response from the protocol server.
            HttpWebResponse httpWebResponse = this.GetResponse(httpWebRequest);
            this.Site.Assert.IsNotNull(httpWebResponse, "The 'HttpWebResponse' object should not be null!");

            // Reserve the last HTTP Response Data.
            WDVMODUUResponse responseWDVMODUU = new WDVMODUUResponse();
            this.lastRawResponse = responseWDVMODUU.ReserveResponseData(httpWebResponse, this.Site);
            this.AssertWDVMODUUResponse(responseWDVMODUU);

            // Return the "WDVMODUUResponse" Object.
            return responseWDVMODUU;
        }

        #endregion Delete method

        #endregion Methods Implementation in IMS_WDVMODUUAdapter interface

        #region Private helper methods

        #region ReserveRequestData Method
        /// <summary>
        /// Reserve the HTTP headers and body content from the HTTP request.
        /// </summary>
        /// <param name="httpRequest">The HTTP Request object.</param>
        /// <param name="httpBodyString">The Http Body string in "httpRequest". </param>
        private void ReserveRequestData(HttpWebRequest httpRequest, string httpBodyString)
        {
            StringBuilder stringBuilder = new StringBuilder();

            // Get the HTTP request title.
            stringBuilder.AppendLine(string.Format("{0} {1} HTTP/{2}", httpRequest.Method, httpRequest.RequestUri.OriginalString.Replace(" ", "20%"), httpRequest.ProtocolVersion));

            // Get the headers in the HTTP request.
            foreach (string key in httpRequest.Headers.AllKeys)
            {
                stringBuilder.AppendLine(string.Format("{0}:{1}", key, httpRequest.Headers[key]));
            }

            // Get the body content in the HTTP request.
            if (httpBodyString != string.Empty)
            {
                stringBuilder.AppendLine(httpBodyString);
            }

            // Reserve the HTTP headers and body content.
            this.lastRawRequest = stringBuilder.ToString();

            // Output raw request in test log.
            this.Site.Log.Add(LogEntryKind.Debug, "The raw request message is: \r\n", this.lastRawRequest);
        }
        #endregion

        #region GetResponse method

        /// <summary>
        /// Get the HTTP response via sending the HTTP request in input parameter "httpRequest".
        /// </summary>
        /// <param name="httpRequest">The HTTP request that will send to the protocol server.</param>
        /// <returns>The HTTP response that received from the protocol server.</returns>
        private HttpWebResponse GetResponse(HttpWebRequest httpRequest)
        {
            HttpWebResponse httpWebResponse = null;
            try
            {
                httpWebResponse = (HttpWebResponse)httpRequest.GetResponse();
            }
            catch (WebException webExceptiion)
            {
                this.Site.Log.Add(LogEntryKind.Comment, "Test Comment: Get following web exception from the server: \r\n{0}", webExceptiion.Message);
                httpWebResponse = (HttpWebResponse)webExceptiion.Response;
                WDVMODUUResponse response = new WDVMODUUResponse();
                this.lastRawResponse = response.ReserveResponseData(httpWebResponse, this.Site);

                // Throw the web exception to test cases.
                throw webExceptiion;
            }

            this.ValidateAndCaptureTransport(httpWebResponse);
            return httpWebResponse;
        }
        #endregion 

        #region ConstructHttpWebRequest Method

        /// <summary>
        /// Construct an "HttpWebRequest" object based on a request URI.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <returns>Return the created "HttpWebRequest" object based on the request URI.</returns>
        private HttpWebRequest ConstructHttpWebRequest(string requestUri)
        {
            this.Site.Assert.IsFalse(string.IsNullOrEmpty(requestUri), "The input request URI should not null or empty.");
            CredentialCache credentials = new CredentialCache();
            Uri resourceUri = new Uri(requestUri);
            credentials.Add(resourceUri, "NTLM", new NetworkCredential(this.userName, this.password, this.domain));
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(resourceUri);
            httpWebRequest.Credentials = credentials;
            httpWebRequest.PreAuthenticate = true;
            return httpWebRequest;
        }

        #endregion

        #region AssertWDVMODUUResponse Method
        /// <summary>
        /// Assert the return object 'WDVMODUUResponse' is not null, make sure all necessary members in the object are not null.
        /// </summary>
        /// <param name="responseWDVMODUU">The return object 'WDVMODUUResponse' that will be checked.</param>
        private void AssertWDVMODUUResponse(WDVMODUUResponse responseWDVMODUU)
        {
            this.Site.Assert.IsNotNull(responseWDVMODUU, "The response object 'responseWDVMODUU' should not be null!");
            this.Site.Assert.IsNotNull(responseWDVMODUU.HttpHeaders, "The response object 'responseWDVMODUU.HttpHeaders' should not be null!");
            this.Site.Assert.IsNotNull(responseWDVMODUU.ProtocolVersion, "The response object 'responseWDVMODUU.ProtocolVersion' should not be null!");
        }
        #endregion

        #endregion Private helper methods
    }
}