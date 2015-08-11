namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class is used as the return type for Adapter methods, 
    /// it contains all data in the HTTP response by parsing the "HttpWebResponse" object, 
    /// and the data are returned and used by the test cases.
    /// </summary>
    public class WDVMODUUResponse
    {
        #region Private fields

        /// <summary>
        /// The protocol version in the HTTP response.
        /// </summary>
        private Version protocolVersion;

        /// <summary>
        /// The status code in the HTTP response.
        /// </summary>
        private HttpStatusCode statusCode;

        /// <summary>
        /// The status description in the HTTP response.
        /// </summary>
        private string statusDescription;

        /// <summary>
        /// The content length in the HTTP response.
        /// </summary>
        private long contentLength;

        /// <summary>
        /// The content type in the HTTP response.
        /// </summary>
        private string contentType;

        /// <summary>
        /// The HTTP header collection in the HTTP response.
        /// </summary>
        private WebHeaderCollection httpHeaders;

        /// <summary>
        /// The XML data in the HTTP response. If there is no XML data in the HTTP response, the field value will be set to null.
        /// </summary>
        private XmlDocument xmlData;

        /// <summary>
        /// The body data in the HTTP response.
        /// </summary>
        private string bodyData;

        #endregion 

        /// <summary>
        /// Initializes a new instance of the WDVMODUUResponse class.
        /// </summary>
        public WDVMODUUResponse()
        {
            this.protocolVersion = null;
            this.httpHeaders = null;
            this.xmlData = null;
            this.bodyData = null;
        }

        #region Public properties

        /// <summary>
        /// Gets the protocol version in the HTTP response.
        /// </summary>
        public Version ProtocolVersion
        {
            get
            {
                return this.protocolVersion;
            }
        }

        /// <summary>
        /// Gets the status code in the HTTP response.
        /// </summary>
        public HttpStatusCode StatusCode
        {
            get
            {
                return this.statusCode;
            }
        }

        /// <summary>
        /// Gets the status description in the HTTP response.
        /// </summary>
        public string StatusDescription
        {
            get
            {
                return this.statusDescription;
            }
        }

        /// <summary>
        /// Gets the content length in the HTTP response.
        /// </summary>
        public long ContentLength
        {
            get
            {
                return this.contentLength;
            }
        }

        /// <summary>
        /// Gets content type in the HTTP response.
        /// </summary>
        public string ContentType
        {
            get
            {
                return this.contentType;
            }
        }

        /// <summary>
        /// Gets the HTTP header collection in the HTTP response.
        /// </summary>
        public WebHeaderCollection HttpHeaders
        {
            get
            {
                return this.httpHeaders;
            }
        }

        /// <summary>
        /// Gets the XML data in the HTTP response.
        /// </summary>
        public XmlDocument BodyXmlData
        {
            get
            {
                return this.xmlData;
            }
        }

        /// <summary>
        /// Gets the body data in the HTTP response.
        /// </summary>
        public string BodyData
        {
            get
            {
                return this.bodyData;
            }
        }

        #endregion Public properties

        /// <summary>
        /// Reserve the HTTP headers and Body data from the HTTP Web Response object.
        /// </summary>
        /// <param name="httpResponse">The HTTP Web Response object.</param>
        /// <param name="site">The ITestSite object is used in log.</param>
        /// <returns>Return the raw HTTP response in a string from the input HTTP Web Response object.</returns>
        public string ReserveResponseData(HttpWebResponse httpResponse, ITestSite site)
        {
            if (site == null)
            {
                return string.Empty;
            }

            site.Assert.IsNotNull(httpResponse, "In method 'ReserveResponseData', the input parameter 'httpResponse' should not be null.");

            string lastRawResponse = string.Empty;
            StringBuilder stringBilder = new StringBuilder();

            // Get and reserve the HTTP response title.
            this.statusCode = httpResponse.StatusCode;
            this.statusDescription = httpResponse.StatusDescription;
            this.protocolVersion = httpResponse.ProtocolVersion;
            stringBilder.AppendLine(string.Format(
                "HTTP/{0} {1} {2}",
                  this.protocolVersion.ToString(),
                (int)this.statusCode,
                this.statusDescription));

            // Get and reserve the headers in the HTTP response.
            this.contentType = httpResponse.ContentType;
            this.contentLength = httpResponse.ContentLength;
            this.httpHeaders = httpResponse.Headers;
            foreach (string key in this.httpHeaders.AllKeys)
            {
                stringBilder.AppendLine(string.Format("{0}:{1}", key, httpResponse.Headers[key]));
            }

            // Get and reserve the body data in the HTTP response.
            if (this.contentLength > 0) 
            {
                // Get and reserve the XML data in the HTTP response when the content body is "XML" and is not empty.
                if (this.contentType.ToLower().IndexOf("xml", 0) >= 0)
                {
                    using (Stream responseStream = httpResponse.GetResponseStream())
                    {
                        this.xmlData = new XmlDocument();

                        // Reserve the XML data in the HTTP response.
                        this.xmlData.Load(responseStream);

                        // Create StringWriter object to get data from XML document.
                        using (StringWriter responseStringWriter = new StringWriter())
                        {
                            using (XmlTextWriter responseXmlTextWriter = new XmlTextWriter(responseStringWriter))
                            {
                                this.xmlData.WriteTo(responseXmlTextWriter);
                                this.bodyData = responseStringWriter.ToString();
                                stringBilder.AppendLine(this.bodyData);
                                responseXmlTextWriter.Close();
                                responseStringWriter.Close();
                                responseStream.Close();
                            }
                        }
                    }
                }
                else
                {
                    using (StreamReader streamReaderResponse = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        this.bodyData = streamReaderResponse.ReadToEnd();
                        stringBilder.AppendLine(this.bodyData);
                        streamReaderResponse.Close();
                    }

                    // If the content type is NOT XML data, the field "xmlData" will be set to null.
                    this.xmlData = null;
                }
            }
            else
            {
                // If there is no content in the HTTP response, the field "xmlData" will be set to null and the field "bodyData" will be set to empty.
                this.xmlData = null;
                this.bodyData = string.Empty;
            }

            // Reserve the HTTP headers and body data in a string and return it.
            lastRawResponse = stringBilder.ToString();

            // Output raw response in test log.
            site.Log.Add(LogEntryKind.Debug, "The raw response message is: \r\n", lastRawResponse);

            return lastRawResponse;
        }
    }
}