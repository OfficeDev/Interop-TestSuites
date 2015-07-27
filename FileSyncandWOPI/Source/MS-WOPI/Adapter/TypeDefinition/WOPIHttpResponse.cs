//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.IO;
    using System.Net;

    /// <summary>
    /// A wrapper class that represents the response of WOPI operation.
    /// </summary>
    public class WOPIHttpResponse : IDisposable
    {
        /// <summary>
        /// A MemoryStream instance which is used to store the body stream of raw response. 
        /// </summary>
        private MemoryStream responseBodyStream;

        /// <summary>
        /// A HttpWebResponse represents the raw Http response from the WOPI server.
        /// </summary>
        private HttpWebResponse rawHttpWebResponseInstance;

        /// <summary>
        /// Initializes a new instance of the <see cref="WOPIHttpResponse"/> class.
        /// </summary>
        /// <param name="rawResponseInstance">A parameter represents the raw Http response from the WOPI server.</param>
        public WOPIHttpResponse(HttpWebResponse rawResponseInstance)
        {
            if (null == rawResponseInstance)
            {
                throw new ArgumentNullException("rawResponseInstance");
            }

            // Read the body stream to local memory stream.
            this.rawHttpWebResponseInstance = rawResponseInstance;
            if (0 != this.rawHttpWebResponseInstance.ContentLength)
            {
                using (Stream rawResponseStream = this.rawHttpWebResponseInstance.GetResponseStream())
                {
                    this.responseBodyStream = new MemoryStream();
                    rawResponseStream.CopyTo(this.responseBodyStream);
                    this.responseBodyStream.Position = 0;
                }

                // Close connection and related resource for the raw http response.
                this.rawHttpWebResponseInstance.Close();
            }
        }

        #region properties

        /// <summary>
        /// Gets the HttpWebResponse instance which represents the raw Http response from the WOPI server. It is not allow to use this instance to get the body response stream, uses ResponseStream property instead.
        /// </summary>
        public HttpWebResponse RawHttpWebResponse
        {
            get
            {
                return this.rawHttpWebResponseInstance;
            }
        }

        /// <summary>
        /// Gets the length of the contents in the response body.
        /// </summary>
        public long ContentLength
        {
            get
            {
                return this.rawHttpWebResponseInstance.ContentLength;
            }
        }

        /// <summary>
        /// Gets the ContentType of the contents in the response body.
        /// </summary>
        public string ContentType
        {
            get
            {
                return this.rawHttpWebResponseInstance.ContentType;
            }
        }

        /// <summary>
        /// Gets the header
        /// </summary>
        public WebHeaderCollection Headers
        {
            get
            {
                return this.rawHttpWebResponseInstance.Headers;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the response contains header items. The 'true' means it has header items.
        /// </summary>
        public bool HasHeaders
        {
            get
            {
                return null != this.rawHttpWebResponseInstance.Headers && 0 != this.rawHttpWebResponseInstance.Headers.Count;
            }
        }

        /// <summary>
        /// Gets the Uri instance which represents the response Uri.
        /// </summary>
        public Uri ResponseUri
        {
            get
            {
                return this.rawHttpWebResponseInstance.ResponseUri;
            }
        }

        /// <summary>
        /// Gets the status code of the current response.
        /// </summary>
        public int StatusCode
        {
            get
            {
                return (int)this.rawHttpWebResponseInstance.StatusCode;
            }
        }

        #endregion 

        /// <summary>
        /// A method is used to get the header value by specified name.
        /// </summary>
        /// <param name="headerName">A parameter represents the header name which is used to find out the header item.</param>
        /// <returns>A return value represents the header value.</returns>
        public string GetHeaderValueByName(string headerName)
        {
            if (string.IsNullOrEmpty(headerName))
            {
                throw new ArgumentNullException("headerName");
            }

            if (!this.HasHeaders)
            {
                throw new ArgumentNullException("HasHeaders");
            }

            string headerValue = this.rawHttpWebResponseInstance.Headers[headerName];

            // If return null, means headers collection does not contain the header named with the specified header name.
            if (null == headerValue)
            {
                throw new InvalidOperationException(string.Format("The header[{0}] dose not exist in response.", headerName));
            }

            return headerValue;
        }

        /// <summary>
        /// A method is used to get the stream copy of response body, so that the stream copy can be used freely and disposed as need.
        /// </summary>
        /// <returns>A return value represents the stream copy of response body. If there are no any contents in the response body, this method will return null.</returns>
        public Stream GetResponseStream()
        {
            if (this.responseBodyStream != null)
            {
                MemoryStream responseStreamCopy = new MemoryStream();
                this.responseBodyStream.CopyTo(responseStreamCopy);
                this.responseBodyStream.Position = 0;
                responseStreamCopy.Position = 0;
                return responseStreamCopy;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.responseBodyStream.Close();
            GC.SuppressFinalize(this);
        }
    }
}
