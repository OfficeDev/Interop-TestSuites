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
    using System.Net;

    /// <summary>
    /// Wrapper class contains all the HTTP information returned from the server
    /// </summary>
    public class ActiveSyncRawResponse
    {
        /// <summary>
        /// Gets or sets the HTTP status code 
        /// </summary>
        public HttpStatusCode StatusCode { get; set; }

        /// <summary>
        /// Gets or sets the HTTP status description
        /// </summary>
        public string StatusDescription { get; set; }

        /// <summary>
        /// Gets the HTTP headers
        /// </summary>
        public WebHeaderCollection Headers { get; private set; }

        /// <summary>
        /// Gets or sets the binary HTTP body content
        /// </summary>
        public byte[] AsHttpRawBody { get; set; }

        /// <summary>
        /// Gets or sets the XML string after WBXML decoding the HTTP body content
        /// </summary>
        public string DecodedAsHttpBody { get; set; }

        /// <summary>
        /// Sets web header collection. 
        /// </summary>
        /// <param name="headers">The web header collection</param>
        public void SetWebHeader(WebHeaderCollection headers)
        {
            this.Headers = headers;
        }
    }
}