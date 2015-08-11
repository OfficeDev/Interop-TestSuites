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