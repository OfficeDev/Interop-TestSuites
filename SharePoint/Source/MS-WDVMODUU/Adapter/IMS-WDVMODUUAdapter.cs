namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System.Collections.Specialized;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This interface is used to generate MS-WDVMODUU protocol request.
    /// </summary>
    public interface IMS_WDVMODUUAdapter : IAdapter
    {
        /// <summary>
        /// Gets the last raw HTTP Response string (include the HTTP headers and XML data if exists) from the protocol server.
        /// </summary>
        string LastRawResponse { get; }

        /// <summary>
        /// Gets the last raw HTTP Request string (include the HTTP headers and body content if exists) that send to the protocol server.
        /// </summary>
        string LastRawRequest { get; }

        /// <summary>
        /// The method is used to send a HTTP request using PROPFIND method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="body">The body content in the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        WDVMODUUResponse PropFind(string requestUri, string body, NameValueCollection headersCollection);

        /// <summary>
        /// The method is used to send a HTTP request using PUT method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="body">The body content in the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        WDVMODUUResponse Put(string requestUri, byte[] body, NameValueCollection headersCollection);

        /// <summary>
        /// The method is used to send a HTTP request using GET method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        WDVMODUUResponse Get(string requestUri, NameValueCollection headersCollection);

        /// <summary>
        /// The method is used to send a HTTP request using DELETE method to the protocol server.
        /// As a result, it will return the "HttpWebResponse" object received from the protocol server.
        /// </summary>
        /// <param name="requestUri">The resource Request_URI for the HTTP request.</param>
        /// <param name="headersCollection">The collections for Name/Value pair of headers that would be inserted in the header of the HTTP request.</param>
        /// <returns>The "WDVMODUUResponse" Object that reserved data from the protocol server for the HTTP request.</returns>
        WDVMODUUResponse Delete(string requestUri, NameValueCollection headersCollection);
    }
}