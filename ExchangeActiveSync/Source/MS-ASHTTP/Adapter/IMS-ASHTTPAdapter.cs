namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASHTTP.
    /// </summary>
    public interface IMS_ASHTTPAdapter : IAdapter
    {
        /// <summary>
        /// Gets the XML request sent to protocol SUT
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the XML response received from protocol SUT
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Send HTTP POST request to the server and get the response.
        /// </summary>
        /// <param name="commandName">The name of the command to send.</param>
        /// <param name="commandParameters">The command parameters.</param>
        /// <param name="requestBody">The plain text request.</param>
        /// <returns>The plain text response.</returns>
        SendStringResponse HTTPPOST(CommandName commandName, IDictionary<CmdParameterName, object> commandParameters, string requestBody);

        /// <summary>
        /// Send HTTP OPTIONS request to the server and get the response.
        /// </summary>
        /// <returns>The HTTP OPTIONS response.</returns>
        OptionsResponse HTTPOPTIONS();

        /// <summary>
        /// Configure the fields in request line or request headers besides command name and command parameters.
        /// </summary>
        /// <param name="requestPrefixFields">The fields in request line or request headers which need to be configured besides command name and command parameters.</param>
        void ConfigureRequestPrefixFields(IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields);
    }
}