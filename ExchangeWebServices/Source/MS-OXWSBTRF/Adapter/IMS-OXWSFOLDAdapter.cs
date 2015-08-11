namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides method CreateFolder defined in MS-OXWSFOLD.
    /// </summary>
    public interface IMS_OXWSFOLDAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Creates folders on the server.
        /// </summary>
        /// <param name="createFolderRequest">Specify the request for the CreateFolder operation.</param>
        /// <returns>The response to this operation request.</returns>
        CreateFolderResponseType CreateFolder(CreateFolderType createFolderRequest);
    }
}