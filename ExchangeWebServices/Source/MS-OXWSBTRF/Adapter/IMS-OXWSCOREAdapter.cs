namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods CreateItem and GetItem defined in MS-OXWSCORE.
    /// </summary>
    public interface IMS_OXWSCOREAdapter : IAdapter
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
        /// Creates items on the server.
        /// </summary>
        /// <param name="createItemRequest">Specify the request for CreateItem operation.</param>
        /// <returns>The response to this operation request.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Gets items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify the request for GetItem operation.</param>
        /// <returns>The response to this operation result.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);
    }
}