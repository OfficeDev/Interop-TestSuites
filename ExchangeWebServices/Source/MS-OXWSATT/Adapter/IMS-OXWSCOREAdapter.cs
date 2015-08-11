namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods CreateItem and DeleteItem defined in MS-OXWSCORE.
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
        /// <param name="createItemRequest">Request message of "CreateItem" operation.</param>
        /// <returns>Response message of "CreateItem" operation.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Deletes items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Request message of "DeleteItem" operation.</param>
        /// <returns>Response message of "DeleteItem" operation.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);
    }
}