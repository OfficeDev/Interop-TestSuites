namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods CreateItem, DeleteItem, GetItem and UpdateItem defined in MS-OXWSCORE.
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
        /// Create items on the server.
        /// </summary>
        /// <param name="createItemRequest">Specify a request to create items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specify a request to delete item on the server.</param>
        /// <returns>A response to this operation request.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);

        /// <summary>
        /// Get items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify a request to get items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);

        /// <summary>
        /// Update items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specify a request to update items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest);
    }
}