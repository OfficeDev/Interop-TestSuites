namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSCONT.
    /// </summary>
    public interface IMS_OXWSCONTAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Get contact item elements on the server.
        /// </summary>
        /// <param name="getItemRequest">The request of GetItem operation.</param>
        /// <returns>A response to GetItem operation request.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);

        /// <summary>
        /// Delete contact item elements on the server.
        /// </summary>
        /// <param name="deleteItemRequest">The request of DeleteItem operation.</param>
        /// <returns>A response to DeleteItem operation request.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);

        /// <summary>
        /// Create contact item elements on the server.
        /// </summary>
        /// <param name="createItemRequest">The request of CreateItem operation.</param>
        /// <returns>A response to CreateItem operation request.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Update contact item elements on the server.
        /// </summary>
        /// <param name="updateItemRequest">The request of UpdateItem operation.</param>
        /// <returns>A response to UpdateItem operation request.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest);

        /// <summary>
        /// Copy contact item elements on the server.
        /// </summary>
        /// <param name="copyItemRequest">The request of CopyItem operation.</param>
        /// <returns>A response to CopyItem operation request.</returns>
        CopyItemResponseType CopyItem(CopyItemType copyItemRequest);

        /// <summary>
        /// Move contact item elements on the server.
        /// </summary>
        /// <param name="moveItemRequest">The request of MoveItem operation.</param>
        /// <returns>A response to MoveItem operation request.</returns>
        MoveItemResponseType MoveItem(MoveItemType moveItemRequest);
    }
}