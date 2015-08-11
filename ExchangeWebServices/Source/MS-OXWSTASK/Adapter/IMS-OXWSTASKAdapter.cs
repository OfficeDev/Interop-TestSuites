namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSTASK.
    /// </summary>
    public interface IMS_OXWSTASKAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XMl request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Gets Task items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specifies a request to get Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);

        /// <summary>
        /// Copies Task items and puts the items in a different folder.
        /// </summary>
        /// <param name="copyItemRequest">Specifies a request to copy Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        CopyItemResponseType CopyItem(CopyItemType copyItemRequest);

        /// <summary>
        /// Creates Task items on the server.
        /// </summary>
        /// <param name="createItemRequest">Specifies a request to create Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Deletes Task items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specifies a request to delete Task item on the server.</param>
        /// <returns>A response to this operation request.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);

        /// <summary>
        /// Moves Task items on the server.
        /// </summary>
        /// <param name="moveItemRequest">Specifies a request to move Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        MoveItemResponseType MoveItem(MoveItemType moveItemRequest);

        /// <summary>
        /// Updates Task items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specifies a request to update Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest);
    }
}