namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSCORE.
    /// </summary>
    public interface IMS_OXWSCOREAdapter : IAdapter
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
        /// Copy items and puts the items in a different folder.
        /// </summary>
        /// <param name="copyItemRequest">Specify a request to copy items on the server.</param>
        /// <returns>A response to CopyItem operation request.</returns>
        CopyItemResponseType CopyItem(CopyItemType copyItemRequest);

        /// <summary>
        /// Create items in the Exchange store
        /// </summary>
        /// <param name="createItemRequest">Specify a request to create items on the server.</param>
        /// <returns>A response to CreateItem operation request.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specify a request to delete item on the server.</param>
        /// <returns>A response to DeleteItem operation request.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);

        /// <summary>
        /// Get items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify a request to get items on the server.</param>
        /// <returns>A response to GetItem operation request.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);

        /// <summary>
        /// Move items on the server.
        /// </summary>
        /// <param name="moveItemRequest">Specify a request to move items on the server.</param>
        /// <returns>A response to MoveItem operation request.</returns>
        MoveItemResponseType MoveItem(MoveItemType moveItemRequest);

        /// <summary>
        /// Send messages and post items on the server.
        /// </summary>
        /// <param name="sendItemRequest">Specify a request to send items on the server.</param>
        /// <returns>A response to SendItem operation request.</returns>
        SendItemResponseType SendItem(SendItemType sendItemRequest);

        /// <summary>
        /// Update items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specify a request to update items on the server.</param>
        /// <returns>A response to UpdateItem operation request.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest);

        /// <summary>
        /// Mark all items in a folder as read.
        /// </summary>
        /// <param name="markAllItemAsReadRequest">Specify a request to mark all items as read.</param>
        /// <returns>A response to MarkAllItemsAsRead operation request.</returns>
        MarkAllItemsAsReadResponseType MarkAllItemsAsRead(MarkAllItemsAsReadType markAllItemAsReadRequest);

        /// <summary>
        /// The MarkAsJunk operation marks an item as junk.
        /// </summary>
        /// <param name="markAsJunkRequest">Specify a request for a MarkAsJunk operation.</param>
        /// <returns>A response to MarkAsJunk operation request.</returns>
        MarkAsJunkResponseType MarkAsJunk(MarkAsJunkType markAsJunkRequest);

        /// <summary>
        /// Switch the current user to a new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        void ConfigureSOAPHeader(Dictionary<string, object> headerValues);
    }
}