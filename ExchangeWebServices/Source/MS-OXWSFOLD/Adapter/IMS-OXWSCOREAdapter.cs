//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods CreateItem, DeleteItem, UpdateItem, GetItem and SwitchUser defined in MS-OXWSCORE.
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
        /// Creates item on the server.
        /// </summary>
        /// <param name="createItemRequest">Create item operation request type.</param>
        /// <returns>Create item operation response type.</returns>
        CreateItemResponseType CreateItem(CreateItemType createItemRequest);

        /// <summary>
        /// Delete item on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Delete item operation request type.</param>
        /// <returns>Delete item operation response type.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest);

        /// <summary>
        /// Update item on the server.
        /// </summary>
        /// <param name="updateItemRequest">Update item operation request type.</param>
        /// <returns>Update item operation response type.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest);

        /// <summary>
        /// Get item on the server.
        /// </summary>
        /// <param name="getItemRequest">Get item operation request type.</param>
        /// <returns>Get item operation response type.</returns>
        GetItemResponseType GetItem(GetItemType getItemRequest);

        /// <summary>
        /// Switch the current user to a new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}