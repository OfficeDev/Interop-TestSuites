//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        /// <summary>
        /// Get message related Item elements of ItemInfoResponseMessageType type in a response message returned from server.
        /// </summary>
        /// <param name="response">A response message returned by an operation: GetItem, CopyItem, CreateItem, MoveItem, or UpdateItem.</param>
        /// <returns>Message related Item elements of ItemInfoResponseMessageType type.</returns>
        public static ItemInfoResponseMessageType[] GetInfoItemsInResponse(BaseResponseMessageType response)
        {
            ItemInfoResponseMessageType[] infoItems = null;
            if (response != null
                && response.ResponseMessages != null
                && response.ResponseMessages.Items != null
                && response.ResponseMessages.Items.Length > 0)
            {
                List<ItemInfoResponseMessageType> infoItemList = new List<ItemInfoResponseMessageType>();
                foreach (ResponseMessageType item in response.ResponseMessages.Items)
                {
                    infoItemList.Add(item as ItemInfoResponseMessageType);
                }

                if (infoItemList.Count > 0)
                {
                    infoItems = infoItemList.ToArray();
                }
            }

            return infoItems;
        }

        /// <summary>
        /// Get the responseMessageItem of type ItemType in the message array returned in the server response by the specified index.
        /// </summary>
        /// <param name="infoItems">The related Item elements of ItemInfoResponseMessageType type in a response message returned from server.</param>
        /// <param name="indexOfInfoItem">The index of the responseMessageItem in infoItems to retrieve.</param>
        /// <param name="indexOfItemTypeItem">The index of the responseMessageItem of ItemType to retrieve.</param>
        /// <returns>The first responseMessageItem in the ItemType array.</returns>
        public static ItemType GetItemTypeItemFromInfoItemsByIndex(ItemInfoResponseMessageType[] infoItems, ushort indexOfInfoItem, ushort indexOfItemTypeItem)
        {
            ItemType item = null;

            if (infoItems != null && infoItems.Length > 0 && infoItems.Length > indexOfInfoItem)
            {
                if (infoItems[indexOfInfoItem].Items != null
                    && infoItems[indexOfInfoItem].Items.Items != null
                    && infoItems[indexOfInfoItem].Items.Items.Length > 0
                    && infoItems[indexOfInfoItem].Items.Items.Length > indexOfItemTypeItem)
                {
                    // The messageResponse is used to save the message returned from server.
                    item = infoItems[indexOfInfoItem].Items.Items[indexOfItemTypeItem];
                }
            }

            return item;
        }
    }
}