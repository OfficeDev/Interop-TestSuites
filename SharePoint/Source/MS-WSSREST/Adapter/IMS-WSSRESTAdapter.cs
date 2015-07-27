//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-WSSREST Adapter class.
    /// </summary>
    public interface IMS_WSSRESTAdapter : IAdapter
    {
        /// <summary>
        /// Insert a list item.
        /// </summary>
        /// <param name="request">The content of the list item that be inserted.</param>
        /// <returns>The list item that be inserted.</returns>
        Entry InsertListItem(Request request);

        /// <summary>
        /// Update a list item.
        /// </summary>
        /// <param name="request">The content of the list item that be updated.</param>
        /// <returns>The ETag of this list item.</returns>
        string UpdateListItem(Request request);

        /// <summary>
        /// Retrieve list item from server.
        /// </summary>
        /// <param name="request">The retrieve condition.</param>
        /// <returns>The response from server.</returns>
        object RetrieveListItem(Request request);

        /// <summary>
        /// Delete a special list item.
        /// </summary>
        /// <param name="request">The special list item.</param>
        /// <returns>True if the list item be deleted success, otherwise false.</returns>
        bool DeleteListItem(Request request);

        /// <summary>
        /// Package many requests(insert,update or delete request) in one batch request.
        /// </summary>
        /// <param name="requests">The multi requests.</param>
        /// <returns>The response from server.</returns>
        string BatchRequests(List<BatchRequest> requests);
    }
}
