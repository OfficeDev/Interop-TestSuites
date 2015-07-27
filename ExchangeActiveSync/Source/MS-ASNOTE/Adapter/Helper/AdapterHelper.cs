//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// The class provides the methods to assist MS_ASNOTEAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        #region Adapter Help Methods
        /// <summary>
        /// Check if the request of ItemOperations command contains the schema element
        /// </summary>
        /// <param name="itemOperationsRequest">ItemOperations command request.</param>
        /// <returns>The boolean value represents whether the request of ItemOperations command contains the schema element</returns>
        public static bool ContainsSchemaElement(ActiveSyncRequestBase<Request.ItemOperations> itemOperationsRequest)
        {
            Request.ItemOperationsFetch fetch = (Request.ItemOperationsFetch)itemOperationsRequest.RequestData.Items[0];

            bool hasSchemaElement = false;

            // Check if the request contains schema
            foreach (object item in fetch.Options.Items)
            {
                if (item.GetType().Equals(typeof(Request.Schema)))
                {
                    hasSchemaElement = true;
                }
            }

            return hasSchemaElement;
        }
        #endregion
    }
}