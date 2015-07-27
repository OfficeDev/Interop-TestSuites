//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Contain the information to be changed for calendar related item.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// Gets or sets the identifier of item to be updated.
        /// </summary>
        public BaseItemIdType ItemId { get; set; }

        /// <summary>
        /// Gets or sets the URIs of well-known element to be updated.
        /// </summary>
        public UnindexedFieldURIType FieldURI { get; set; }

        /// <summary>
        /// Gets or sets the item used to store the element to be updated.
        /// </summary>
        public ItemType Item { get; set; }
    }
}
