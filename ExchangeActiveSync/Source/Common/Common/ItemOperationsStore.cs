//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// The wrapper class for the fetched result returned from the server
    /// </summary>
    public class ItemOperationsStore
    {
        /// <summary>
        /// The list of the ItemOperations items
        /// </summary>
        private Collection<ItemOperations> items = new Collection<ItemOperations>();

        /// <summary>
        /// Gets or sets the status of this itemOperations operation
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets all the information fetch from the itemOperations
        /// </summary>
        public Collection<ItemOperations> Items
        {
            get { return this.items; }
        }
    }
}