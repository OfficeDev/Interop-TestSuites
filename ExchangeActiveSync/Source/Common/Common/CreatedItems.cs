//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// The items created by case.
    /// </summary>
    public class CreatedItems
    {
        /// <summary>
        /// Define a collection used to store the subject of created items.
        /// </summary>
        private Collection<string> itemSubject;

        /// <summary>
        /// Initializes a new instance of the CreatedItems class.
        /// </summary>
        public CreatedItems()
        {
            // Initialize the itemSubject collection.
            if (this.itemSubject != null)
            {
                this.itemSubject.Clear();
            }
            else
            {
                this.itemSubject = new Collection<string>();
            }
        }

        /// <summary>
        /// Gets or sets the CollectionId of the created item's parent folder.
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets the subjects of the created items.
        /// </summary>
        public Collection<string> ItemSubject
        {
            get { return this.itemSubject; }
        }
    }
}