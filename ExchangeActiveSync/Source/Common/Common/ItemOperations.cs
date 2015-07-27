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
    /// <summary>
    /// Wrapper class of ItemOperations response
    /// </summary>
    public class ItemOperations
    {
        /// <summary>
        /// Gets or sets the Class
        /// </summary>
        public string Class { get; set; }

        /// <summary>
        /// Gets or sets the ServerId
        /// </summary>
        public string ServerId { get; set; }

        /// <summary>
        /// Gets or sets the Status
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets CollectionId
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets or sets the Calendar
        /// </summary>
        public Calendar Calendar { get; set; }

        /// <summary>
        /// Gets or sets Email
        /// </summary>
        public Email Email { get; set; }

        /// <summary>
        /// Gets or sets Note
        /// </summary>
        public Note Note { get; set; }

        /// <summary>
        /// Gets or sets Contact
        /// </summary>
        public Contact Contact { get; set; }

        /// <summary>
        /// Gets or sets task
        /// </summary>
        public Task Task { get; set; }
    }
}