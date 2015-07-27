//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Wrapper conversion item class for the conversation id and the server id
    /// </summary>
    public class ConversationItem
    {
        /// <summary>
        /// The server id collection in the current conversion.
        /// </summary>
        private Collection<string> serverId;

        /// <summary>
        /// Initializes a new instance of the ConversationItem class.
        /// </summary>
        public ConversationItem()
        {
            if (this.serverId != null)
            {
                this.serverId.Clear();
            }
            else
            {
                this.serverId = new Collection<string>();
            }
        }

        /// <summary>
        /// Gets or sets the conversation id
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets the server id collection
        /// </summary>
        public Collection<string> ServerId 
        {
            get { return this.serverId; }
        }
    }
}