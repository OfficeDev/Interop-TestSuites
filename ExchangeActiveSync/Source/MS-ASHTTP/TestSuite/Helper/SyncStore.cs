//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System.Collections.ObjectModel;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// Wrapper class of the result data from Sync command response.
    /// </summary>
    public class SyncStore
    {
        /// <summary>
        /// Define a collection used to store SyncItem.
        /// </summary>
        private Collection<SyncItem> addCommands;

        /// <summary>
        /// Define a collection used to store SyncCollectionsCollectionResponsesAdd.
        /// </summary>
        private Collection<Response.SyncCollectionsCollectionResponsesAdd> addResponses;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncStore" /> class.
        /// </summary>
        public SyncStore()
        {
            // Initialize the AddCommands collection.
            if (this.addCommands != null)
            {
                this.addCommands.Clear();
            }
            else
            {
                this.addCommands = new Collection<SyncItem>();
            }

            // Initialize the AddResponses collection.
            if (this.addResponses != null)
            {
                this.addResponses.Clear();
            }
            else
            {
                this.addResponses = new Collection<Response.SyncCollectionsCollectionResponsesAdd>();
            }
        }

        /// <summary>
        /// Gets or sets the SyncKey returned by the server, used in next Sync command.
        /// </summary>
        public string SyncKey { get; set; }

        /// <summary>
        /// Gets or sets the Status of the Sync command.
        /// </summary>
        public byte Status { get; set; }

        /// <summary>
        /// Gets or sets the CollectionId of the Sync command.
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets the add sync response information of the Sync command. 
        /// </summary>
        public Collection<Response.SyncCollectionsCollectionResponsesAdd> AddResponses
        {
            get { return this.addResponses; }
        }

        /// <summary>
        /// Gets the sync item list of add command.
        /// </summary>
        public Collection<SyncItem> AddCommands
        {
            get { return this.addCommands; }
        }
    }
}