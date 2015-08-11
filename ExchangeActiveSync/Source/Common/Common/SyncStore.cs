namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Wrapper class of sync result data for a sync operation
    /// </summary>
    public class SyncStore
    {
        /// <summary>
        /// The collection of add responses
        /// </summary>
        private Collection<Response.SyncCollectionsCollectionResponsesAdd> addResponses = new Collection<Response.SyncCollectionsCollectionResponsesAdd>();

        /// <summary>
        /// The collection of change responses
        /// </summary>
        private Collection<Response.SyncCollectionsCollectionResponsesChange> changeResponses = new Collection<Response.SyncCollectionsCollectionResponsesChange>();

        /// <summary>
        /// The collection of add element
        /// </summary>
        private Collection<Sync> addElements = new Collection<Sync>();

        /// <summary>
        /// The collection of change element
        /// </summary>
        private Collection<Sync> changeElements = new Collection<Sync>();

        /// <summary>
        /// The collection of delete element
        /// </summary>
        private Collection<string> deleteElements = new Collection<string>();

        /// <summary>
        /// Gets or sets the Sync Key returned by the server, used to next sync operation
        /// </summary>
        public string SyncKey { get; set; }

        /// <summary>
        /// Gets or sets the Status in the Collection element
        /// </summary>
        public byte CollectionStatus { get; set; }

        /// <summary>
        /// Gets or sets the Status of the sync operation
        /// </summary>
        public byte Status { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Status of the sync operation exists
        /// </summary>
        public bool StatusSpecified { get; set; }

        /// <summary>
        /// Gets or sets the collection id
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets the add sync response information for this sync. 
        /// </summary>
        public Collection<Response.SyncCollectionsCollectionResponsesAdd> AddResponses
        {
            get { return this.addResponses; }
        }

        /// <summary>
        /// Gets the change sync response information for this sync. 
        /// </summary>
        public Collection<Response.SyncCollectionsCollectionResponsesChange> ChangeResponses
        {
            get { return this.changeResponses; }
        }

        /// <summary>
        /// Gets the sync item list of add elements
        /// </summary>
        public Collection<Sync> AddElements
        {
            get { return this.addElements; }
        }

        /// <summary>
        /// Gets the sync item list of change elements
        /// </summary>
        public Collection<Sync> ChangeElements
        {
            get { return this.changeElements; }
        }

        /// <summary>
        /// Gets the serverId list of delete elements
        /// </summary>
        public Collection<string> DeleteElements
        {
            get { return this.deleteElements; }
        }
    }
}