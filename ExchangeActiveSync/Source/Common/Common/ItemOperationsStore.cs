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