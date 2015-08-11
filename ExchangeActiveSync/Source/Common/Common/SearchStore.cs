namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// The wrapper class for the search result returned from the server
    /// </summary>
    public class SearchStore
    {
        /// <summary>
        /// The list of return results from the server
        /// </summary>
        private Collection<Search> searchResults = new Collection<Search>();

        /// <summary>
        /// Gets or sets the status in Search element
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets the status in Store element
        /// </summary>
        public string StoreStatus { get; set; }

        /// <summary>
        /// Gets or sets the total number
        /// </summary>
        public int Total { get; set; }

        /// <summary>
        /// Gets or sets the range returned
        /// </summary>
        public string Range { get; set; }

        /// <summary>
        /// Gets the list result returned from the server
        /// </summary>
        public Collection<Search> Results
        {
            get { return this.searchResults; }
        }
    }
}