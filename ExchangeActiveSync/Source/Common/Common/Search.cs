namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    /// <summary>
    /// Wrapper class of Search response
    /// </summary>
    public class Search
    {
        /// <summary>
        /// Gets or sets the long id of the calendar returned by the search operation
        /// </summary>
        public string LongId { get; set; }

        /// <summary>
        ///  Gets or sets the class name of the calendar
        /// </summary>
        public string Class { get; set; }

        /// <summary>
        /// Gets or sets the collection id
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets or sets the calendar information
        /// </summary>
        public Calendar Calendar { get; set; }

        /// <summary>
        /// Gets or sets the note item
        /// </summary>
        public Note Note { get; set; }

        /// <summary>
        /// Gets or sets the contact item
        /// </summary>
        public Contact Contact { get; set; }

        /// <summary>
        /// Gets or sets the email item
        /// </summary>
        public Email Email { get; set; }

        /// <summary>
        /// Gets or sets the task item
        /// </summary>
        public Task Task { get; set; }
    }
}