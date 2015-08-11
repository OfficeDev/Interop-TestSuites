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