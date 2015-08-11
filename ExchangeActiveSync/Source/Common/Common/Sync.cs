namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    /// <summary>
    /// Wrapper class of Sync response
    /// </summary>
    public class Sync
    {
        /// <summary>
        /// Gets or sets ServerId
        /// </summary>
        public string ServerId { get; set; }

        /// <summary>
        /// Gets or sets calendar item
        /// </summary>
        public Calendar Calendar { get; set; }

        /// <summary>
        /// Gets or sets email
        /// </summary>
        public Email Email { get; set; }

        /// <summary>
        ///  Gets or sets note
        /// </summary>
        public Note Note { get; set; }

        /// <summary>
        ///  Gets or sets contact
        /// </summary>
        public Contact Contact { get; set; }

        /// <summary>
        ///  Gets or sets task
        /// </summary>
        public Task Task { get; set; }
    }
}