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