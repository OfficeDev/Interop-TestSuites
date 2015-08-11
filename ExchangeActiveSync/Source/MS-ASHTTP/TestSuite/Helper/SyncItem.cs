namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    /// <summary>
    /// Wrapper class of ApplicationData from Sync command response.
    /// </summary>
    public class SyncItem
    {
        /// <summary>
        /// Gets or sets the ServerId.
        /// </summary>
        public string ServerId { get; set; }

        /// <summary>
        /// Gets or sets the Subject of the item.
        /// </summary>
        public string Subject { get; set; }
    }
}