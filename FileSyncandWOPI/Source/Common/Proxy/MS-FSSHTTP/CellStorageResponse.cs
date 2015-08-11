namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// A class encapsulates the ResponseVersion and Response collection object.
    /// </summary>
    public class CellStorageResponse
    {
        /// <summary>
        /// Gets or sets version information of the response message.
        /// </summary>
        public ResponseVersion ResponseVersion { get; set; }

        /// <summary>
        /// Gets or sets a collection of sub responses.
        /// </summary>
        public ResponseCollection ResponseCollection { get; set; }
    }
}