namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    /// <summary>
    /// A base class of response body for Address Book Server Endpoint.
    /// </summary>
    public abstract class AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the status of the request.
        /// </summary>
        public uint StatusCode { get; protected set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field.
        /// </summary>
        public uint AuxiliaryBufferSize { get; protected set; }

        /// <summary>
        /// Gets or sets an array of bytes that constitute the auxiliary payload data returned from the server. 
        /// </summary>
        public byte[] AuxiliaryBuffer { get; protected set; }
    }
}