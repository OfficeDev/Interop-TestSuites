namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A base class of request body for Address Book Server Endpoint.
    /// </summary>
    public abstract class AddressBookRequestBodyBase : IRequestBody
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the size, in bytes, of the AuxiliaryBuffer field. 
        /// </summary>
        public uint AuxiliaryBufferSize { get; set; }

        /// <summary>
        /// Gets or sets an array of bytes that constitute the auxiliary payload data returned from the server. 
        /// </summary>
        public byte[] AuxiliaryBuffer { get; set; }

        /// <summary>
        /// Serialize the object to an array of bytes.
        /// </summary>
        /// <returns>The serialized data to be returned.</returns>
        public abstract byte[] Serialize();
    }
}