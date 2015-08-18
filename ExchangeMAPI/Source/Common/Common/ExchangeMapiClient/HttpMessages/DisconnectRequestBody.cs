namespace Microsoft.Protocols.TestSuites.Common
{ 
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A disconnect request body for Mailbox Server Endpoint.
    /// </summary>
    public class DisconnectRequestBody : MailboxRequestBodyBase
    {
        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>The serialized data to be returned.</returns>
        public override byte[] Serialize()
        {
            List<byte> rawData = new List<byte>();

            rawData.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            rawData.AddRange(this.AuxiliaryBuffer);

            return rawData.ToArray();
        }
    }
}