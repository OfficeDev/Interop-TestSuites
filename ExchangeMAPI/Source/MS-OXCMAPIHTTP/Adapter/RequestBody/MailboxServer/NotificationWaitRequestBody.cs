namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the NotificationWait request type.
    /// </summary>
    public class NotificationWaitRequestBody : MailboxRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the reserved flag. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the NotificationWait request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> rawData = new List<byte>();

            rawData.AddRange(BitConverter.GetBytes(this.Flags));
            rawData.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            rawData.AddRange(this.AuxiliaryBuffer);

            return rawData.ToArray();
        }
    }
}