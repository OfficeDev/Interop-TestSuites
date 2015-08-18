namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the GetPropList request type.
    /// </summary>
    public class GetPropListRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets a set of bit flags that specify options to the server.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets a MinimalEntryID structure that specifies the object for which to return properties.
        /// </summary>
        public uint MinmalId { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the code page that the server is being requested to use for string values of properties.
        /// </summary>
        public uint CodePage { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Flags));
            listByte.AddRange(BitConverter.GetBytes(this.MinmalId));
            listByte.AddRange(BitConverter.GetBytes(this.CodePage));
            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}