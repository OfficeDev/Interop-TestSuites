namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the GetSpecialTable request type.
    /// </summary>
    public class GetSpecialTableRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets a set of bit flags that specify options to the server.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether has state
        /// </summary>
        public bool HasState { get; set; }

        /// <summary>
        /// Gets or sets a STAT structure that specifies the state of a specific address book container.
        /// </summary>
        public STAT State { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Version field is present.
        /// </summary>
        public bool HasVersion { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the version number of the address book hierarchy table that the client has.
        /// </summary>
        public uint Version { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Flags));
            listByte.AddRange(BitConverter.GetBytes(this.HasState));
            if (this.HasState)
            {
                listByte.AddRange(this.State.Serialize());
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasVersion));
            if (this.HasVersion)
            {
                listByte.AddRange(BitConverter.GetBytes(this.Version));
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}