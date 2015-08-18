namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the ResortRestriction request type.
    /// </summary>
    public class ResortRestrictionRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the reserved flag field. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint Reserved { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether has state
        /// </summary>
        public bool HasState { get; set; }

        /// <summary>
        /// Gets or sets a STAT structure that specifies the state of a specific address book container.
        /// </summary>
        public STAT State { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the MinimalIdCount and MinimalIds fields are present.
        /// </summary>
        public bool HasMinimalIds { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures in the MinimalIds field.
        /// </summary>
        public uint MinimalIdCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures that compose a restricted address book container.
        /// </summary>
        public uint[] MinimalIds { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Reserved));
            listByte.AddRange(BitConverter.GetBytes(this.HasState));
            if (this.HasState)
            {
                listByte.AddRange(this.State.Serialize());
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasMinimalIds));
            if (this.HasMinimalIds)
            {
                listByte.AddRange(BitConverter.GetBytes(this.MinimalIdCount));
                for (int i = 0; i < this.MinimalIdCount; i++)
                {
                   listByte.AddRange(BitConverter.GetBytes(this.MinimalIds[i])); 
                }  
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}