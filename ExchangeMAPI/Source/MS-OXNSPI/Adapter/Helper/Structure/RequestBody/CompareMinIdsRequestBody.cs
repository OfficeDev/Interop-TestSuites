namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the CompareMIds request type.
    /// </summary>
    public class CompareMinIdsRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the reserved property. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
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
        /// Gets or sets an unsigned integer to the MinimalId1 that specifies the Minimal Entry ID of the first object. 
        /// </summary>
        public uint MinimalId1 { get; set; }

        /// <summary>
        ///  Gets or sets an unsigned integer to the MinimalId1 that specifies the Minimal Entry ID of the first object. 
        /// </summary>
        public uint MinimalId2 { get; set; }

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
            
            listByte.AddRange(BitConverter.GetBytes((uint)this.MinimalId1));
            listByte.AddRange(BitConverter.GetBytes((uint)this.MinimalId2));
            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}