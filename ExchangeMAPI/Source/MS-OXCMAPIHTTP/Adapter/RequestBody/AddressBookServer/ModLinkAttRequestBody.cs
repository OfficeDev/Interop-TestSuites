namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the ModLinkAtt request type.
    /// </summary>
    public class ModLinkAttRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets a set of bit flags that specify options to the server.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets a PropertyTag structure that specifies the property to be modified.
        /// </summary>
        public PropertyTag PropertyTag { get; set; }

        /// <summary>
        /// Gets or sets a MinimalEntryID structure that specifies the Minimal Entry ID of the address book row to be modified.
        /// </summary>
        public uint MinimalId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the EntryIdCount and EntryIds fields are present.
        /// </summary>
        public bool HasEntryIds { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the count structures in the EntryIds field.
        /// </summary>
        public uint EntryIdCount { get; set; }

        /// <summary>
        /// Gets or sets the buffer of EntryIds
        /// </summary>
        public byte[][] EntryIDs { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Flags));
            listByte.AddRange(this.PropertyTag.Serialize());

            listByte.AddRange(BitConverter.GetBytes(this.MinimalId));

            listByte.AddRange(BitConverter.GetBytes(this.HasEntryIds));
            if (this.HasEntryIds)
            {
                listByte.AddRange(BitConverter.GetBytes(this.EntryIdCount));
                for (int i = 0; i < this.EntryIDs.Length; i++)
                {
                    listByte.AddRange(this.EntryIDs[i]);
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}