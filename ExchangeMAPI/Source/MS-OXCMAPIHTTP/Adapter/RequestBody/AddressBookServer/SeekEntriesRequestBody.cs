namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the SeekEntries request type.
    /// </summary>
    public class SeekEntriesRequestBody : AddressBookRequestBodyBase
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
        /// Gets or sets a value indicating whether the Target field is present.
        /// </summary>
        public bool HasTarget { get; set; }

        /// <summary>
        /// Gets or sets a PropertyValue_r structure that specifies the property value being sought.
        /// </summary>
        public PropertyValue_r Target { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the ExplicitTableCount and ExplicitTable field are present.
        /// </summary>
        public bool HasExplicitTable { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures present in the ExplicitTable field.
        /// </summary>
        public uint ExplicitableCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures that constitute an Explicit Table.
        /// </summary>
        public uint[] ExplicitTable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the columns field is present.
        /// </summary>
        public bool HasColumns { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the columns that the client is requesting.
        /// </summary>
        public LargePropertyTagArray Columns { get; set; }

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

            listByte.AddRange(BitConverter.GetBytes(this.HasTarget));
            if (this.HasTarget)
            {
                listByte.AddRange(this.Target.PropTag.Serialize());

                // According to the product behavior, string or binary buffer start with "FF".
                listByte.Add(0xFF);
                listByte.AddRange(this.Target.Value);
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasExplicitTable));
            if (this.HasExplicitTable)
            {
                listByte.AddRange(BitConverter.GetBytes(this.ExplicitableCount));
                for (int i = 0; i < this.ExplicitableCount; i++)
                {
                    listByte.AddRange(BitConverter.GetBytes(this.ExplicitTable[i]));
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasColumns));
            if (this.HasColumns)
            {
                listByte.AddRange(BitConverter.GetBytes(this.Columns.PropertyTagCount));
                for (int i = 0; i < this.Columns.PropertyTagCount; i++)
                {
                    listByte.AddRange(this.Columns.PropertyTags[i].Serialize());
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}