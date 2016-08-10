namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the GetMatches request type.
    /// </summary>
    public class GetMatchesRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the reserved property. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint Reserved { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether has state.
        /// </summary>
        public bool HasState { get; set; }

        /// <summary>
        /// Gets or sets a STAT structure that specifies the state of a specific address book container.
        /// </summary>
        public STAT State { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether has Minimal IDs.
        /// </summary>
        public bool HasMinimalIds { get; set; }

        /// <summary>
        /// Gets or sets the count of the Minimal IDs.
        /// </summary>
        public uint MinimalIdCount { get; set; }

        /// <summary>
        /// Gets or sets an array of the Minimal IDs.
        /// </summary>
        public uint[] MinimalIds { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer to the Interface option Flags field. Reserved. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint InterfaceOptionFlags { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Filter field is present.
        /// </summary>
        public bool HasFilter { get; set; }

        /// <summary>
        /// Gets or sets a restriction that is to be applied to the rows in the address book container.
        /// </summary>
        public byte[] Filter { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the PropertyNameGUID and PropertyNameId fields are present.
        /// </summary>
        public bool HasPropertyName { get; set; }

        /// <summary>
        /// Gets or sets the GUID of the property to be opened.
        /// </summary>
        public Guid PropertyNameGuid { get; set; }

        /// <summary>
        /// Gets or sets a 4-byte value that specifies the ID of the property to be opened.
        /// </summary>
        public uint PropertyNameId { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of rows the client is requesting.
        /// </summary>
        public uint RowCount { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Columns field is present.
        /// </summary>
        public bool HasColumns { get; set; }

        /// <summary>
        /// Gets or sets a LargePropertyTagArray structure that specifies the columns that the client is requesting.
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

            listByte.AddRange(BitConverter.GetBytes(this.HasMinimalIds));
            if (this.HasMinimalIds)
            {
                listByte.AddRange(BitConverter.GetBytes(this.MinimalIdCount));
                for (int i = 0; i < this.MinimalIdCount; i++)
                {
                    listByte.AddRange(BitConverter.GetBytes(this.MinimalIds[i]));
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.InterfaceOptionFlags));
            listByte.AddRange(BitConverter.GetBytes(this.HasFilter));
            if (this.HasFilter)
            {
                listByte.AddRange(this.Filter);
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasPropertyName));
            if (this.HasPropertyName)
            {
                listByte.AddRange(this.PropertyNameGuid.ToByteArray());
                listByte.AddRange(BitConverter.GetBytes(this.PropertyNameId));
            }

            listByte.AddRange(BitConverter.GetBytes(this.RowCount));
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