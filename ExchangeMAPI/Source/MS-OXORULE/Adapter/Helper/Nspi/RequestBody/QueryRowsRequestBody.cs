namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the QueryRows request type.
    /// </summary>
    public class QueryRowsRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets a set of bit flags that specify the options the server.
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
        /// Gets or sets an unsigned integer that specifies the number of structures present in the ExplicitTable field. This value is limited to 100000.
        /// </summary>
        public uint ExplicitTableCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures that constitute the Explicit Table.
        /// </summary>
        public uint[] ExplicitTable { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of rows the client is requesting.
        /// </summary>
        public uint RowCount { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Columns fields is present. 
        /// </summary>
        public bool HasColumns { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the properties that the client requires for each row returned.
        /// </summary>
        public LargePropTagArray Columns { get; set; }

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

            listByte.AddRange(BitConverter.GetBytes(this.ExplicitTableCount));
            if (this.ExplicitTableCount != 0)
            {
                for (int i = 0; i < this.ExplicitTableCount; i++)
                {
                    listByte.AddRange(BitConverter.GetBytes(this.ExplicitTable[i]));
                }  
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