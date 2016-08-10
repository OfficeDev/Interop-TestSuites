namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the ResolveNames request type.
    /// </summary>
    public class ResolveNamesRequestBody : AddressBookRequestBodyBase
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
        /// Gets or sets a value indicating whether the Property Tags field is present.
        /// </summary>
        public bool HasPropertyTags { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the properties that client requires for the row returned.
        /// </summary>
        public LargePropertyTagArray PropertyTags { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Names field is present.
        /// </summary>
        public bool HasNames { get; set; }

        /// <summary>
        /// Gets or sets a WStringArray_r structure that specifies the values on which the client is requesting that the server perform ambiguous name resolution.
        /// </summary>
        public WStringsArray_r Names { get; set; }

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

            listByte.AddRange(BitConverter.GetBytes(this.HasPropertyTags));
            if (this.HasPropertyTags)
            {
                listByte.AddRange(BitConverter.GetBytes(this.PropertyTags.PropertyTagCount));
                for (int i = 0; i < this.PropertyTags.PropertyTagCount; i++)
                {
                    listByte.AddRange(this.PropertyTags.PropertyTags[i].Serialize());
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasNames));

            if (this.HasNames)
            {
                listByte.AddRange(BitConverter.GetBytes(this.Names.CValues));
                for (int i = 0; i < this.Names.CValues; i++)
                {
                    StringBuilder nameStringBuilder = new StringBuilder(this.Names.LppszW[i]);
                    nameStringBuilder.Append("\0\0");
                    listByte.AddRange(
                        System.Text.Encoding.Unicode.GetBytes(nameStringBuilder.ToString()));
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}