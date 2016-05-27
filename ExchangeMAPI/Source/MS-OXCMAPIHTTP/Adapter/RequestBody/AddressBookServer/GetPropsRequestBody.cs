namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the GetProps request type.
    /// </summary>
    public class GetPropsRequestBody : AddressBookRequestBodyBase
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
        /// Gets or sets a value indicating whether the PropertyTags field is present.
        /// </summary>
        public bool HasPropertyTags { get; set; }

        /// <summary>
        /// Gets or sets a LargePropertyTagArray structure that contains the property tags of the properties that the client requesting.
        /// </summary>
        public LargePropertyTagArray PropertyTags { get; set; }

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

            listByte.AddRange(BitConverter.GetBytes(this.HasPropertyTags));
            if (this.HasPropertyTags)
            {
                listByte.AddRange(BitConverter.GetBytes(this.PropertyTags.PropertyTagCount));
                for (int i = 0; i < this.PropertyTags.PropertyTagCount; i++)
                {
                    listByte.AddRange(this.PropertyTags.PropertyTags[i].Serialize());
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}