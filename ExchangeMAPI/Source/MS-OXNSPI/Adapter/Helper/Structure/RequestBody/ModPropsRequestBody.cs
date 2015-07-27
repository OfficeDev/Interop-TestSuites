//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A class indicates the ModProps request type.
    /// </summary>
    public class ModPropsRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the reserved flag. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
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
        /// Gets or sets a value indicating whether the PropertyTagsToRemove field is present.
        /// </summary>
        public bool HasPropertyTagsToRemove { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the properties that the client is requesting to be removed.
        /// </summary>
        public LargePropTagArray PropertyTagsToRemove { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the PropertyValues field is present.
        /// </summary>
        public bool HasPropertyValues { get; set; }

        /// <summary>
        /// Gets or sets a buffer of AddressBookPropValueList structure that specifies the properties to modified.
        /// </summary>
        public AddressBookPropValueList PropertyVaules { get; set; }

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

            listByte.AddRange(BitConverter.GetBytes(this.HasPropertyTagsToRemove));
            if (this.HasPropertyTagsToRemove)
            {
                listByte.AddRange(BitConverter.GetBytes(this.PropertyTagsToRemove.PropertyTagCount));
                for (int i = 0; i < this.PropertyTagsToRemove.PropertyTagCount; i++)
                {
                    listByte.AddRange(this.PropertyTagsToRemove.PropertyTags[i].Serialize());
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.HasPropertyValues));
            if (this.HasPropertyValues)
            {
                listByte.AddRange(BitConverter.GetBytes(this.PropertyVaules.PropertyValueCount));
                for (int i = 0; i < this.PropertyVaules.PropertyValueCount; i++)
                {
                    listByte.AddRange(this.PropertyVaules.PropertyValues[i].Serialize());
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}
