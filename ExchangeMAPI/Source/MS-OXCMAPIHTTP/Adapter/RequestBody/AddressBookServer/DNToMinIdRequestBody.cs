//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the DNToMId request type.
    /// </summary>
    public class DNToMinIdRequestBody : AddressBookRequestBodyBase 
    {
        /// <summary>
        /// Gets or sets the reserved property. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint Reserved { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether has names
        /// </summary>
        public bool HasNames { get; set; }

        /// <summary>
        /// Gets or sets an array of string to the names field that contains the list of distinguished names to be mapped to Minimal Entry IDs.
        /// </summary>
        public StringArray_r Names { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Reserved));
            listByte.AddRange(BitConverter.GetBytes(this.HasNames));
            if (this.HasNames)
            {
                listByte.AddRange(BitConverter.GetBytes(this.Names.CValues));
                for (int i = 0; i < this.Names.CValues; i++)
                {
                     StringBuilder name = new StringBuilder(this.Names.LppzA[i]);
                     name.Append("\0");
                     listByte.AddRange(System.Text.Encoding.ASCII.GetBytes(name.ToString()));
                }
            }

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}
