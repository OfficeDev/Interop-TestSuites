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
    using System.Text;

    /// <summary>
    /// A class indicates the GetMailboxUrl request type.
    /// </summary>
    public class GetMailboxUrlRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets the unsigned integer to the Flags field. Not used. The client MUST set this field to 0x00000000 and the server MUST ignore this field.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets a null-terminated Unicode string that specifies the distinguished name of the mailbox server for which to look up the URL.
        /// </summary>
        public string ServerDn { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Flags));
            StringBuilder serverDnStringBuilder = new StringBuilder(this.ServerDn);
            serverDnStringBuilder.Append("\0\0");
            listByte.AddRange(
                System.Text.Encoding.Unicode.GetBytes(serverDnStringBuilder.ToString()));
            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);
            return listByte.ToArray();
        }
    }
}
