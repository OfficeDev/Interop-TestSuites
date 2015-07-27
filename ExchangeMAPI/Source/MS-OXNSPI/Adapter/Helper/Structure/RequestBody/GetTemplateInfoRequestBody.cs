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
    /// A class indicates the GetTemplateInfo request type.
    /// </summary>
    public class GetTemplateInfoRequestBody : AddressBookRequestBodyBase
    {
        /// <summary>
        /// Gets or sets a set of bit flags that specify options to the server.
        /// </summary>
        public uint Flags { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the display type of the template for which information is requested.
        /// </summary>
        public uint DisplayType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the TemplateDN field is present.
        /// </summary>
        public bool HasTemplateDn { get; set; }

        /// <summary>
        /// Gets or sets a string that specifies the distinguished name of the template requested.
        /// </summary>
        public string TemplateDn { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the code page of template for which information is requested.
        /// </summary>
        public uint CodePage { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the language code identifier(LCID) of the template for which information is requested.
        /// </summary>
        public uint LocaleId { get; set; }

        /// <summary>
        /// Serialize the instance to byte array.
        /// </summary>
        /// <returns>An array byte of the request body.</returns>
        public override byte[] Serialize()
        {
            List<byte> listByte = new List<byte>();

            listByte.AddRange(BitConverter.GetBytes(this.Flags));
            listByte.AddRange(BitConverter.GetBytes(this.DisplayType));
            listByte.AddRange(BitConverter.GetBytes(this.HasTemplateDn));
            if (this.HasTemplateDn == true)
            {
                StringBuilder rgbTemplatDNStringBuilder = new StringBuilder(this.TemplateDn);
                rgbTemplatDNStringBuilder.Append("\0");
                listByte.AddRange(
                    System.Text.Encoding.ASCII.GetBytes(rgbTemplatDNStringBuilder.ToString()));
            }

            listByte.AddRange(BitConverter.GetBytes(this.CodePage));
            listByte.AddRange(BitConverter.GetBytes(this.LocaleId));

            listByte.AddRange(BitConverter.GetBytes(this.AuxiliaryBufferSize));
            listByte.AddRange(this.AuxiliaryBuffer);

            return listByte.ToArray();
        }
    }
}
