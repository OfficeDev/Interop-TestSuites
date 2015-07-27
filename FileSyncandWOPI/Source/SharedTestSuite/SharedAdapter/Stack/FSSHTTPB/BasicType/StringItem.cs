//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The count and content of an arbitrary wide character string.
    /// </summary>
    public class StringItem : BasicObject
    {
        /// <summary>
        /// Initializes a new instance of the StringItem class.
        /// </summary>
        public StringItem()
        {
            this.Count = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets the count of character.
        /// </summary>
        public Compact64bitInt Count { get; set; }

        /// <summary>
        /// Gets or sets an array of UTF-16 characters that specify the string. It MUST NOT be null-terminated.
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// This method is used to convert the element of StringItem basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of StringItem.</returns>
        public override List<byte> SerializeToByteList()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// This method is used to deserialize the StringItem basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the StringItem basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.Count = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.Content = System.Text.Encoding.Unicode.GetString(byteArray, index, Convert.ToInt32(this.Count.DecodedValue * 2));
            index += (int)this.Count.DecodedValue * 2;
            return index - startIndex;
        }
    }
}