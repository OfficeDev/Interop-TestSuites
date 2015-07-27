//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;

    /// <summary>
    /// The AuxBlock
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class AuxBlock : SerializableBase
    { 
        /// <summary>
        /// An unsigned 32-bit integer value that specifies the size in bytes of 
        /// the BlockBytes field. The value of BlockBytesCount is zero (0x00000000) 
        /// if BlockBytes contains no data.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private uint blockBytesCount;

        /// <summary>
        /// A PtypBinary ([MS-OXCDATA] section 2.11.1) value. Semantics are determined
        /// by the value of the BlockType field. MUST be exactly BlockBytesCount bytes
        /// long.
        /// </summary>
        [SerializableFieldAttribute(3)]
        private byte[] blockBytes;

        /// <summary>
        /// Serialize fields to a stream
        /// </summary>
        /// <param name="stream">The stream where serialized instance will be wrote.</param>
        /// <returns>Bytes wrote to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            if (this.blockBytesCount != this.blockBytes.Length)
            {
                AdapterHelper.Site.Assert.Fail("The serialization operation should be successful.");
            }

            return base.Serialize(stream);
        }

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">stream contains a serialized instance of this class</param>
        /// <param name="size">The number of bytes can read if -1, no limitation, MUST be -1</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual actual is {0}.", size);

            int bytesRead = 0;
            bytesRead += 2;
            this.blockBytesCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.blockBytes = new byte[this.blockBytesCount];
            bytesRead += stream.Read(this.blockBytes, 0, (int)this.blockBytesCount);

            AdapterHelper.Site.Assert.Fail("The bytes length to read is not equal to the stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
        
            return bytesRead;
        }
    }
}