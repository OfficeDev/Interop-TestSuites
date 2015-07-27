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
    /// The SizedXid.
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class SizedXid : SerializableBase
    {
        /// <summary>
        ///  An unsigned 8-bit integer. MUST be equal to the size of the XID field 
        ///  in bytes.
        /// </summary>
        private byte xidSize;

        /// <summary>
        ///  A structure of type XID that contains the value of the internal identifier 
        ///  of an object, or internal or external identifier of a change number. 
        /// </summary>
        private XID xid;

        /// <summary>
        /// Gets or sets the xidSize.
        /// </summary>
        public byte XidSize
        {
            get
            {
                return this.xidSize;
            }

            set
            {
                this.xidSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the XID.
        /// </summary>
        public XID XID
        {
            get
            {
                return this.xid;
            }

            set
            {
                this.xid = value;
            }
        }

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">Stream contains a serialized instance of this class.</param>
        /// <param name="size">How many bytes can read if -1, no limitation. MUST be -1.</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual value is {0}.", size);

            int bytesRead = 0;
            this.xidSize = StreamHelper.ReadUInt8(stream);
            bytesRead += 1;
            this.xid = new XID();
            bytesRead += this.xid.Deserialize(stream, (int)this.xidSize);
            return bytesRead;
        }

        /// <summary>
        /// Serialize fields to a stream.
        /// </summary>
        /// <param name="stream">The stream where serialized instance will be wrote.</param>
        /// <returns>The number of bytes wrote to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            int bytesWriten = 0;
            stream.WriteByte(this.xidSize);
            bytesWriten += 1;
            bytesWriten += this.xid.Serialize(stream);
            return bytesWriten;
        }
    }
}