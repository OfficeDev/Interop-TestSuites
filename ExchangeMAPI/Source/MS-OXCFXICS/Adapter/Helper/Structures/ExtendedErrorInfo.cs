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
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Contains extended and contextual information about an error 
    /// that has occurred when producing a FastTransfer stream.
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class ExtendedErrorInfo : SerializableBase
    {
        /// <summary>
        /// An unsigned 16-bit integer that determines the format of the structure. 
        /// The format shown in the preceding packet diagram corresponds 
        /// to version "0x00000000", which is the only version defined for 
        /// the protocol.
        /// </summary>
        private ushort version;

        /// <summary>
        /// SHOULD be set to zeros and MUST be ignored by the clients.
        /// </summary>
        private ushort padding1;

        /// <summary>
        /// One of the error codes defined in [MS-OXCDATA] that describes 
        /// the reason for the failure.
        /// </summary>
        private uint errorCode;

        /// <summary>
        ///  A GID structure that identifies the folder that was in context 
        ///  at the time the error occurred.
        ///  MUST be filled with zeros, if no folders were in context.
        /// </summary>
        private GID folderGID;

        /// <summary>
        /// SHOULD be set to zeros and MUST be ignored by the clients.
        /// </summary>
        private ushort padding2;

        /// <summary>
        /// A GID structure that identifies the message that was in context 
        /// at the time the error occurred. MUST be filled with zeros, 
        /// if no messages were in context.
        /// </summary>
        private GID messageGID;

        /// <summary>
        /// SHOULD be set to zeros and MUST be ignored by the clients.
        /// </summary>
        private ushort padding3;

        /// <summary>
        /// SHOULD be set to zeros and SHOULD be ignored by clients.
        /// </summary>
        private byte[] reserved1;

        /// <summary>
        /// An unsigned 32-bit integer value that specifies 
        /// the size of the AuxBytes field. If set to 0, AuxBytes is missing.
        /// </summary>
        private uint auxBytesCount;

        /// <summary>
        /// An unsigned 32-bit integer value that specifies the offset 
        /// in bytes of Auxbytes from the beginning of the structure.
        /// </summary>
        private uint auxBytesOffset;

        /// <summary>
        /// Optional. SHOULD be set to zeros and SHOULD be ignored by clients.
        /// </summary>
        private byte[] reserved2;

        /// <summary>
        /// An optional PtypBinary ([MS-OXCDATA] section 2.11.1) value that MUST 
        /// be present and reside at AuxBytesOffset from the beginning of the structure, 
        /// IFF AuxBytesCount > 0. If present, MUST consist of one or more AuxBlock 
        /// structures serialized sequentially without any padding.
        /// </summary>
        private byte[] auxBytes;

        /// <summary>
        /// Serialize this instance to a stream
        /// </summary>
        /// <param name="stream">A data stream contains serialized object.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            int previousPosition = (int)stream.Position;
            int bytesWritten = 0;
            bytesWritten += StreamHelper.WriteUInt16(stream, this.version);
            bytesWritten += StreamHelper.WriteUInt16(stream, this.padding1);
            bytesWritten += StreamHelper.WriteUInt32(stream, this.errorCode);
            byte[] buffer = StructureSerializer.Serialize(this.folderGID);
            bytesWritten += StreamHelper.WriteBuffer(stream, buffer);
            bytesWritten += StreamHelper.WriteUInt16(stream, this.padding2);
            buffer = StructureSerializer.Serialize(this.messageGID);
            bytesWritten += StreamHelper.WriteBuffer(stream, buffer);
            bytesWritten += StreamHelper.WriteUInt16(stream, this.padding3);

            if (this.reserved1 == null || this.reserved1.Length != 24)
            {
                AdapterHelper.Site.Assert.Fail("The Reserved field should not be null and its length MUST be 24, but the actual length is {0}.", this.reserved1.Length);
            }

            for (int i = 0; i < this.reserved1.Length; i++)
            {
                AdapterHelper.Site.Assert.AreEqual(0, this.reserved1[i], "Reserved (24 bytes):  SHOULD be set to zeros and SHOULD be ignored by clients.");
            }

            bytesWritten += StreamHelper.WriteBuffer(stream, this.reserved1);
            bytesWritten += StreamHelper.WriteUInt32(stream, this.auxBytesCount);
            bytesWritten += StreamHelper.WriteUInt32(stream, this.auxBytesOffset);
            if (this.reserved2 != null)
            {
                bytesWritten += StreamHelper.WriteBuffer(stream, this.reserved2);
            }

            AdapterHelper.Site.Assert.AreEqual((int)this.auxBytesOffset, previousPosition + bytesWritten, "The offset and writen length are not equal, the offset is {0} and writen length is {1}.", this.auxBytesOffset, previousPosition + bytesWritten);

            AdapterHelper.Site.Assert.AreEqual((int)this.auxBytesCount, this.auxBytes.Length, "The original and serialized auxBytes length are not equal, the original length is {0} and serialized length is {1}.", this.auxBytesCount, this.auxBytes.Length);

            bytesWritten += StreamHelper.WriteBuffer(stream, this.auxBytes);
            return bytesWritten;
        }

        /// <summary>
        /// Get AuxBlock list.
        /// </summary>
        /// <returns>A list contains AuxBlocks.</returns>
        public List<AuxBlock> GetAuxBlockList()
        {
            List<AuxBlock> auxBlocks = null;
            if (this.auxBytes != null)
            {
                auxBlocks = new List<AuxBlock>();
                using (MemoryStream stream = new MemoryStream(this.auxBytes, false))
                {
                    stream.Position = 0;
                    while (stream.Position < stream.Length)
                    {
                        AuxBlock block = new AuxBlock();
                        block.Deserialize(stream, -1);
                        auxBlocks.Add(block);
                    }
                }
            }

            return auxBlocks;
        }

        /// <summary>
        /// Deserialize from a stream.
        /// </summary>
        /// <param name="stream">A stream contains serialize.</param>
        /// <param name="size">Must be -1.</param>
        /// <returns>The number of bytes read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            int previousPosition = (int)stream.Position;
            int bytesRead = 0;
            this.version = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            this.padding1 = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            this.errorCode = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            byte[] buffer = new byte[22];
            stream.Read(buffer, 0, 22);
            this.folderGID = StructureSerializer.Deserialize<GID>(buffer);
            bytesRead += 22;
            this.padding2 = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            buffer = new byte[22];
            stream.Read(buffer, 0, 22);
            this.messageGID = StructureSerializer.Deserialize<GID>(buffer);
            bytesRead += 22;
            this.padding3 = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            this.reserved1 = new byte[24];
            stream.Read(this.reserved1, 0, 24);
            bytesRead += 24;
            this.auxBytesCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            this.auxBytesOffset = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            int reservedBytesCount = (int)this.auxBytesOffset - previousPosition - bytesRead;
            if (reservedBytesCount > 0)
            {
                this.reserved2 = new byte[reservedBytesCount];
                bytesRead += stream.Read(this.reserved2, 0, reservedBytesCount);
            }
            else if (reservedBytesCount == 0)
            {
                this.reserved2 = null;
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The deserialization operation should be successful.");
            }

            this.auxBytes = new byte[this.auxBytesCount];
            bytesRead += stream.Read(this.auxBytes, 0, (int)this.auxBytesCount);
            AdapterHelper.Site.Assert.AreEqual(size, bytesRead, "The stream size is not equal to the bytes to read. The stream size is {0} and the read bytes is {1}.", size, bytesRead);

            return bytesRead;
        }
    }
}