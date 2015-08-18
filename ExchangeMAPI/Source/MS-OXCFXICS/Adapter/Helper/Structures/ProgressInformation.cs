namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;

    /// <summary>
    /// The ProgressInformation.
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class ProgressInformation : SerializableBase
    {
        /// <summary>
        /// An unsigned 16-bit value that contains a number that identifies 
        /// the binary structure of the data that follows.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private ushort version;

        /// <summary>
        /// The padding.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private ushort padding1;

        /// <summary>
        /// An unsigned 32-bit integer value that contains 
        /// the total number of changes to FAI messages that 
        /// are scheduled for download during the current 
        /// synchronization operation.
        /// </summary>
        [SerializableFieldAttribute(3)]
        private uint faiMessageCount;

        /// <summary>
        /// An unsigned 64-bit integer value that contains 
        /// the size in bytes of all changes to FAI messages 
        /// that are scheduled for download during the current 
        /// synchronization operation.
        /// </summary>
        [SerializableFieldAttribute(4)]
        private ulong faiMessageTotalSize;

        /// <summary>
        /// An unsigned 32-bit integer value that contains 
        /// the total number of changes to normal messages 
        /// that are scheduled for download during the current 
        /// synchronization operation.
        /// </summary>
        [SerializableFieldAttribute(5)]
        private uint normalMessageCount;

        /// <summary>
        /// SHOULD be set to zeros and MUST be ignored by clients.
        /// </summary>
        [SerializableFieldAttribute(6)]
        private uint padding2;

        /// <summary>
        /// An unsigned 64-bit integer value that contains the size 
        /// in bytes of all changes to normal messages  that are scheduled 
        /// for download during the current synchronization operation.
        /// </summary>
        [SerializableFieldAttribute(7)]
        private ulong normalMessageTotalSize;
        #region Properties

        /// <summary>
        /// Gets or sets the version.
        /// </summary>
        public ushort Version
        {
            get
            {
                return this.version;
            }

            set
            {
                this.version = value;
            }
        }

        /// <summary>
        /// Gets or sets the padding1.
        /// </summary>
        public ushort Padding1
        {
            get
            {
                return this.padding1;
            }

            set
            {
                this.padding1 = value;
            }
        }

        /// <summary>
        /// Gets or sets the faiMessageCount.
        /// </summary>
        public uint FAIMessageCount
        {
            get
            {
                return this.faiMessageCount;
            }

            set
            {
                this.faiMessageCount = value;
            }
        }

        /// <summary>
        /// Gets or sets the faiMessageTotalSize.
        /// </summary>
        public ulong FAIMessageTotalSize
        {
            get
            {
                return this.faiMessageTotalSize;
            }

            set
            {
                this.faiMessageTotalSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the  normalMessageCount.
        /// </summary>
        public uint NormalMessageCount
        {
            get
            {
                return this.normalMessageCount;
            }

            set
            {
                this.normalMessageCount = value;
            }
        }

        /// <summary>
        /// Gets or sets the padding2.
        /// </summary>
        public uint Padding2
        {
            get
            {
                return this.padding2;
            }

            set
            {
                this.padding2 = value;
            }
        }

        /// <summary>
        /// Gets or sets the normalMessageTotalSize.
        /// </summary>
        public ulong NormalMessageTotalSize
        {
            get
            {
                return this.normalMessageTotalSize;
            }

            set
            {
                this.normalMessageTotalSize = value;
            }
        }

        #endregion

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">Stream contains a serialized instance of this class.</param>
        /// <param name="size">The number of bytes can read if -1, no limitation.</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            int bytesRead = 0;
            this.version = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.padding1 = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.faiMessageCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.faiMessageTotalSize = StreamHelper.ReadUInt64(stream);
            bytesRead += 8;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.normalMessageCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.padding2 = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            this.normalMessageTotalSize = StreamHelper.ReadUInt64(stream);
            bytesRead += 8;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }
            
            return bytesRead;
        }
    }
}