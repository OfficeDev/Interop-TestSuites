namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Represents CN structure contains a change number that identifies a version of a messaging object. 
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class CN : SerializableBase
    {
        /// <summary>
        /// A 16-bit unsigned integer identifying the server replica in which the messaging object was last changed.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private uint replicaId;

        /// <summary>
        /// An unsigned 48-bit integer identifying the change to the messaging object.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private ulong globalCounter;

        /// <summary>
        /// Gets or sets the replicaId.
        /// </summary>
        public uint ReplicaId
        {
            get
            {
                return this.replicaId;
            }

            set
            {
                this.replicaId = value;
            }
        }

        /// <summary>
        /// Gets or sets the globalCounter.
        /// </summary>
        public ulong GlobalCounter
        {
            get
            {
                return this.globalCounter;
            }

            set
            {
                this.globalCounter = value;
            }
        }

        /// <summary>
        /// Deserialize an object from a stream
        /// </summary>
        /// <param name="stream">A stream contains object fields.</param>
        /// <param name="size">Max length can used by this deserialization
        /// if -1 no limitation except stream length.
        /// </param>
        /// <returns>The number of bytes read from the stream.</returns>
        public override int Deserialize(System.IO.Stream stream, int size)
        {
            int bytesRead = 0;
            this.replicaId = StreamHelper.ReadUInt16(stream);
            bytesRead += 2;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            byte[] globalCounterByte = new byte[8];
            stream.Read(globalCounterByte, 0, 6);
            this.globalCounter = BitConverter.ToUInt64(globalCounterByte, 0);
            bytesRead += 6;
            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            return bytesRead;
        }
    }
}