namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Represents an external identifier for an entity within a data store.
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class XID : SerializableBase
    {
        /// <summary>
        /// A 128-bit GUID that identifies the namespace 
        /// that the identifier specified by LocalId belongs to
        /// </summary>
        [SerializableFieldAttribute(1)]
        private Guid namespaceGuid;
        
        /// <summary>
        /// A variable binary value that contains the ID of
        /// the entity in the namespace specified by NamespaceGuid.
        /// </summary>
        [SerializableFieldAttribute(2)]
        private byte[] localId;

        /// <summary>
        /// Gets or sets the nameSpaceGuid.
        /// </summary>
        public Guid NamespaceGuid
        {
            get
            {
                return this.namespaceGuid;
            }

            set
            {
                this.namespaceGuid = value;
            }
        }

        /// <summary>
        /// Gets or sets the localId.
        /// </summary>
        public byte[] LocalId
        {
            get
            {
                return this.localId;
            }

            set
            {
                this.localId = value;
            }
        }

        /// <summary>
        /// Deserialize an object from a stream.
        /// </summary>
        /// <param name="stream">A stream contains object fields.</param>
        /// <param name="size">Max length can used by this deserialization
        /// if -1 no limitation except stream length.
        /// </param>
        /// <returns>The number of bytes read from the stream.</returns>
        public override int Deserialize(System.IO.Stream stream, int size)
        {
            if (size == -1)
            {
                size = (int)(stream.Length - stream.Position);
            }

            this.namespaceGuid = StreamHelper.ReadGuid(stream);
            int bufferSize = size - GuidSize;
            if (bufferSize >= 0)
            {
                this.localId = new byte[bufferSize];
                stream.Read(this.localId, 0, this.localId.Length);
                return size;
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The specified size should be larger than the size of a Guid in bytes.", "size");
                return -1;
            }
        }
    }
}