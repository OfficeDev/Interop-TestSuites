namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;

    /// <summary>
    /// The PropertyGroupInfo structure describes a single property 
    /// mapping between a group index and property tags within a property group.
    /// </summary>
    [SerializableObjectAttribute(false, true)]
    public class PropertyGroupInfo : SerializableBase
    {
        /// <summary>
        /// (4 bytes):  An unsigned 32-bit integer value that identifies 
        /// a property mapping within the current synchronization download.
        /// </summary>
        [SerializableFieldAttribute(1)]
        private uint groupId; 

        /// <summary>
        ///  (4 bytes):  This value MUST be set to "0x00000000".
        /// </summary>
        [SerializableFieldAttribute(2)]
        private uint reserved;

        /// <summary>
        ///  (4 bytes):  An unsigned 32-bit integer value that 
        ///  specifies how many PropertyGroup structures are present 
        ///  in the Groups field. MUST NOT be zero (0x00000000).
        /// </summary>
        [SerializableFieldAttribute(3)]
        private uint groupCount;

        /// <summary>
        /// An array of PropertyGroup structures. 
        /// This field MUST contain GroupCount PropertyGroup elements.
        /// </summary>
        [SerializableFieldAttribute(4)]
        private PropertyGroup[] groups;

        /// <summary>
        /// Gets or sets the groupId.
        /// </summary>
        public uint GroupId
        {
            get
            {
                return this.groupId;
            }

            set
            {
                this.groupId = value;
            }
        }

        /// <summary>
        /// Gets or sets the reserved.
        /// </summary>
        public uint Reserved
        {
            get
            {
                return this.reserved;
            }

            set
            {
                this.reserved = value;
            }
        }

        /// <summary>
        /// Gets or sets the groupCount.
        /// </summary>
        public uint GroupCount
        {
            get
            {
                return this.groupCount;
            }

            set
            {
                this.groupCount = value;
            }
        }

        /// <summary>
        /// Gets or sets the groups.
        /// </summary>
        public PropertyGroup[] Groups
        {
            get
            {
                return this.groups;
            }

            set
            {
                this.groups = value;
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
        public override int Deserialize(Stream stream, int size)
        {
            int i;
            int bytesRead = 0;

            this.groupId = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            this.reserved = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            this.groupCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;

            this.groups = new PropertyGroup[this.groupCount];
            for (i = 0; i < this.groupCount; i++)
            {
                this.groups[i] = new PropertyGroup();
                bytesRead += this.groups[i].Deserialize(stream, -1);
            }

            if (size > 0 && bytesRead != size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read should be equal to the stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            return bytesRead;
        }
    }
}