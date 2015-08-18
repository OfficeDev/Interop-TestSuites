namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Contain PropertyTags. 
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class PropertyGroup : SerializableBase
    {
        /// <summary>
        /// An unsigned 32-bit integer value that specifies 
        /// how many PropertyTag structures are present in PropertyTags
        /// </summary>
        private uint propertyTagCount;

        /// <summary>
        /// A variable length array of PropertyTag GroupPropertyName tuple,
        /// If the TUPLE's PropertyTag is not a named property, GroupPropertyName is null.
        /// else GroupPropertyName is not null.
        /// </summary>
        private Tuple<PropertyTag, GroupPropertyName>[] propertyTags;

        /// <summary>
        /// Gets or sets the propertyTagCount.
        /// </summary>
        public uint PropertyTagCount
        {
            get
            {
                return this.propertyTagCount;
            }

            set
            {
                this.propertyTagCount = value;
            }
        }

        /// <summary>
        /// Gets or sets the propertyTags.
        /// </summary>
        public Tuple<PropertyTag, GroupPropertyName>[] PropertyTags
        {
            get
            {
                return this.propertyTags;
            }

            set
            {
                this.propertyTags = value;
            }
        }

        /// <summary>
        /// Serialize current instance to a stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The number of bytes written to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            int size = 0;
            StreamHelper.WriteUInt32(stream, this.propertyTagCount);
            size += 4;
            AdapterHelper.Site.Assert.AreEqual((int)this.propertyTagCount, this.propertyTags.Length, "This field MUST contain PropertyTagCount tags. The expected count is {0}, the actual count is {1}.", this.propertyTagCount, this.propertyTags.Length);

            for (int i = 0; i < this.propertyTagCount; i++)
            {
                PropertyTag tag = this.propertyTags[i].Item1;
                StreamHelper.WriteUInt16(stream, tag.PropertyType);
                StreamHelper.WriteUInt16(stream, tag.PropertyId);
                size += 4;
                if (this.IsNamedProperty(tag))
                {
                    GroupPropertyName name = this.propertyTags[i].Item2;
                    AdapterHelper.Site.Assert.IsNotNull(name, "The property name should not be null.");
                    size += name.Serialize(stream);
                }
            }

            return size;
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
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual value is {0}.", size);

            int bytesRead = 0;
            int i;
            this.propertyTagCount = StreamHelper.ReadUInt32(stream);
            bytesRead += 4;
            this.propertyTags = new Tuple<PropertyTag, GroupPropertyName>[this.propertyTagCount];
            for (i = 0; i < this.propertyTagCount; i++)
            {
                PropertyTag tag = new PropertyTag
                {
                    PropertyType = StreamHelper.ReadUInt16(stream)
                };
                bytesRead += 2;
                tag.PropertyId = StreamHelper.ReadUInt16(stream);
                bytesRead += 2;
                GroupPropertyName name = null;
                if (this.IsNamedProperty(tag))
                {
                    name = new GroupPropertyName();
                    bytesRead += name.Deserialize(stream, -1);
                }

                this.propertyTags[i] = new Tuple<PropertyTag, GroupPropertyName>(tag, name);
            }

            if (size >= 0 && bytesRead > size)
            {
                AdapterHelper.Site.Assert.Fail("The bytes length to read is larger than stream size, the stream size is {0} and the bytes to read is {1}.", size, bytesRead);
            }

            return bytesRead;
        }

        /// <summary>
        /// Decide whether a propertyTag is a named property,
        /// PropertyId greater than or equal to 0x8000
        /// </summary>
        /// <param name="tag">The PropertyTag to be decided.</param>
        /// <returns>If propertyTag's propertyId >=0x8000, return true, else false.</returns>
        private bool IsNamedProperty(PropertyTag tag)
        {
            return tag.PropertyId >= 0x8000;
        }
    }
}