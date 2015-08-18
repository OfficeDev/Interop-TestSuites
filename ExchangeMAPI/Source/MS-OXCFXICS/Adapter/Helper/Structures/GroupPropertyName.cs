namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.IO;
    using System.Text;

    /// <summary>
    /// The GroupPropertyName.
    /// </summary>
    [SerializableObjectAttribute(true, true)]
    public class GroupPropertyName : SerializableBase
    {
        /// <summary>
        /// The GUID that identifies the property set for the named property.
        /// </summary>
        private Guid guid;

        /// <summary>
        /// The following are possible values for the Kind field: 
        /// Name    Value
        /// 0x00000000  The property is identified by the LID field. 
        /// 0x00000001  The property is identified by the Name field.
        /// </summary>
        private uint kind;

        /// <summary>
        ///  Present only if Kind is set to 0x00000000. 
        ///  An unsigned integer that identifies the named property 
        ///  within its property set.
        /// </summary>
        private uint? lid;

        /// <summary>
        ///  Present only if Kind is set to 0x00000001.
        ///  Identifies the number of bytes in the Name string.
        /// </summary>
        private uint? nameSize;

        /// <summary>
        /// Present only if Kind is set to 0x00000001. 
        /// A Unicode (UTF-16) string, followed by two zero bytes 
        /// as a null terminator, that identifies the property 
        /// within its property set. 
        /// </summary>
        private char[] name;

        /// <summary>
        /// Gets or sets the GUID.
        /// </summary>
        public Guid GUID
        {
            get
            {
                return this.guid;
            }

            set
            {
                this.guid = value;
            }
        }

        /// <summary>
        /// Gets or sets the kind.
        /// </summary>
        public uint Kind
        {
            get
            {
                return this.kind;
            }

            set
            {
                this.kind = value;
            }
        }

        /// <summary>
        /// Gets or sets the lid.
        /// </summary>
        public uint? LID
        {
            get
            {
                return this.lid;
            }

            set
            {
                this.lid = value;
            }
        }

        /// <summary>
        /// Gets or sets the nameSize.
        /// </summary>
        public uint? NameSize
        {
            get
            {
                return this.nameSize;
            }

            set
            {
                this.nameSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public char[] Name
        {
            get 
            {
                return this.name;
            }

            set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Serialize fields to a stream.
        /// </summary>
        /// <param name="stream">The stream where serialized instance will be wrote.</param>
        /// <returns>Bytes wrote to the stream.</returns>
        public override int Serialize(Stream stream)
        {
            int bytesWritten = 0;
            bytesWritten += StreamHelper.WriteGuid(stream, this.guid);
            bytesWritten += StreamHelper.WriteUInt32(stream, this.kind);
            if (this.kind == 0x00000000)
            {
                AdapterHelper.Site.Assert.IsNotNull(this.lid, "The value of GroupPropertyName.lid should not be null.");
                bytesWritten += StreamHelper.WriteUInt32(stream, (uint)this.lid);
            }
            else if (this.kind == 0x00000001)
            {
                AdapterHelper.Site.Assert.IsNotNull(this.nameSize, "The value of GroupPropertyName.nameSize is null.");

                bytesWritten += StreamHelper.WriteUInt32(stream, (uint)this.nameSize);
                byte[] buffer = Encoding.Unicode.GetBytes(this.name, 0, this.name.Length);
                bytesWritten += StreamHelper.WriteBuffer(stream, buffer);
            }

            return bytesWritten;
        }

        /// <summary>
        /// Deserialize fields in this class from a stream.
        /// </summary>
        /// <param name="stream">Stream contains a serialized instance of this class.</param>
        /// <param name="size">The number of bytes can read if -1, no limitation. MUST be -1.</param>
        /// <returns>Bytes have been read from the stream.</returns>
        public override int Deserialize(Stream stream, int size)
        {
            int bytesRead = 0;
            AdapterHelper.Site.Assert.AreEqual(-1, size, "The size value MUST be -1, the actual value is {0}.", size);

            this.guid = StreamHelper.ReadGuid(stream);
            bytesRead += 0x10;
            this.kind = StreamHelper.ReadUInt32(stream);
            if (this.kind == 0x00000000)
            {
                this.nameSize = null;
                this.name = null;
                this.lid = StreamHelper.ReadUInt32(stream);
                bytesRead += 4;
            }
            else if (this.kind == 0x00000001)
            {
                this.lid = null;
                this.nameSize = StreamHelper.ReadUInt32(stream);
                bytesRead += 4;
                byte[] buffer = new byte[(int)this.nameSize];
                stream.Read(buffer, 0, (int)this.nameSize);
                this.name = Encoding.Unicode.GetChars(buffer);
                bytesRead += (int)this.nameSize;
            }

            return bytesRead;
        }
    }
}