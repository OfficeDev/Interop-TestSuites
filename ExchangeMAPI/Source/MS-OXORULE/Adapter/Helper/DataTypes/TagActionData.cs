namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// An OP_TAG action Data Buffer is a PropertyValue structure, packaged as specified in [MS-OXCDATA] section 2.11.2.1.
    /// This structure contains a PropertyTag_r structure.
    /// </summary>
    public class TagActionData : IActionData
    {
        /// <summary>
        /// Encodes the property tag with the value represented 
        /// </summary>
        private PropertyTag propertyTag;

        /// <summary>
        /// This is the actual value of the property represented by the PropertyValue_r structure. The Type value is specified by the PropertyTag field.
        /// </summary>
        private byte[] propertyValue;

        /// <summary>
        /// Gets or sets the property represented by the PropertyValue_r structure. The Type value is specified by the PropertyTag field.
        /// </summary>
        public byte[] PropertyValue
        {
            get { return this.propertyValue; }
            set { this.propertyValue = value; }
        }

        /// <summary>
        /// Gets or sets encodes the property tag with the value represented 
        /// </summary>
        public PropertyTag PropertyTag
        {
            get { return this.propertyTag; }
            set { this.propertyTag = value; }
        }

        /// <summary>
        /// The total Size of this ActionData buffer
        /// </summary>
        /// <returns>Number of bytes in this ActionData buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this ActionData
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> result = new List<byte>();
            result.AddRange(this.PropertyTag.Serialize());
            result.AddRange(this.PropertyValue);
            return result.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to a TagActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contains data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyType = bufferReader.ReadUInt16(),
                PropertyId = bufferReader.ReadUInt16()
            };
            this.PropertyTag = propertyTag;
            uint size = bufferReader.Position;
            this.PropertyValue = AdapterHelper.ReadValueByType(this.PropertyTag.PropertyType, bufferReader.ReadToEnd());
            size += (uint)this.PropertyValue.Length;
            return size;
        }
    }
}