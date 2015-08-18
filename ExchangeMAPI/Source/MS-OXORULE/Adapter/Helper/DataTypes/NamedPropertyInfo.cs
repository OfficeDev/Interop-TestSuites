namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The Named Property Information Format provides context to any named property tags which are present in the structure 
    /// it precedes. For every distinct (unique) named property used in the structure it precedes, the Named Property 
    /// Information structure MUST contain one PropId-NamedProperty pair. Each PropId field is a property ID that 
    /// MUST have a value of 0x8000 or greater and uniquely identifies the named property within an extended rule.
    /// </summary>
    public class NamedPropertyInfo
    {
        /// <summary>
        /// Specifies the number of named property mappings that are packed in this buffer. If no named properties are used in the structure that follows the Named Property Information buffer, the value of this field MUST be 0x0000.
        /// Note that if there are no named properties to be listed, the Named Property Information reduces to a 2-byte WORD value of 0x0000.
        /// </summary>
        private ushort namedPropsNumber;

        /// <summary>
        /// Array of properties ID.
        /// </summary>
        private uint[] propId;

        /// <summary>
        /// The total Size, in bytes, of the following fields. Only present if NoOfNamedProps is greater than zero.
        /// </summary>
        private uint namedPropertiesSize;

        /// <summary>
        /// Array of PropertyName structure.
        /// </summary>
        private PropertyName[] namedProperty;

        /// <summary>
        /// Gets or sets the number of named property mappings that are packed in this buffer. If no named properties are used in the structure that follows the Named Property Information buffer, the value of this field MUST be 0x0000.
        /// Note that if there are no named properties to be listed, the Named Property Information reduces to a 2-byte WORD value of 0x0000.
        /// </summary>
        public ushort NoOfNamedProps
        {
            get { return this.namedPropsNumber; }
            set { this.namedPropsNumber = value; }
        }

        /// <summary>
        /// Gets or sets array of properties ID.
        /// </summary>
        public uint[] PropId
        {
            get { return this.propId; }
            set { this.propId = value; }
        }

        /// <summary>
        /// Gets or sets the total Size, in bytes, of the following fields. Only present if NoOfNamedProps is greater than zero.
        /// </summary>
        public uint NamedPropertiesSize
        {
            get { return this.namedPropertiesSize; }
            set { this.namedPropertiesSize = value; }
        }

        /// <summary>
        /// Gets or sets array of PropertyName structure.
        /// </summary>
        public PropertyName[] NamedProperty
        {
            get { return this.namedProperty; }
            set { this.namedProperty = value; }
        }

        /// <summary>
        /// Get the total Size of ActionData
        /// </summary>
        /// <returns>The Size of ActionData buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> bytes = new List<byte>();
            bytes.AddRange(BitConverter.GetBytes(this.NoOfNamedProps));
            if (this.NoOfNamedProps == 0x0000)
            {
                return bytes.ToArray();
            }

            for (int i = 0; i < this.NoOfNamedProps; i++)
            {
                bytes.AddRange(BitConverter.GetBytes(this.PropId[i]));
            }

            bytes.AddRange(BitConverter.GetBytes(this.NamedPropertiesSize));

            for (int j = 0; j < this.NoOfNamedProps; j++)
            {
                bytes.AddRange(this.NamedProperty[j].Serialize());
            }

            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to an ActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            this.NoOfNamedProps = bufferReader.ReadUInt16();
            if (this.NoOfNamedProps == 0x0000)
            {
                return bufferReader.Position;
            }

            this.PropId = new uint[this.NoOfNamedProps];
            for (int i = 0; i < this.NoOfNamedProps; i++)
            {
                this.PropId[i] = bufferReader.ReadUInt32();
            }

            this.NamedPropertiesSize = bufferReader.ReadUInt32();
            this.NamedProperty = new PropertyName[this.NoOfNamedProps];

            uint length = bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();
            int startIndex = 0;
            for (int j = 0; j < this.NoOfNamedProps; j++)
            {
                this.NamedProperty[j] = new PropertyName();
                startIndex += this.NamedProperty[j].Deserialize(tmpArray, startIndex);
            }

            length += (uint)startIndex;
            return length;
        }
    }
}