namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The ComparePropertiesRestriction structure specifies a comparison between the values of two properties using a relational operator.
    /// </summary>
    public class ComparePropertiesRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 8-bit integer. The value indicates the relational operator used to compare the two properties.
        /// </summary>
        private RelationalOperator relOp;

        /// <summary>
        /// This value indicates the property tag of the property that MUST be compared.
        /// </summary>
        private PropertyTag propTag1;

        /// <summary>
        /// This value is the PropertyTag of the second property that MUST be compared.
        /// </summary>
        private PropertyTag propTag2;

        /// <summary>
        /// Initializes a new instance of the ComparePropertiesRestriction class.
        /// </summary>
        public ComparePropertiesRestriction()
        {
            this.RestrictType = RestrictionType.ComparePropertiesRestriction;
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. The value indicates the relational operator used to compare the two properties.
        /// </summary>
        public RelationalOperator RelOp
        {
            get { return this.relOp; }
            set { this.relOp = value; }
        }

        /// <summary>
        /// Gets or sets the property tag of the property that MUST be compared.
        /// </summary>
        public PropertyTag PropTag1
        {
            get { return this.propTag1; }
            set { this.propTag1 = value; }
        }

        /// <summary>
        /// Gets or sets the PropertyTag of the second property that MUST be compared.
        /// </summary>
        public PropertyTag PropTag2
        {
            get { return this.propTag2; }
            set { this.propTag2 = value; }
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer.</returns>
        public override int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this structure
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public override byte[] Serialize()
        {
            List<byte> bytes = new List<byte>
            {
                (byte)RestrictType, (byte)this.RelOp
            };
            bytes.AddRange(this.PropTag1.Serialize());
            bytes.AddRange(this.PropTag2.Serialize());

            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public override uint Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)reader.ReadByte();
            this.RelOp = (RelationalOperator)reader.ReadByte();
            this.propTag1.PropertyId = reader.ReadUInt16();
            this.propTag1.PropertyType = reader.ReadUInt16();
            this.propTag2.PropertyId = reader.ReadUInt16();
            this.propTag2.PropertyType = reader.ReadUInt16();

            return reader.Position;
        }
    }
}