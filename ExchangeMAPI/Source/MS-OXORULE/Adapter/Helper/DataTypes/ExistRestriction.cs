namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The ExistRestriction structure tests whether a particular property value exists on a row of the table.
    /// </summary>
    public class ExistRestriction : Restrictions
    {
        /// <summary>
        /// This value is the PropertyTag of the column to be tested for existence in each row.
        /// </summary>
        private PropertyTag propTag;

        /// <summary>
        /// Initializes a new instance of the ExistRestriction class.
        /// </summary>
        public ExistRestriction()
        {
            this.RestrictType = RestrictionType.ExistRestriction;
        }

        /// <summary>
        /// Gets or sets the PropertyTag of the column to be tested for existence in each row.
        /// </summary>
        public PropertyTag PropTag
        {
            get { return this.propTag; }
            set { this.propTag = value; }
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
                (byte)RestrictType
            };
            bytes.AddRange(this.PropTag.Serialize());

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
            this.propTag.PropertyType = reader.ReadUInt16();
            this.propTag.PropertyId = reader.ReadUInt16();

            return reader.Position;
        }
    }
}