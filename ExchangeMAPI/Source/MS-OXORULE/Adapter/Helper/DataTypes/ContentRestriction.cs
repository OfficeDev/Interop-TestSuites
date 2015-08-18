namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The ContentRestriction structure describes a content restriction, which is used to limit a table view to only those rows that include a column with contents matching a search string.
    /// </summary>
    public class ContentRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 16-bit integer. This field specifies the level of precision that the server enforces when checking for a match against a ContentRestriction
        /// </summary>
        private FuzzyLevelLowValue fuzzyLevelLow;

        /// <summary>
        /// This field applies only to string valued properties
        /// </summary>
        private FuzzyLevelHighValue fuzzyLevelHigh;

        /// <summary>
        /// This value indicates the property tag of the column whose value MUST be matched against the value specified by the TaggedValue field
        /// </summary>
        private PropertyTag propertyTag;

        /// <summary>
        /// A TaggedPropertyValue structure,This structure contains the value to be matched
        /// </summary>
        private TaggedPropertyValue taggedValue;

        /// <summary>
        /// Initializes a new instance of the ContentRestriction class.
        /// </summary>
        public ContentRestriction()
        {
            this.RestrictType = RestrictionType.ContentRestriction;
        }

        /// <summary>
        /// FuzzyLevelLow applies to both binary and string properties and MUST be set to one of the following values.
        /// </summary>
        public enum FuzzyLevelLowValue : ushort
        {
            /// <summary>
            /// The value stored in TaggedValue and the value of the column PropertyTag matches in their entirety.
            /// </summary>
            FL_FULLSTRING = 0x0000,

            /// <summary>
            /// The value stored in TaggedValue matches some portion of the value of the column PropertyTag.
            /// </summary>
            FL_SUBSTRING = 0x0001,

            /// <summary>
            /// The value stored in TaggedValue matches a starting portion of the value of the column PropertyTag.
            /// </summary>
            FL_PREFIX = 0x0002
        }

        /// <summary>
        /// FuzzyLevelHigh can be set to the following bit values in any combination. FuzzyLevelHigh values can be OR'd together.
        /// </summary>
        public enum FuzzyLevelHighValue : ushort
        {
            /// <summary>
            /// The comparison does not consider case.
            /// </summary>
            FL_IGNORECASE = 0x0001,

            /// <summary>
            /// The comparison ignores Unicode-defined nonspacing characters such as diacritical marks.
            /// </summary>
            FL_IGNORENONSPACE = 0x0002,

            /// <summary>
            /// The comparison results in a match whenever possible, ignoring case and nonspacing characters.
            /// </summary>
            FL_LOOSE
        }

        /// <summary>
        /// Gets or sets unsigned 16-bit integer. This field specifies the level of precision that the server enforces when checking for a match against a ContentRestriction
        /// </summary>
        public FuzzyLevelLowValue FuzzyLevelLow
        {
            get { return this.fuzzyLevelLow; }
            set { this.fuzzyLevelLow = value; }
        }

        /// <summary>
        /// Gets or sets the field applies only to string valued properties
        /// </summary>
        public FuzzyLevelHighValue FuzzyLevelHigh
        {
            get { return this.fuzzyLevelHigh; }
            set { this.fuzzyLevelHigh = value; }
        }

        /// <summary>
        /// Gets or sets the property tag of the column whose value MUST be matched against the value specified by the TaggedValue field
        /// </summary>
        public PropertyTag PropertyTag
        {
            get { return this.propertyTag; }
            set { this.propertyTag = value; }
        }

        /// <summary>
        /// Gets or sets a TaggedPropertyValue structure,This structure contains the value to be matched
        /// </summary>
        public TaggedPropertyValue TaggedValue
        {
            get { return this.taggedValue; }
            set { this.taggedValue = value; }
        }

        /// <summary>
        /// Get serialized byte array for this structure
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public override byte[] Serialize()
        {
            List<byte> result = new List<byte>
            {
                (byte)this.RestrictType
            };

            // Add RestrictType

            // Add FuzzyLevelLow
            result.AddRange(BitConverter.GetBytes((ushort)this.FuzzyLevelLow));

            // Add FuzzyLevelHigh
            result.AddRange(BitConverter.GetBytes((ushort)this.FuzzyLevelHigh));

            // Add PropertyTag
            result.AddRange(this.PropertyTag.Serialize());

            // Add TaggedValue
            result.AddRange(this.TaggedValue.Serialize());

            return result.ToArray();
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
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public override uint Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)reader.ReadByte();
            this.FuzzyLevelLow = (FuzzyLevelLowValue)reader.ReadUInt16();
            this.FuzzyLevelHigh = (FuzzyLevelHighValue)reader.ReadUInt16();
            this.propertyTag.PropertyId = reader.ReadUInt16();
            this.propertyTag.PropertyType = reader.ReadUInt16();

            uint size = reader.Position;
            byte[] tmpArray = reader.ReadToEnd();
            this.TaggedValue = AdapterHelper.ReadTaggedProperty(tmpArray);
            size += (uint)this.TaggedValue.Size();

            return size;
        }
    }
}