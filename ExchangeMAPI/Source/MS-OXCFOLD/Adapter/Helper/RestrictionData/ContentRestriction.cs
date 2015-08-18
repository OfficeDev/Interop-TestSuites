namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    
    /// <summary>
    /// The ContentRestriction structure.
    /// </summary>
    public class ContentRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the ContentRestriction class.
        /// </summary>
        public ContentRestriction()
        {
            this.RestrictType = RestrictType.ContentRestriction;
        }

        /// <summary>
        /// Gets or sets the field specifies the level of precision that the server enforces when checking for a match against a ContentRestriction structure. 
        /// </summary>
        public FuzzyLevelLowValues FuzzyLevelLow
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the field applies only to string-value properties and can be set to the bit values listed in the following table, in any combination.
        /// </summary>
        public FuzzyLevelHighValues FuzzyLevelHigh
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value indicates the property tag of the column whose value MUST be matched against the value specified in the TaggedValue field.
        /// </summary>
        public PropertyTag PropertyTag
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a TaggedPropertyValue structure, this structure contains the value to be matched
        /// </summary>
        public TaggedPropertyValue TaggedValue
        {
            get;
            set;
        }

        /// <summary>
        /// Deserialize the AndRestriction data.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        public override void Deserialize(byte[] restrictionData)
        {
            int index = 0;
            this.RestrictType = (RestrictType)restrictionData[index];

            if (this.RestrictType != RestrictType.ContentRestriction)
            {
                throw new ArgumentException("The restrict type is not ContentRestriction type.");
            }

            index++;

            this.FuzzyLevelLow = (FuzzyLevelLowValues)BitConverter.ToInt16(restrictionData, index);
            index += 2;

            this.FuzzyLevelHigh = (FuzzyLevelHighValues)BitConverter.ToInt16(restrictionData, index);
            index += 2;

            index += this.PropertyTag.Deserialize(restrictionData, index);

            Context.Instance.PropertyBytes = restrictionData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property((PropertyType)this.PropertyTag.PropertyType);
            this.TaggedValue = new TaggedPropertyValue();
            this.TaggedValue.Parse(Context.Instance);
        }

        /// <summary>
        /// Serialize the Restriction data.
        /// </summary>
        /// <returns>Format the type of Restriction data to byte array.</returns>
        public override byte[] Serialize()
        {
            int index = 0;
            byte[] restrictionData = new byte[this.Size()];

            restrictionData[index++] = (byte)this.RestrictType;
            Array.Copy(BitConverter.GetBytes((short)this.FuzzyLevelLow), 0, restrictionData, index, BitConverter.GetBytes((short)this.FuzzyLevelLow).Length);
            index += 2;
            Array.Copy(BitConverter.GetBytes((short)this.FuzzyLevelHigh), 0, restrictionData, index, BitConverter.GetBytes((short)this.FuzzyLevelHigh).Length);
            index += 2;
           
            Array.Copy(this.PropertyTag.Serialize(), 0, restrictionData, index, this.PropertyTag.Serialize().Length);
            index += this.PropertyTag.Serialize().Length;

            Array.Copy(this.TaggedValue.Serialize(), 0, restrictionData, index, this.TaggedValue.Serialize().Length);
            index += this.TaggedValue.Serialize().Length;

            return restrictionData;
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            return sizeof(byte) + (2 * sizeof(short)) + sizeof(uint) + this.TaggedValue.Size();
        }
    }
}