namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The PropertyRestriction class.
    /// </summary>
    public class PropertyRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the PropertyRestriction class.
        /// </summary>
        public PropertyRestriction()
        {
            this.RestrictType = Restrictions.PropertyRestriction;
        }

        /// <summary>
        /// Gets or sets an unsigned integer. This value indicates the relational operator that is used to compare the property on the object with the value of the TaggedValue..
        /// </summary>
        public byte RelOp
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value indicates the type of restriction (2) and MUST be set to 0x08.
        /// </summary>
        public PropertyTag PropTag
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a TaggedValue structure.
        /// </summary>
        public TaggedPropertyValue TaggedValue
        {
            get;
            set;
        }

        /// <summary>
        /// Deserialize the PropertyRestriction data.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        public override void Deserialize(byte[] restrictionData)
        {
            int index = 0;
            this.RestrictType = (Restrictions)restrictionData[index];

            if (this.RestrictType != Restrictions.PropertyRestriction)
            {
                throw new ArgumentException("The restrict type is not PropertyRestriction type.");
            }

            index++;

            this.RelOp = restrictionData[index];
            index++;

            index += this.PropTag.Deserialize(restrictionData, index);
            Context.Instance.PropertyBytes = restrictionData;
            Context.Instance.CurIndex = index;
            Context.Instance.CurProperty = new Property(PropertyType.PtypUnspecified);
            TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue();
            taggedPropertyValue.Parse(Context.Instance);
            this.TaggedValue = taggedPropertyValue;
        }

        /// <summary>
        /// Serialize the Restriction data.
        /// </summary>
        /// <returns>Format the type of Restriction data to byte array.</returns>
        public override byte[] Serialize()
        {
            List<byte> restrictionData = new List<byte>();
            restrictionData.Add((byte)this.RestrictType);
            restrictionData.Add((byte)this.RelOp);
            restrictionData.AddRange(this.PropTag.Serialize());
            restrictionData.AddRange(this.TaggedValue.Serialize());
           
            return restrictionData.ToArray();
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            int size = 0;
            size += sizeof(byte) + sizeof(byte) + this.PropTag.Size() + this.TaggedValue.Size();

            return size;
        }
    }
}