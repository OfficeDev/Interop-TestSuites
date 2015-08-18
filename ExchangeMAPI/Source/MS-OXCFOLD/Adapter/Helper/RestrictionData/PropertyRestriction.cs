namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;

    /// <summary>
    /// The PropertyRestriction structure.
    /// </summary>
    public class PropertyRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the PropertyRestriction class.
        /// </summary>
        public PropertyRestriction()
        {
            this.RestrictType = RestrictType.PropertyRestriction;
        }

        /// <summary>
        /// Gets or sets the value indicates the relational operator that is used to compare the property on the object with the value of the TaggedValue field.
        /// </summary>
        public byte RelOp
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value indicates the property tag of the property that MUST be compared.
        /// </summary>
        public uint PropTag
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the TaggedValue structure, this structure describes the property value to be compared with. 
        /// </summary>
        public byte[] TaggedValue
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

            if (this.RestrictType != RestrictType.PropertyRestriction)
            {
                throw new ArgumentException("The restrict type is not PropertyRestriction type.");
            }

            index++;

            this.RelOp = restrictionData[index];
            index++;

            this.PropTag = BitConverter.ToUInt32(restrictionData, index);
            index += 4;

            int taggedValueLength = restrictionData.Length - (2 * sizeof(byte)) - sizeof(uint);
            this.TaggedValue = new byte[taggedValueLength];
            
            for (int i = 0; i < taggedValueLength; i++)
            {
                this.TaggedValue[i] = restrictionData[index];
                index++;
            }
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
            restrictionData[index++] = this.RelOp;

            Array.Copy(BitConverter.GetBytes(this.PropTag), 0, restrictionData, index, BitConverter.GetBytes(this.PropTag).Length);
            index += 4;

            for (int i = 0; i < this.TaggedValue.Length; i++)
            {
                restrictionData[index++] = this.TaggedValue[i];
            }

            return restrictionData;
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            return sizeof(byte) + sizeof(byte) + sizeof(uint) + this.TaggedValue.Length;
        }
    }
}