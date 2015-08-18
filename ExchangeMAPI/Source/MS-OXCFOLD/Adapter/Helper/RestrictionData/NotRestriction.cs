namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;

    /// <summary>
    /// The NotRestriction structure.
    /// </summary>
    public class NotRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the NotRestriction class.
        /// </summary>
        public NotRestriction()
        {
            this.RestrictType = RestrictType.NotRestriction;
        }

        /// <summary>
        /// Gets or sets the value specifies the restriction (2) that the logical NOT operation applies to.
        /// </summary>
        public Restriction Restricts
        {
            get;
            set;
        }

        /// <summary>
        /// Deserialize the NotRestriction data.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        public override void Deserialize(byte[] restrictionData)
        {
            int index = 0;
            this.RestrictType = (RestrictType)restrictionData[index];

            if (this.RestrictType != RestrictType.NotRestriction)
            {
                throw new ArgumentException("The restrict type is not NotRestriction type.");
            }

            index++;

            byte[] subData = new byte[restrictionData.Length - index];
            Array.Copy(restrictionData, index, subData, 0, subData.Length);
            Restriction restrictsTemp = RestrictsFactory.Deserialize(subData);
            this.Restricts = restrictsTemp;
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

            Array.Copy(this.Restricts.Serialize(), 0, restrictionData, index, this.Restricts.Serialize().Length);
            index += this.Restricts.Serialize().Length;

            return restrictionData;
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            return sizeof(byte) + this.Restricts.Size();
        }
    }
}