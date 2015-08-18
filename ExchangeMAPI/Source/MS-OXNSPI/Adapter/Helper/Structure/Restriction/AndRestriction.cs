namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The AndRestriction class.
    /// </summary>
    public class AndRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the AndRestriction class.
        /// </summary>
        public AndRestriction()
        {
            this.RestrictType = Restrictions.AndRestriction;
        }

        /// <summary>
        /// Gets or sets an unsigned integer specifies how many restriction structures are present in the Restricts field.
        /// </summary>
        public uint RestrictCount
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets an array of restriction structures.
        /// </summary>
        public byte[][] Restricts
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
            this.RestrictType = (Restrictions)restrictionData[index];

            if (this.RestrictType != Restrictions.AndRestriction)
            {
                throw new ArgumentException("The restrict type is not AndRestriction type.");
            }

            index++;

            this.RestrictCount = BitConverter.ToUInt32(restrictionData, index);
            index += 4;
        }

        /// <summary>
        /// Serialize the Restriction data.
        /// </summary>
        /// <returns>Format the type of Restriction data to byte array.</returns>
        public override byte[] Serialize()
        {
            List<byte> restrictionData = new List<byte>();

            restrictionData.Add((byte)this.RestrictType);
            restrictionData.AddRange(BitConverter.GetBytes(this.RestrictCount));

            for (int i = 0; i < this.RestrictCount; i++)
            {
                restrictionData.AddRange(this.Restricts[i]);  
            }

            return restrictionData.ToArray();
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            int size = 0;
            size += sizeof(byte) + sizeof(uint);

            return size;
        }
    }
}