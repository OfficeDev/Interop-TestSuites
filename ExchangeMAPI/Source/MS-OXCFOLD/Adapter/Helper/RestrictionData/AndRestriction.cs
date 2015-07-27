//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The AndRestriction structure.
    /// </summary>
    public class AndRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the AndRestriction class.
        /// </summary>
        public AndRestriction()
        {
            this.RestrictType = RestrictType.AndRestriction;
        }

        /// <summary>
        /// Gets or sets the value specifies how many restriction structures are present in the Restricts field.
        /// </summary>
        public short RestrictCount
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets an array of restriction structures. 
        /// </summary>
        public List<Restriction> Restricts
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

            if (this.RestrictType != RestrictType.AndRestriction)
            {
                throw new ArgumentException("The restrict type is not AndRestriction type.");
            }

            index++;

            this.RestrictCount = BitConverter.ToInt16(restrictionData, index);
            index += 2;

            int count = 0;
            this.Restricts = new List<Restriction>();
            while (count < this.RestrictCount)
            {
                byte[] subData = new byte[restrictionData.Length - index];
                Array.Copy(restrictionData, index, subData, 0, subData.Length);
                Restriction restrictsTemp = RestrictsFactory.Deserialize(subData);
                this.Restricts.Add(restrictsTemp);
                count++;
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
            Array.Copy(BitConverter.GetBytes(this.RestrictCount), 0, restrictionData, index, BitConverter.GetBytes(this.RestrictCount).Length);
            index += 2;

            foreach (Restriction restrict in this.Restricts)
            {
                Array.Copy(restrict.Serialize(), 0, restrictionData, index, restrict.Serialize().Length);
                index += restrict.Serialize().Length;
            }

            return restrictionData;
        }

        /// <summary>
        /// Get the size of the restriction data.
        /// </summary>
        /// <returns>Returns the size of restriction data.</returns>
        public override int Size()
        {
            int size = 0;
            size += sizeof(byte) + sizeof(short);

            foreach (Restriction restrict in this.Restricts)
            {
                size += restrict.Size();
            }

            return size;
        }
    }
}