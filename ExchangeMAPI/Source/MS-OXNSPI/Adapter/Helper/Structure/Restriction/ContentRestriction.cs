//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The ContentRestriction class.
    /// </summary>
    public class ContentRestriction : Restriction
    {
        /// <summary>
        /// Initializes a new instance of the ContentRestriction class.
        /// </summary>
        public ContentRestriction()
        {
            this.RestrictType = Restrictions.ContentRestriction;
        }

        /// <summary>
        /// Gets or sets the value indicates the type of restriction (2) and MUST be set to 0x08.
        /// </summary>
        public uint FuzzyLevel
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
        /// Deserialize the OrRestriction data.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        public override void Deserialize(byte[] restrictionData)
        {
            int index = 0;
            this.RestrictType = (Restrictions)restrictionData[index];

            if (this.RestrictType != Restrictions.ContentRestriction)
            {
                throw new ArgumentException("The restrict type is not ContentRestriction type.");
            }

            index++;
            this.FuzzyLevel = BitConverter.ToUInt32(restrictionData, index);
            index += 4;

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
            restrictionData.AddRange(BitConverter.GetBytes(this.FuzzyLevel));
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
            size += sizeof(byte) + sizeof(uint) + this.PropTag.Size() + this.TaggedValue.Size();

            return size;
        }
    }
}
