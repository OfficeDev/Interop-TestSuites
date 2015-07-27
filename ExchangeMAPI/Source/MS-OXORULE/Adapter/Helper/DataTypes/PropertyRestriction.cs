//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The PropertyRestriction structure describes a property restriction that is used to match a constant with the value of a property.
    /// </summary>
    public class PropertyRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 8-bit integer. The value indicates the relational operator that is used to compare the property on the object with TaggedValue. 
        /// </summary>
        private RelationalOperator relOp;

        /// <summary>
        /// This value indicates the property tag of the property that MUST be compared.
        /// </summary>
        private PropertyTag propTag;

        /// <summary>
        /// This structure describes the property value to be compared against.
        /// </summary>
        private TaggedPropertyValue taggedValue;

        /// <summary>
        /// Initializes a new instance of the PropertyRestriction class.
        /// </summary>
        public PropertyRestriction()
        {
            this.RestrictType = RestrictionType.PropertyRestriction;
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. The value indicates the relational operator that is used to compare the property on the object with TaggedValue. 
        /// </summary>
        public RelationalOperator RelOp
        {
            get { return this.relOp; }
            set { this.relOp = value; }
        }

        /// <summary>
        /// Gets or sets the property tag of the property that MUST be compared.
        /// </summary>
        public PropertyTag PropTag
        {
            get { return this.propTag; }
            set { this.propTag = value; }
        }

        /// <summary>
        /// Gets or sets the structure describes the property value to be compared against.
        /// </summary>
        public TaggedPropertyValue TaggedValue
        {
            get { return this.taggedValue; }
            set { this.taggedValue = value; }
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
            bytes.AddRange(this.PropTag.Serialize());
            bytes.AddRange(this.TaggedValue.Serialize());

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
            this.propTag.PropertyType = reader.ReadUInt16();
            this.propTag.PropertyId = reader.ReadUInt16();

            uint size = reader.Position;
            byte[] tmpArray = reader.ReadToEnd();
            this.TaggedValue = AdapterHelper.ReadTaggedProperty(tmpArray);
            size += (uint)this.TaggedValue.Size();

            return size;
        }
    }
}