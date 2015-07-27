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
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The SizeRestriction structure describes a Size restriction which compares the Size (in bytes) of a property value with a given Size.
    /// </summary>
    public class SizeRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 8-bit integer. The value indicates the relational operator used in the Size comparison. 
        /// </summary>
        private RelationalOperator relOp;

        /// <summary>
        /// This value indicates the property tag of the property, the Size of whose value are testing.
        /// </summary>
        private PropertyTag propTag;

        /// <summary>
        /// Unsigned 32-bit integer. This value indicates the Size, as a count of bytes, which is to be used in the comparison.
        /// </summary>
        private uint sizeValue;

        /// <summary>
        /// Initializes a new instance of the SizeRestriction class.
        /// </summary>
        public SizeRestriction()
        {
            this.RestrictType = RestrictionType.SizeRestriction;
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. The value indicates the relational operator used in the Size comparison. 
        /// </summary>
        public RelationalOperator RelOp
        {
            get { return this.relOp; }
            set { this.relOp = value; }
        }

        /// <summary>
        /// Gets or sets the property tag of the property, the Size of whose value are testing.
        /// </summary>
        public PropertyTag PropTag
        {
            get { return this.propTag; }
            set { this.propTag = value; }
        }

        /// <summary>
        /// Gets or sets unsigned 32-bit integer. This value indicates the Size, as a count of bytes, that is to be used in the comparison.
        /// </summary>
        public uint SizeValue
        {
            get { return this.sizeValue; }
            set { this.sizeValue = value; }
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
            bytes.AddRange(BitConverter.GetBytes(this.SizeValue));

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
            this.propTag.PropertyId = reader.ReadUInt16();
            this.propTag.PropertyType = reader.ReadUInt16();
            this.SizeValue = reader.ReadUInt32();

            return reader.Position;
        }
    }
}