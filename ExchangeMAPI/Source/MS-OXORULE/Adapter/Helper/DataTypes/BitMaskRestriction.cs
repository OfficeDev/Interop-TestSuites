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
    /// The BitMaskRestriction structure describes a bitmask restriction, which performs a bitwise AND operation and compares the result with zero.
    /// </summary>
    public class BitMaskRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 8-bit integer. The value specifies how the server MUST perform the masking operation. 
        /// </summary>
        private BitmapRelOpValue bitmapRelOp;

        /// <summary>
        /// Unsigned 32-bit integer. This value is the PropertyTag of the property to be tested. Its property Type MUST be single-valued int.
        /// </summary>
        private PropertyTag propTag;

        /// <summary>
        /// Unsigned 32 bit integer. The bitmask to use for the AND operation.
        /// </summary>
        private uint mask;

        /// <summary>
        /// Initializes a new instance of the BitMaskRestriction class.
        /// </summary>
        public BitMaskRestriction()
        {
            this.RestrictType = RestrictionType.BitMaskRestriction;
        }

        /// <summary>
        /// Specifies how the server MUST perform the masking operation. 
        /// </summary>
        public enum BitmapRelOpValue : byte
        {
            /// <summary>
            /// Perform a bitwise AND operation of the value of Mask with the value of the property PropTag and test for being equal to zero.
            /// </summary>
            BMR_EQZ,

            /// <summary>
            /// Perform a bitwise AND operation of the value of Mask with the value of the property PropTag and test for NOT being equal to zero.
            /// </summary>
            BMR_NEZ
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. The value specifies how the server MUST perform the masking operation. 
        /// </summary>
        public BitmapRelOpValue BitmapRelOp
        {
            get { return this.bitmapRelOp; }
            set { this.bitmapRelOp = value; }
        }

        /// <summary>
        /// Gets or sets unsigned 32-bit integer. This value is the PropertyTag of the property to be tested. Its property Type MUST be single-valued int.
        /// </summary>
        public PropertyTag PropTag
        {
            get { return this.propTag; }
            set { this.propTag = value; }
        }

        /// <summary>
        /// Gets or sets unsigned 32 bit integer. The bitmask to use for the AND operation.
        /// </summary>
        public uint Mask
        {
            get { return this.mask; }
            set { this.mask = value; }
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
                (byte)RestrictType, (byte)this.BitmapRelOp
            };
            bytes.AddRange(this.PropTag.Serialize());
            bytes.AddRange(BitConverter.GetBytes(this.Mask));

            return bytes.ToArray();
        }

        /// <summary>
        /// De-serialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public override uint Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)reader.ReadByte();
            this.BitmapRelOp = (BitmapRelOpValue)reader.ReadByte();
            this.propTag.PropertyId = reader.ReadUInt16();
            this.propTag.PropertyType = reader.ReadUInt16();
            this.Mask = reader.ReadUInt32();

            return reader.Position;
        }
    }
}