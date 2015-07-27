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
    /// The CountRestriction class
    /// </summary>
    public class CountRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the limit on the number of matches to be returned when SubRestriction is evaluated.
        /// </summary>
        private uint count;

        /// <summary>
        /// A restriction structure. This field specifies the restriction to be limited.
        /// </summary>
        private Restrictions subRestriction;

        /// <summary>
        /// Initializes a new instance of the CountRestriction class.
        /// </summary>
        public CountRestriction()
        {
            this.RestrictType = RestrictionType.CountRestriction;
            this.CountType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the CountRestriction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public CountRestriction(CountByte countType)
        {
            this.RestrictType = RestrictionType.CountRestriction;
            this.CountType = countType;
        }

        /// <summary>
        ///  Gets or sets Unsigned 32-bit integer. This value specifies the limit on the number of matches to be returned when SubRestriction is evaluated.
        /// </summary>
        public uint Count
        {
            get { return this.count; }
            set { this.count = value; }
        }

        /// <summary>
        /// Gets or sets a restriction structure. This field specifies the restriction to be limited.
        /// </summary>
        public Restrictions SubRestriction
        {
            get { return this.subRestriction; }
            set { this.subRestriction = value; }
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
                (byte)RestrictType
            };
            bytes.AddRange(BitConverter.GetBytes(this.Count));
            bytes.AddRange(this.SubRestriction.Serialize());

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
            this.Count = reader.ReadUInt32();

            uint size = reader.Position;
            byte[] tmpArray = reader.ReadToEnd();

            RestrictionType restrictionType = (RestrictionType)tmpArray[0];
            switch (restrictionType)
            {
                case RestrictionType.AndRestriction:
                    this.SubRestriction = new AndRestriction(this.CountType);
                    break;
                case RestrictionType.BitMaskRestriction:
                    this.SubRestriction = new BitMaskRestriction();
                    break;
                case RestrictionType.CommentRestriction:
                    this.SubRestriction = new CommentRestriction(this.CountType);
                    break;
                case RestrictionType.ComparePropertiesRestriction:
                    this.SubRestriction = new ComparePropertiesRestriction();
                    break;
                case RestrictionType.ContentRestriction:
                    this.SubRestriction = new ContentRestriction();
                    break;
                case RestrictionType.CountRestriction:
                    this.SubRestriction = new CountRestriction(this.CountType);
                    break;
                case RestrictionType.ExistRestriction:
                    this.SubRestriction = new ExistRestriction();
                    break;
                case RestrictionType.NotRestriction:
                    this.SubRestriction = new NotRestriction(this.CountType);
                    break;
                case RestrictionType.OrRestriction:
                    this.SubRestriction = new OrRestriction(this.CountType);
                    break;
                case RestrictionType.PropertyRestriction:
                    this.SubRestriction = new PropertyRestriction();
                    break;
                case RestrictionType.SizeRestriction:
                    this.SubRestriction = new SizeRestriction();
                    break;
                case RestrictionType.SubObjectRestriction:
                    this.SubRestriction = new SubObjectRestriction(this.CountType);
                    break;
            }

            size += this.SubRestriction.Deserialize(tmpArray);

            return size;
        }
    }
}