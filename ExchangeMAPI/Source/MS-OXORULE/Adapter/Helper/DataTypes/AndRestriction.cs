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
    /// The AndRestriction structure describes an AND restriction, which is used to join a group of restrictions using a logical AND operation.
    /// The result of an AndRestriction is TRUE if all of its child restrictions evaluate to TRUE, and FALSE if any child restriction evaluates to FALSE.
    /// </summary>
    public class AndRestriction : Restrictions
    {
        /// <summary>
        /// Array of restriction structures. This field MUST contain RestrictCount structures.
        /// </summary>
        private IRestriction[] restricts;

        /// <summary>
        /// This value specifies how many restriction structures are present in Restricts. The width of this field is 16 bits in the context of ROPs, and 32 bits in the context of extended rules.
        /// </summary>
        private object restrictCount;

        /// <summary>
        /// Initializes a new instance of the AndRestriction class.
        /// </summary>
        public AndRestriction()
        {
            this.RestrictType = RestrictionType.AndRestriction;
            this.CountType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the AndRestriction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public AndRestriction(CountByte countType)
        {
            this.RestrictType = RestrictionType.AndRestriction;
            this.CountType = countType;
        }

        /// <summary>
        /// Gets or sets array of restriction structures. This field MUST contain RestrictCount structures.
        /// </summary>
        public IRestriction[] Restricts
        {
            get { return this.restricts; }
            set { this.restricts = value; }
        }

        /// <summary>
        /// Gets or sets this value specifies how many restriction structures are present in Restricts. The width of this field is 16 bits in the context of ROPs, and 32 bits in the context of extended rules.
        /// </summary>
        public object RestrictCount
        {
            get
            {
                return this.restrictCount;
            }

            set
            {
                if (this.CountType == CountByte.TwoBytesCount)
                {
                    if (value is int)
                    {
                        this.restrictCount = (ushort)(int)value;
                    }
                    else
                    {
                        this.restrictCount = (ushort)value;
                    }
                }
                else
                {
                    if (value is int)
                    {
                        this.restrictCount = (uint)(int)value;
                    }
                    else
                    {
                        this.restrictCount = (uint)value;
                    }
                }
            }
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
                (byte)this.RestrictType
            };
            bytes.AddRange(BitConverter.GetBytes(this.CountType == CountByte.TwoBytesCount ? (ushort)this.RestrictCount : (uint)this.RestrictCount));
            foreach (IRestriction restriction in this.Restricts)
            {
                bytes.AddRange(restriction.Serialize());
            }

            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public override uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)bufferReader.ReadByte();
            this.restrictCount = (this.CountType == CountByte.TwoBytesCount) ? bufferReader.ReadUInt16() : bufferReader.ReadUInt32();
            int count = (this.CountType == CountByte.TwoBytesCount) ? (int)(ushort)this.RestrictCount : (int)(uint)this.RestrictCount;
            this.Restricts = new IRestriction[count];

            uint size = bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();
            for (int i = 0; i < count; i++)
            {
                RestrictionType restrictionType = (RestrictionType)tmpArray[0];
                switch (restrictionType)
                {
                    case RestrictionType.AndRestriction:
                        this.Restricts[i] = new AndRestriction(this.CountType);
                        break;
                    case RestrictionType.BitMaskRestriction:
                        this.Restricts[i] = new BitMaskRestriction();
                        break;
                    case RestrictionType.CommentRestriction:
                        this.Restricts[i] = new CommentRestriction(this.CountType);
                        break;
                    case RestrictionType.ComparePropertiesRestriction:
                        this.Restricts[i] = new ComparePropertiesRestriction();
                        break;
                    case RestrictionType.ContentRestriction:
                        this.Restricts[i] = new ContentRestriction();
                        break;
                    case RestrictionType.CountRestriction:
                        this.Restricts[i] = new CountRestriction(this.CountType);
                        break;
                    case RestrictionType.ExistRestriction:
                        this.Restricts[i] = new ExistRestriction();
                        break;
                    case RestrictionType.NotRestriction:
                        this.Restricts[i] = new NotRestriction(this.CountType);
                        break;
                    case RestrictionType.OrRestriction:
                        this.Restricts[i] = new OrRestriction(this.CountType);
                        break;
                    case RestrictionType.PropertyRestriction:
                        this.Restricts[i] = new PropertyRestriction();
                        break;
                    case RestrictionType.SizeRestriction:
                        this.Restricts[i] = new SizeRestriction();
                        break;
                    case RestrictionType.SubObjectRestriction:
                        this.Restricts[i] = new SubObjectRestriction(this.CountType);
                        break;
                }

                uint tmpLength = this.Restricts[i].Deserialize(tmpArray);
                size += tmpLength;
                bufferReader = new BufferReader(tmpArray);
                tmpArray = bufferReader.ReadBytes(tmpLength, (uint)(tmpArray.Length - tmpLength));
            }

            return size;
        }
    }
}