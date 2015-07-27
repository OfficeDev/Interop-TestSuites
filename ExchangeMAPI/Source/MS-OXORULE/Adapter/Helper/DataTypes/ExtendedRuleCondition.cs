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
    /// Extended rule conditions contain information about any named properties contained inside of them. 
    /// </summary>
    public class ExtendedRuleCondition
    {
        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// Specifies information about named properties used in this condition.
        /// </summary>
        private NamedPropertyInfo namedPropertyInformation;

        /// <summary>
        /// A structure containing the condition to be evaluated, represented as a Restriction structure. The format of this Restriction structure is defined in [MS-OXCDATA] section 2.12, using a 4-byte COUNT value.
        /// </summary>
        private IRestriction ruleRestriction;

        /// <summary>
        /// Initializes a new instance of the ExtendedRuleCondition class.
        /// </summary>
        public ExtendedRuleCondition()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the ExtendedRuleCondition class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public ExtendedRuleCondition(CountByte countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets named properties used in this condition.
        /// </summary>
        public NamedPropertyInfo NamedPropertyInformation
        {
            get { return this.namedPropertyInformation; }
            set { this.namedPropertyInformation = value; }
        }

        /// <summary>
        /// Gets or sets a structure containing the condition to be evaluated, represented as a Restriction structure. The format of this Restriction structure is defined in [MS-OXCDATA] section 2.12, using a 4-byte COUNT value.
        /// </summary>
        public IRestriction RuleRestriction
        {
            get { return this.ruleRestriction; }
            set { this.ruleRestriction = value; }
        }

        /// <summary>
        /// Gets type of Count
        /// </summary>
        public CountByte CountType
        {
            get
            {
                return this.countType;
            }
        }

        /// <summary>
        /// Get the total Size of ExtendedRuleCondition
        /// </summary>
        /// <returns>The Size of Restriction buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> bytes = new List<byte>();
            bytes.AddRange(this.NamedPropertyInformation.Serialize());
            bytes.AddRange(this.RuleRestriction.Serialize());
            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to an ExtendedRuleCondition instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            this.NamedPropertyInformation = new NamedPropertyInfo();
            uint namedPropInfoLength = this.NamedPropertyInformation.Deserialize(buffer);
            BufferReader bufferReader = new BufferReader(buffer);
            byte[] tmpArray = bufferReader.ReadBytes(namedPropInfoLength, (uint)(buffer.Length - namedPropInfoLength));
            RestrictionType restrictionType = (RestrictionType)tmpArray[0];
            switch (restrictionType)
            {
                case RestrictionType.AndRestriction:
                    this.RuleRestriction = new AndRestriction(this.CountType);
                    break;
                case RestrictionType.BitMaskRestriction:
                    this.RuleRestriction = new BitMaskRestriction();
                    break;
                case RestrictionType.CommentRestriction:
                    this.RuleRestriction = new CommentRestriction(this.CountType);
                    break;
                case RestrictionType.ComparePropertiesRestriction:
                    this.RuleRestriction = new ComparePropertiesRestriction();
                    break;
                case RestrictionType.ContentRestriction:
                    this.RuleRestriction = new ContentRestriction();
                    break;
                case RestrictionType.CountRestriction:
                    this.RuleRestriction = new CountRestriction(this.CountType);
                    break;
                case RestrictionType.ExistRestriction:
                    this.RuleRestriction = new ExistRestriction();
                    break;
                case RestrictionType.NotRestriction:
                    this.RuleRestriction = new NotRestriction(this.CountType);
                    break;
                case RestrictionType.OrRestriction:
                    this.RuleRestriction = new OrRestriction(this.CountType);
                    break;
                case RestrictionType.PropertyRestriction:
                    this.RuleRestriction = new PropertyRestriction();
                    break;
                case RestrictionType.SizeRestriction:
                    this.RuleRestriction = new SizeRestriction();
                    break;
                case RestrictionType.SubObjectRestriction:
                    this.RuleRestriction = new SubObjectRestriction(this.CountType);
                    break;
            }

            uint totalLength = this.RuleRestriction.Deserialize(tmpArray);
            totalLength += namedPropInfoLength;
            return totalLength;
        }
    }
}