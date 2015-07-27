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
    /// An extended rule's PidTagExtendedRuleMessageActions property contains additional information about the Version of the rule and the named properties stored in the rule action, as well as information about the actions to be performed by this rule. 
    /// </summary>
    public class ExtendedRuleActions : IActionData
    {
        /// <summary>
        /// Specifies information about named properties used in this action.
        /// </summary>
        private NamedPropertyInfo namedPropertyInformation;

        /// <summary>
        /// Specifies the extended rules Version format. This document defines Version 1, and this value MUST be set to 0x00000001.
        /// </summary>
        private uint ruleVersion;

        /// <summary>
        /// A structure containing the actions to be executed when the rule condition for the rule to which these actions apply evaluates to TRUE. The format of this structure is defined in section 2.2.5, using a 4-byte COUNT value.
        /// </summary>
        private RuleAction ruleActionBuffer;

        /// <summary>
        /// Gets or sets information about named properties used in this action.
        /// </summary>
        public NamedPropertyInfo NamedPropertyInformation
        {
            get { return this.namedPropertyInformation; }
            set { this.namedPropertyInformation = value; }
        }

        /// <summary>
        /// Gets or sets the extended rules Version format. This document defines Version 1, and thus this value MUST be set to 0x00000001.
        /// </summary>
        public uint RuleVersion
        {
            get { return this.ruleVersion; }
            set { this.ruleVersion = value; }
        }

        /// <summary>
        /// Gets or sets a structure containing the actions to be executed when the rule condition for the rule to which these actions apply evaluates to TRUE. The format of this structure is defined in section 2.2.5, using a 4-byte COUNT value.
        /// </summary>
        public RuleAction RuleActionBuffer
        {
            get { return this.ruleActionBuffer; }
            set { this.ruleActionBuffer = value; }
        }

        /// <summary>
        /// Get the total Size of ActionData
        /// </summary>
        /// <returns>The Size of ActionData buffer.</returns>
        public int Size()
        {
            int size = this.NamedPropertyInformation.Size()
                + sizeof(uint)
                + this.RuleActionBuffer.Size();
            return size;
        }

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> bytes = new List<byte>();
            bytes.AddRange(this.NamedPropertyInformation.Serialize());
            bytes.AddRange(BitConverter.GetBytes(this.RuleVersion));
            bytes.AddRange(this.RuleActionBuffer.Serialize());

            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to an ActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            this.NamedPropertyInformation = new NamedPropertyInfo();
            uint namedPropertyInfoLength = this.NamedPropertyInformation.Deserialize(buffer);
            bufferReader = new BufferReader(bufferReader.ReadBytes(namedPropertyInfoLength, (uint)(buffer.Length - namedPropertyInfoLength)));
            this.RuleVersion = bufferReader.ReadUInt32();
            this.RuleActionBuffer = new RuleAction(CountByte.FourBytesCount);
            uint actionLength = this.RuleActionBuffer.Deserialize(bufferReader.ReadToEnd());
            uint length = namedPropertyInfoLength + sizeof(uint) + actionLength;
            return length;
        }
    }
}