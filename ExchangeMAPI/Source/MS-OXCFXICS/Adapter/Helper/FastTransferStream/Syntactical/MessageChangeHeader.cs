//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The MessageChangeHeader element contains a fixed set of information about 
    /// the message change that follows this element in the FastTransfer stream. 
    /// The information in the header is sufficient for message identification 
    /// and conflict detection.
    /// MessageChangeHeader  = propList
    /// </summary>
    public class MessageChangeHeader : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the MessageChangeHeader class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public MessageChangeHeader(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets LastModificationTime.
        /// </summary>
        public DateTime LastModificationTime
        {
            get
            {
                ulong value = (ulong)this.PropList.GetPropValue(0x3008, 0x0040);
                return DateTime.FromBinary((long)value);
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagMid.
        /// </summary>
        public bool HasPidTagMid
        {
            get
            {
                return this.PropList.HasPropertyTag(0x674A, 0x0014);
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagMessageSize.
        /// </summary>
        public bool HasPidTagMessageSize
        {
            get
            {
                return this.PropList.HasPropertyTag(0x0E08, 0x0003);
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagChangeNumber.
        /// </summary>
        public bool HasPidTagChangeNumber
        {
            get
            {
                return this.PropList.HasPropertyTag(0x67A4, 0x0014);
            }
        }

        /// <summary>
        /// Gets the sourceKey property.
        /// </summary>
        public byte[] SourceKey
        {
            get
            {
                return this.PropList.GetPropValue(0x65E0, 0x0102) as byte[];
            }
        }

        /// <summary>
        ///  Gets the PidTagChangKey property.
        /// </summary>
        public byte[] PidTagChangeKey
        {
            get
            {
                return this.PropList.GetPropValue(0x65E2, 0x0102) as byte[];
            }
        }

        /// <summary>
        /// Gets the PidTagMid property.
        /// </summary>
        public long PidTagMid
        {
            get
            {
                return (long)this.PropList.GetPropValue(0x674A, 0x0014);
            }
        }

        /// <summary>
        /// Gets the PidTagChangeNumber property.
        /// </summary>
        public long PidTagChangeNumber
        {
            get
            {
                return (long)this.PropList.GetPropValue(0x67a4, 0x0014);
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized MessageChangeHeader.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized MessageChangeHeader, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return PropList.Verify(stream);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.propList = new PropList(stream);
        }
    }
}