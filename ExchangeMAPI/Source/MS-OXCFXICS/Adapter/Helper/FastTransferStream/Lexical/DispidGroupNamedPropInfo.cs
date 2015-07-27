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
    /// A GroupNamedPropInfo has a dispid.
    /// </summary>
    public class DispidGroupNamedPropInfo : GroupNamedPropInfo
    {
        /// <summary>
        /// The dispid.
        /// </summary>
        private uint dispid;

        /// <summary>
        /// Initializes a new instance of the DispidGroupNamedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public DispidGroupNamedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the dispid.
        /// </summary>
        public uint Dispid
        {
            get { return this.dispid; }
            set { this.dispid = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized DispidGroupNamedPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>If the stream's current position contains 
        /// a serialized DispidGroupNamedPropInfo, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyUInt32(Guid.Empty.ToByteArray().Length) ==
                0x00000000;
        }

        /// <summary>
        /// Deserialize a DispidGroupNamedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>A DispidGroupNamedPropInfo instance </returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new DispidGroupNamedPropInfo(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.dispid = stream.ReadUInt32();
        }
    }
}