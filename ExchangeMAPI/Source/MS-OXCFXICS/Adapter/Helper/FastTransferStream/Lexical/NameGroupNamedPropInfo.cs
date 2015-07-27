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
    /// <summary>
    /// The NameGroupNamedPropInfo class.
    /// </summary>
    public class NameGroupNamedPropInfo : GroupNamedPropInfo
    {
        /// <summary>
        /// The name of the GroupNamedPropInfo.
        /// </summary>
        private string name;

        /// <summary>
        /// Initializes a new instance of the NameGroupNamedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public NameGroupNamedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.name = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized NameGroupNamedPropInfo
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized NameGroupNamedPropInfo, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return 0x01000000 == stream.VerifyUInt32(LexicalBase.GuidLength);
        }

        /// <summary>
        /// Deserialize a NameGroupNamedPropInfo instance from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A NameGroupNamedPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new NameGroupNamedPropInfo(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.name = stream.ReadString();
        }
    }
}