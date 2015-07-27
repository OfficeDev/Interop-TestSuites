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
    /// The TaggedPropIdGroupTypedPropInfo class.
    /// </summary>
    public class TaggedPropIdGroupTypedPropInfo : GroupTypedPropInfo
    {
        /// <summary>
        /// Initializes a new instance of the TaggedPropIdGroupTypedPropInfo class.
        /// </summary>
        /// <param name="stream">a FastTransferStream</param>
        public TaggedPropIdGroupTypedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized TaggedPropIdGroupTypedPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized TaggedPropIdGroupTypedPropInfo, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return LexicalTypeHelper.IsTaggedPropertyID(stream.VerifyUInt16(2));
        }

        /// <summary>
        /// Deserialize a TaggedPropIdGroupTypedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A TaggedPropIdGroupTypedPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new TaggedPropIdGroupTypedPropInfo(stream);
        }
    }
}