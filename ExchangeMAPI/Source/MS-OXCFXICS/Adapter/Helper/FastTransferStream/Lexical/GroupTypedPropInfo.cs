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
    /// GroupTypedPropInfo class.
    /// </summary>
    public class GroupTypedPropInfo : LexicalBase
    {
        /// <summary>
        /// The propID.
        /// </summary>
        private ushort propID;

        /// <summary>
        /// The propType.
        /// </summary>
        private ushort propType;

        /// <summary>
        /// Initializes a new instance of the GroupTypedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public GroupTypedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propType.
        /// </summary>
        public ushort PropType
        {
            get { return this.propType; }
            set { this.propType = value; }
        }

        /// <summary>
        /// Gets or sets the propID.
        /// </summary>
        public ushort PropID
        {
            get { return this.propID; }
            set { this.propID = value; }
        }

        /// <summary>
        /// Verify rhat a stream's current position contains a serialized GroupTypedPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>If the stream's current position contains 
        /// a serialized GroupTypedPropInfo, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            ushort t = stream.VerifyUInt16();
            PropertyDataType type = (PropertyDataType)t;
            return LexicalTypeHelper.IsFixedType(type)
                || LexicalTypeHelper.IsMVType(type)
                || LexicalTypeHelper.IsVarType(type);
        }

        /// <summary>
        /// Deserialize a GroupTypedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A GroupTypedPropInfo instance.</returns>
        public static LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            if (TaggedPropIdGroupTypedPropInfo.Verify(stream))
            {
                return TaggedPropIdGroupTypedPropInfo.DeserializeFrom(stream);
            }
            else if (NamedPropIdGroupTypedPropInfo.Verify(stream))
            {
                return NamedPropIdGroupTypedPropInfo.DeserializeFrom(stream);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
                return null;
            }
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.propType = stream.ReadUInt16();
            this.propID = stream.ReadUInt16();
        }
    }
}