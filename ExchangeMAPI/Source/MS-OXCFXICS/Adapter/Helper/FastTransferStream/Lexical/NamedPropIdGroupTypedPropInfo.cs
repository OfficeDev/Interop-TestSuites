namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The NamedPropIdGroupTypedPropInfo class.
    /// </summary>
    public class NamedPropIdGroupTypedPropInfo : GroupTypedPropInfo
    {
        /// <summary>
        /// Initializes a new instance of the NamedPropIdGroupTypedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public NamedPropIdGroupTypedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the NamedPropInfo.
        /// </summary>
        public NameGroupNamedPropInfo NamedPropInfo
        {
            get;
            set;
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized NamedPropIdGroupTypedPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized NamedPropIdGroupTypedPropInfo, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return LexicalTypeHelper.IsNamedPropertyID(stream.VerifyUInt16());
        }

        /// <summary>
        /// Deserialize a NamedPropIdGroupTypedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A NamedPropIdGroupTypedPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            if (DispidGroupNamedPropInfo.Verify(stream))
            {
                return DispidGroupNamedPropInfo.DeserializeFrom(stream);
            }
            else if (NameGroupNamedPropInfo.Verify(stream))
            {
                return NameGroupNamedPropInfo.DeserializeFrom(stream);
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
            this.NamedPropInfo = NameGroupNamedPropInfo.DeserializeFrom(stream)
                as NameGroupNamedPropInfo;
        }
    }
}