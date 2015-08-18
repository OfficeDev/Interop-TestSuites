namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The TaggedPropIdPropInfo class.
    /// </summary>
    public class TaggedPropIdPropInfo : PropInfo
    {
        /// <summary>
        /// Initializes a new instance of the TaggedPropIdPropInfo class.
        /// </summary>
        /// <param name="stream">a FastTransferStream</param>
        public TaggedPropIdPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized TaggedPropIdPropInfo
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized TaggedPropIdPropInfo, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return LexicalTypeHelper.IsTaggedPropertyID(stream.VerifyUInt16());
        }

        /// <summary>
        /// Deserialize a TaggedPropIdPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A TaggedPropIdPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new TaggedPropIdPropInfo(stream);
        }
    }
}