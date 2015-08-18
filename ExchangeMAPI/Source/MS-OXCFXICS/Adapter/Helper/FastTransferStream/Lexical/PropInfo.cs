namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The PropInfo class.
    /// </summary>
    public abstract class PropInfo : LexicalBase
    {
        /// <summary>
        /// The property id.
        /// </summary>
        private ushort propID;

        /// <summary>
        /// Initializes a new instance of the PropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        protected PropInfo(FastTransferStream stream)
            : base(stream)
        {
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
        /// Verify that a stream's current position contains a serialized PropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized PropInfo, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return TaggedPropIdPropInfo.Verify(stream)
                || NamedPropIdPropInfo.Verify(stream);
        }

        /// <summary>
        /// Deserialize a PropInfo instance from a FastTransferStream
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A PropInfo instance.</returns>
        public static LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            if (TaggedPropIdPropInfo.Verify(stream))
            {
                return TaggedPropIdPropInfo.DeserializeFrom(stream);
            }
            else if (NamedPropIdPropInfo.Verify(stream))
            {
                return NamedPropIdPropInfo.DeserializeFrom(stream);
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
            this.propID = stream.ReadUInt16();
        }
    }
}