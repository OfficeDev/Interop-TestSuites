namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The PropValue element represents 
    /// identification information and the value of the property.
    /// </summary>
    public abstract class PropValue : LexicalBase
    {
        /// <summary>
        /// The propType.
        /// </summary>
        private ushort propType;

        /// <summary>
        /// Initializes a new instance of the PropValue class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public PropValue(FastTransferStream stream)
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
        /// Gets or sets the PropInfo.
        /// </summary>
        public PropInfo PropInfo
        {
            get;
            set;
        }

        /// <summary>
        /// Indicate whether the stream's position is IsPidTagIdsetGiven.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>True if the stream's position is IsPidTagIdsetGiven,
        /// else false.
        /// </returns>
        public static bool IsPidTagIdsetGiven(FastTransferStream stream)
        {
            ushort type = stream.VerifyUInt16();
            ushort id = stream.VerifyUInt16(2);
            return type == (ushort)PropertyDataType.PtypInteger32
                && id == 0x4017;
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized PropValue.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized PropValue, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return !stream.IsEndOfStream
                && (FixedPropTypePropValue.Verify(stream)
                || VarPropTypePropValue.Verify(stream)
                || MvPropTypePropValue.Verify(stream));
        }

        /// <summary>
        /// Deserialize a PropValue instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A PropValue instance.</returns>
        public static LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            if (FixedPropTypePropValue.Verify(stream))
            {
                return FixedPropTypePropValue.DeserializeFrom(stream);
            }
            else if (VarPropTypePropValue.Verify(stream))
            {
                return VarPropTypePropValue.DeserializeFrom(stream);
            }
            else if (MvPropTypePropValue.Verify(stream))
            {
                return MvPropTypePropValue.DeserializeFrom(stream);
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
            PropInfo = PropInfo.DeserializeFrom(stream) as PropInfo;
        }
    }
}