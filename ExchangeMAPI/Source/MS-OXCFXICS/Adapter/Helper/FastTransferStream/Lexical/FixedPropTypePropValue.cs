namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Represent a fixedPropType PropValue.
    /// </summary>
    public class FixedPropTypePropValue : PropValue
    {
        /// <summary>
        /// A fixed value.
        /// </summary>
        private object fixedValue;

        /// <summary>
        /// Initializes a new instance of the FixedPropTypePropValue class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public FixedPropTypePropValue(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the fixedValue.
        /// </summary>
        public object FixedValue
        {
            get { return this.fixedValue; }
            set { this.fixedValue = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized FixedPropTypePropValue.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized FixedPropTypePropValue, return true, else false</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            ushort tmp = stream.VerifyUInt16();
            return LexicalTypeHelper.IsFixedType((PropertyDataType)tmp)
                && !PropValue.IsPidTagIdsetGiven(stream);
        }

        /// <summary>
        /// Deserialize a DispidNamedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A DispidNamedPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new FixedPropTypePropValue(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            PropertyDataType type = (PropertyDataType)this.PropType;

            switch (type)
            {
                case PropertyDataType.PtypInteger16:
                    this.fixedValue = stream.ReadInt16();
                    break;
                case PropertyDataType.PtypInteger32:
                    this.fixedValue = stream.ReadInt32();
                    break;
                case PropertyDataType.PtypFloating32:
                    this.fixedValue = stream.ReadFloating32();
                    break;
                case PropertyDataType.PtypFloating64:
                    this.fixedValue = stream.ReadFloating64();
                    break;
                case PropertyDataType.PtypCurrency:
                    this.fixedValue = stream.ReadCurrency();
                    break;
                case PropertyDataType.PtypFloatingTime:
                    this.fixedValue = stream.ReadFloatingTime();
                    break;
                case PropertyDataType.PtypBoolean:
                    this.fixedValue = stream.ReadBoolean();
                    break;
                case PropertyDataType.PtypInteger64:
                    this.fixedValue = stream.ReadInt64();
                    break;
                case PropertyDataType.PtypTime:
                    this.fixedValue = stream.ReadTime();
                    break;
                case PropertyDataType.PtypGuid:
                    this.fixedValue = stream.ReadGuid();
                    break;
            }
        }
    }
}