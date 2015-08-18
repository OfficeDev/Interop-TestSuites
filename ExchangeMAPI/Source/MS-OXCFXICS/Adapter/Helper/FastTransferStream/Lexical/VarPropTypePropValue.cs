namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The VarPropTypePropValue class.
    /// </summary>
    public class VarPropTypePropValue : PropValue
    {
        /// <summary>
        /// The length of a variate type value.
        /// </summary>
        private int length;

        /// <summary>
        /// The valueArray.
        /// </summary>
        private byte[] valueArray;

        /// <summary>
        /// Initializes a new instance of the VarPropTypePropValue class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public VarPropTypePropValue(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the length.
        /// </summary>
        public int Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        /// <summary>
        /// Gets or sets the value array.
        /// </summary>
        public byte[] ValueArray
        {
            get { return this.valueArray; }
            set { this.valueArray = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized VarPropTypePropValue.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized VarPropTypePropValue, return true, else false</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            ushort tmp = stream.VerifyUInt16();
            return LexicalTypeHelper.IsVarType((PropertyDataType)tmp)
                || PropValue.IsPidTagIdsetGiven(stream)
                || (tmp & 0x8000) == 0x8000;  // Adapter all code page specified in [MS-OXCFXICS]2.2.4.1.1.1 Code Page Property Types
        }

        /// <summary>
        /// Deserialize a VarPropTypePropValue instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A VarPropTypePropValue instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new VarPropTypePropValue(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.length = stream.ReadInt32();
            this.valueArray = new byte[this.length];
            stream.Read(this.valueArray, 0, this.valueArray.Length);
        }
    }
}