namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Represents a NamedPropInfo has a dispid.
    /// </summary>
    public class DispidNamedPropInfo : NamedPropInfo
    {
        /// <summary>
        /// The dispid in lexical definition.
        /// </summary>
        private int dispid;

        /// <summary>
        /// Initializes a new instance of the DispidNamedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public DispidNamedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the dispid.
        /// </summary>
        public int Dispid
        {
            get { return this.dispid; }
            set { this.dispid = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized DispidNamedPropInfo.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>If the stream's current position contains 
        /// a serialized DispidNamedPropInfo, return true, else false</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.Verify(0x00, Guid.Empty.ToByteArray().Length);
        }

        /// <summary>
        /// Deserialize a DispidNamedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A DispidNamedPropInfo instance.</returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new DispidNamedPropInfo(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.dispid = stream.ReadInt32(); 
        }
    }
}