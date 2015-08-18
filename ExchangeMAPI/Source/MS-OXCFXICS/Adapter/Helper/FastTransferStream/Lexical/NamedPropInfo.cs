namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The NamedPropInfo class.
    /// </summary>
    public class NamedPropInfo : LexicalBase
    {
        /// <summary>
        /// The propertySet item in lexical definition.
        /// </summary>
        private Guid propertySet;

        /// <summary>
        /// The flag variable.
        /// </summary>
        private byte flag;

        /// <summary>
        /// Initializes a new instance of the NamedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public NamedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propertySet.
        /// </summary>
        public Guid PropertySet
        {
            get { return this.propertySet; }
            set { this.propertySet = value; }
        }

        /// <summary>
        /// Gets or sets the flag.
        /// </summary>
        public byte Flag
        {
            get { return this.flag; }
            set { this.flag = value; }
        }

        /// <summary>
        /// Deserialize a NamedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A NamedPropInfo instance.</returns>
        public static LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            if (DispidNamedPropInfo.Verify(stream))
            {
                return new DispidNamedPropInfo(stream);
            }
            else if (NameNamedPropInfo.Verify(stream))
            {
                return new NameNamedPropInfo(stream);
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
            byte[] buffer = new byte[Guid.Empty.ToByteArray().Length];
            stream.Read(buffer, 0, buffer.Length);
            this.propertySet = new Guid(buffer);
            int tmp = stream.ReadByte();
            if (tmp == -1)
            {
                AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
            }
            else if (tmp > 0)
            {
                this.flag = (byte)tmp;
            }
        }
    }
}