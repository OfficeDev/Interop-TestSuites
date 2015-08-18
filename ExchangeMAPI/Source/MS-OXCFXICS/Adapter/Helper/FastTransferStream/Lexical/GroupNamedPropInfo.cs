namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The GroupNamedPropInfo
    /// </summary>
    public class GroupNamedPropInfo : LexicalBase
    {
        /// <summary>
        /// The property set in lexical definition.
        /// </summary>
        private Guid propertySet;

        /// <summary>
        /// The value of flag.
        /// </summary>
        private uint flag;

        /// <summary>
        /// Initializes a new instance of the GroupNamedPropInfo class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public GroupNamedPropInfo(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the  propertySet.
        /// </summary>
        public Guid PropertySet
        {
            get { return this.propertySet; }
            set { this.propertySet = value; }
        }

        /// <summary>
        /// Gets or sets flag.
        /// </summary>
        public uint Flag
        {
            get { return this.flag; }
            set { this.flag = value; }
        }

        /// <summary>
        /// Deserialize a GroupNamedPropInfo instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>A GroupNamedPropInfo instance.</returns>
        public static LexicalBase DeserializeFrom(FastTransferStream stream)
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
        /// <param name="stream">A FastTransferStream</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            this.propertySet = stream.ReadGuid();
            this.flag = stream.ReadUInt32();
        }
    }
}