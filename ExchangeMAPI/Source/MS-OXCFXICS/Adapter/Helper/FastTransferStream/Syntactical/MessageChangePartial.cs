namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The MessageChangePartial element represents the difference 
    /// in message content since the last download, as identified by 
    /// the initial ICS state.
    /// MessageChangePartial = [groupInfo] [PidTagIncrSyncGroupId]
    ///                 IncrSyncChgPartial messageChangeHeader
    ///                 *( PidTagIncrementalSyncMessagePartial propList )
    ///                 MessageChildren
    /// </summary>
    public class MessageChangePartial : MessageChange
    {
        /// <summary>
        /// A groupInfo value.
        /// Provides a definition for the property group mapping.
        /// </summary>
        private GroupInfo groupInfo;

        /// <summary>
        /// A messageChangeHeader value.
        /// Contains a fixed set of information about the message change that follows this element in the FastTransfer stream.
        /// </summary>
        private MessageChangeHeader messageChangeHeader;

        /// <summary>
        /// A list of  propList values.
        /// </summary>
        private List<PropList> propListList;

        /// <summary>
        /// A uint value after PidTagIncrSyncGroupId.
        /// Specifies an identifier of a property group mapping.
        /// </summary>
        private uint? incrSyncGroupId;

        /// <summary>
        /// A uint value after PidTagIncrementalSyncMessagePartial.
        /// Specifies an index of a property group within 
        /// a property group mapping currently in context.
        /// </summary>
        private uint? incrementalSyncMessagePartial;

        /// <summary>
        /// Initializes a new instance of the MessageChangePartial class.
        /// </summary>
        /// <param name="stream">A FastTransferStream object.</param>
        public MessageChangePartial(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets the last modification time in the messageChangeHeader.
        /// </summary>
        public override DateTime LastModificationTime
        {
            get
            {
                return this.MessageChangeHeader.LastModificationTime;
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagMid.
        /// </summary>
        public override bool HasPidTagMid
        {
            get
            {
                return this.MessageChangeHeader.HasPidTagMid;
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagMessageSize.
        /// </summary>
        public override bool HasPidTagMessageSize
        {
            get
            {
                return this.MessageChangeHeader.HasPidTagMessageSize;
            }
        }

        /// <summary>
        /// Gets a value indicating whether has PidTagChangeNumber.
        /// </summary>
        public override bool HasPidTagChangeNumber
        {
            get
            {
                return this.MessageChangeHeader.HasPidTagChangeNumber;
            }
        }

        /// <summary>
        /// Gets the sourceKey property.
        /// </summary>
        public override byte[] SourceKey
        {
            get
            {
                return this.MessageChangeHeader.SourceKey;
            }
        }

        /// <summary>
        /// Gets a value indicating the PidTagChangeKey.
        /// </summary>
        public override byte[] PidTagChangeKey
        {
            get
            {
                return this.messageChangeHeader.PidTagChangeKey;
            }
        }

        /// <summary>
        /// Gets a value indicating the PidTagChangeNumber.
        /// </summary>
        public override long PidTagChangeNumber
        {
            get
            {
                return this.messageChangeHeader.PidTagChangeNumber;
            }
        }

        /// <summary>
        /// Gets a value indicating the PidTagMid.
        /// </summary>
        public override long PidTagMid
        {
            get { return this.messageChangeHeader.PidTagMid; }
        }

        /// <summary>
        /// Gets or sets incrementalSyncMessagePartial.
        /// </summary>
        public uint? IncrementalSyncMessagePartial
        {
            get { return this.incrementalSyncMessagePartial; }
            set { this.incrementalSyncMessagePartial = value; }
        }

        /// <summary>
        /// Gets or sets incrSyncGroupId.
        /// </summary>
        public uint? IncrSyncGroupId
        {
            get { return this.incrSyncGroupId; }
            set { this.incrSyncGroupId = value; }
        }

        /// <summary>
        /// Gets or sets groupInfo.
        /// </summary>
        public GroupInfo GroupInfo
        {
            get { return this.groupInfo; }
            set { this.groupInfo = value; }
        }

        /// <summary>
        /// Gets or sets propList list.
        /// </summary>
        public List<PropList> PropListList
        {
            get { return this.propListList; }
            set { this.propListList = value; }
        }

        /// <summary>
        /// Gets or sets messageChangeHeader.
        /// </summary>
        public MessageChangeHeader MessageChangeHeader
        {
            get { return this.messageChangeHeader; }
            set { this.messageChangeHeader = value; }
        }

        /// <summary>
        /// Gets or sets MessageChildren.
        /// </summary>
        public MessageChildren MessageChildren
        {
            get;
            set;
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized MessageChangePartial.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized MessageChangePartial, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return GroupInfo.Verify(stream)
                || stream.VerifyMetaProperty(MetaProperties.MetaTagIncrSyncGroupId)
                || stream.VerifyMarker(Markers.PidTagIncrSyncChgPartial);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (GroupInfo.Verify(stream))
            {
                this.groupInfo = new GroupInfo(stream);
            }

            if (stream.VerifyMetaProperty(MetaProperties.MetaTagIncrSyncGroupId))
            {
                stream.ReadMarker();
                this.incrSyncGroupId = stream.ReadUInt32();
            }

            if (stream.ReadMarker(Markers.PidTagIncrSyncChgPartial))
            {
                this.messageChangeHeader = new MessageChangeHeader(stream);
                this.propListList = new List<PropList>();
                while (stream.VerifyMetaProperty(MetaProperties.MetaTagIncrementalSyncMessagePartial))
                {
                    stream.ReadMarker();
                    this.incrementalSyncMessagePartial = stream.ReadUInt32();
                    this.propListList.Add(new PropList(stream));
                }

                this.MessageChildren = new MessageChildren(stream);
                return;
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}