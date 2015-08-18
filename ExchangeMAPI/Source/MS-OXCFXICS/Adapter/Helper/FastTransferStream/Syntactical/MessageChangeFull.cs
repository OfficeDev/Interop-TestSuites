namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The messageChangeFull element contains the complete content of 
    /// a new or changed message: the message properties, the recipients,
    /// and the attachments.
    /// messageChangeFull    = IncrSyncChg messageChangeHeader 
    ///                  IncrSyncMessage propList 
    ///               MessageChildren
    /// </summary>
    public class MessageChangeFull : MessageChange
    {
        /// <summary>
        /// A messageChangeHeader value.
        /// </summary>
        private MessageChangeHeader messageChangeHeader;

        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// A MessageChildren value.
        /// </summary>
        private MessageChildren messageChildren;

        /// <summary>
        /// Initializes a new instance of the MessageChangeFull class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public MessageChangeFull(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets LastModificationTime.
        /// </summary>
        public override DateTime LastModificationTime
        {
            get
            {
                return this.MessageChangeHeader.LastModificationTime;
            }
        }

        /// <summary>
        /// Gets LastModificationTime.
        /// </summary>
        public DateTime MessageDeliveryTime
        {
            get
            {
                if (this.PropList != null)
                {
                    ulong t = Convert.ToUInt64(this.PropList.GetPropValue(0x0E06, 0x0040));
                    return DateTime.FromBinary((long)t);
                }

                AdapterHelper.Site.Assert.Fail("The PropList should not be null.");
                return DateTime.Now;
            }
        }

        /// <summary>
        /// Gets a value indicating whether has a rtf body.
        /// </summary>
        public bool IsRTFFormat
        {
            get
            {
                if (this.PropList != null)
                {
                    return this.PropList.HasPropertyTag(0x1009, 0x0102);
                }

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether all subObjects of this are in rtf format.
        /// </summary>
        public bool IsAllRTFFormat
        {
            get
            {
                if (this.PropList != null)
                {
                    bool flag = this.PropList.HasPropertyTag(0x1009, 0x0102);
                    if (flag)
                    {
                        if (this.messageChildren != null
                            && this.messageChildren.Attachments != null
                            && this.messageChildren.Attachments.Count > 0)
                        {
                            foreach (Attachment atta in this.messageChildren.Attachments)
                            {
                                flag = flag && atta.IsRTFFormat;
                                if (!flag)
                                {
                                    return false;
                                }
                            }
                        }
                    }

                    return flag;
                }

                return false;
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
        /// Gets a value indicating the sourceKey property.
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
        /// Gets or sets messageChangeHeader.
        /// </summary>
        public MessageChangeHeader MessageChangeHeader
        {
            get { return this.messageChangeHeader; }
            set { this.messageChangeHeader = value; }
        }

        /// <summary>
        /// Gets or sets propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets or sets MessageChildren.
        /// </summary>
        public MessageChildren MessageChildren
        {
            get { return this.messageChildren; }
            set { this.messageChildren = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized messageChangeFull.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized messageChangeFull, return true, else false.</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncChg);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (stream.ReadMarker(Markers.PidTagIncrSyncChg))
            {
                this.messageChangeHeader = new MessageChangeHeader(stream);
                if (stream.ReadMarker(Markers.PidTagIncrSyncMessage))
                {
                    this.propList = new PropList(stream);
                    this.messageChildren = new MessageChildren(stream);
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}