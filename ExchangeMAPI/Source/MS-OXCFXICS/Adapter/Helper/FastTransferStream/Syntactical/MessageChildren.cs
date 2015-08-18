namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// The MessageChildren element represents children of the Message objects: 
    /// Recipient and Attachment objects.
    ///  MessageChildren = [ PidTagFXDelProp ] [ *Recipient ]
    ///                     [ PidTagFXDelProp ] [ *attachment ]
    /// </summary>
    public class MessageChildren : SyntacticalBase
    {
        /// <summary>
        /// A list of recipients.
        /// </summary>
        private List<Recipient> recipients;

        /// <summary>
        /// A list of attachments.
        /// </summary>
        private List<Attachment> attachments;

        /// <summary>
        /// A list of FXDelProp values.
        /// </summary>
        private List<uint> fxdelPropsBeforeRecipient;

        /// <summary>
        /// A list of FXDelProp values.
        /// </summary>
        private List<uint> fxdelPropsBeforeAttachment;

        /// <summary>
        /// A list of all the FXDelProp.
        /// </summary>
        private List<uint> fxdelPropsAll;

        /// <summary>
        /// Initializes a new instance of the MessageChildren class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public MessageChildren(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets fxdelPropsBeforeRecipient.
        /// </summary>
        public List<uint> FXDelPropsBeforeRecipient
        {
            get { return this.fxdelPropsBeforeRecipient; }
            set { this.fxdelPropsBeforeRecipient = value; }
        }

        /// <summary>
        /// Gets or sets fxdelPropsBeforeAttachment.
        /// </summary>
        public List<uint> FXDelPropsBeforeAttachment
        {
            get { return this.fxdelPropsBeforeAttachment; }
            set { this.fxdelPropsBeforeAttachment = value; }
        }

        /// <summary>
        /// Gets or sets FXDelProps.
        /// </summary>
        public List<uint> FXDelProps
        {
            get { return this.fxdelPropsAll; }
            set { this.fxdelPropsAll = value; }
        }

        /// <summary>
        /// Gets or sets the attachment list.
        /// </summary>
        public List<Attachment> Attachments
        {
            get { return this.attachments; }
            set { this.attachments = value; }
        }

        /// <summary>
        /// Gets or sets the Recipient list.
        /// </summary>
        public List<Recipient> Recipients
        {
            get { return this.recipients; }
            set { this.recipients = value; }
        }

        /// <summary>
        /// Gets a value indicating whether all subObjects of this are in rtf format.
        /// </summary>
        public bool IsAllRTFFormat
        {
            get
            {
                if (this.Attachments != null && this.Attachments.Count > 0)
                {
                    bool flag = this.Attachments[0].IsRTFFormat;
                    for (int i = 1; i < this.Attachments.Count && flag; i++)
                    {
                        flag = flag && this.Attachments[i].IsRTFFormat;
                    }

                    return flag;
                }

                return false;
            }
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.fxdelPropsAll = new List<uint>();
            this.fxdelPropsBeforeRecipient = new List<uint>();
            this.fxdelPropsBeforeAttachment = new List<uint>();
            this.attachments = new List<Attachment>();
            this.recipients = new List<Recipient>();
            while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp))
            {
                stream.ReadMarker();
                this.fxdelPropsBeforeRecipient.Add(stream.ReadUInt32());
            }

            if (Recipient.Verify(stream))
            {
                this.recipients = new List<Recipient>();
                while (Recipient.Verify(stream))
                {
                    this.recipients.Add(new Recipient(stream));
                }
            }

            while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp))
            {
                stream.ReadMarker();
                this.fxdelPropsBeforeAttachment.Add(stream.ReadUInt32());
            }

            while (Attachment.Verify(stream))
            {
                this.attachments.Add(new Attachment(stream));
            }

            for (int i = 0; i < this.fxdelPropsBeforeRecipient.Count; i++)
            {
                this.fxdelPropsAll.Add(this.fxdelPropsBeforeRecipient[i]);
            }

            for (int i = 0; i < this.fxdelPropsBeforeAttachment.Count; i++)
            {
                this.fxdelPropsAll.Add(this.fxdelPropsBeforeAttachment[i]);
            }
        }
    }
}