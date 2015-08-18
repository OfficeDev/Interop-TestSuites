namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The MessageContent element represents the content of a message:
    /// its properties, the recipients, and the attachments.
    /// MessageContent       = propList MessageChildren
    /// </summary>
    public class MessageContent : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Represents children of the Message objects: Recipient and Attachment objects.
        /// </summary>
        private MessageChildren messageChildren;

        /// <summary>
        /// Initializes a new instance of the MessageContent class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public MessageContent(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets the propList.
        /// </summary>
        public PropList PropList
        {
            get
            {
                return this.propList;
            }
        }

        /// <summary>
        /// Gets the MessageChildren.
        /// </summary>
        public MessageChildren MessageChildren
        {
            get
            {
                return this.messageChildren;
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
                        if (this.MessageChildren != null
                                && this.MessageChildren.Attachments != null
                                && this.MessageChildren.Attachments.Count > 0)
                        {
                            foreach (Attachment atta in this.MessageChildren.Attachments)
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
        /// Gets a value indicating whether is a false message.
        /// </summary>
        public bool IsFAIMessage
        {
            get
            {
                if (this.PropList != null)
                {
                    uint val = (uint)this.PropList.GetPropValue(0x0e07, 0x0003);

                    // mfFAI 0x00000040 The message is an FAI message.
                    return 0 != (val & 0x00000040);
                }

                return false;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized MessageContent.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized MessageContent, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return PropList.Verify(stream);
        }

        /// <summary>
        /// Get the corresponding AbstractFastTransferStream.
        /// </summary>
        /// <returns>The corresponding AbstractFastTransferStream.</returns>
        public AbstractFastTransferStream GetAbstractFastTransferStream()
        {
            AbstractFastTransferStream abstractFastTransferStream = new AbstractFastTransferStream
            {
                StreamType = FastTransferStreamType.MessageContent
            };

            AbstractMessageContent abstractMessageContent = new AbstractMessageContent();

            if (this.PropList != null && this.PropList.PropValues.Count > 0)
            {
                // Check the propList of MessageContent if one PropValue is a PtypString value, 
                // the StringPropertiesInUnicode of the AbstractMessageContent is true.
                for (int i = 0; i < this.PropList.PropValues.Count; i++)
                {
                    PropValue p = this.PropList.PropValues[i];
                    if (p.PropType == 0x1f)
                    {
                        abstractMessageContent.StringPropertiesInUnicode = true;
                        break;
                    }
                }

                for (int i = 0; i < this.PropList.PropValues.Count; i++)
                {
                    PropValue p = this.PropList.PropValues[i];

                    // If server stored the string using Unicode, the code page property type should be 0x84B0, which specifies the Unicode (1200) code page, specified in section 2.2.4.1.1.1. 
                    if (p.PropType == 0x84b0)
                    {
                        abstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                        break;
                    }
                }

                for (int i = 0; i < this.PropList.PropValues.Count; i++)
                {
                    PropValue p = this.PropList.PropValues[i];

                    // If server supports other formats (not Unicode), the PropType should not be 0X84b0, which specifies the Unicode (1200) code page, specified in section 2.2.4.1.1.1.
                    if (p.PropType - 0x8000 >= 0 && p.PropType != 0x84b0)
                    {
                        abstractMessageContent.StringPropertiesInOtherCodePage = true;
                        break;
                    }
                }
            }
            else if (this.MessageChildren != null)
            {
                // Check the propList of MessageContent if one PropValue is a PtypString value, 
                // the StringPropertiesInUnicode of the AbstractMessageContent is true.
                foreach (Recipient rec in this.MessageChildren.Recipients)
                {
                    if (rec.PropList.HasPropertyType(0x1f))
                    {
                        abstractMessageContent.StringPropertiesInUnicode = true;
                        break;
                    }

                    // If server stored the string using Unicode, the code page property type should be 0x84B0, which specifies the Unicode (1200) code page, specified in section 2.2.4.1.1.1. 
                    if (rec.PropList.HasPropertyType(0x84b0))
                    {
                        abstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                        break;
                    }

                    for (int i = 0; i < rec.PropList.PropValues.Count; i++)
                    {
                        // If server supports other formats (not Unicode), the PropType should not be 0X84b0, which specifies the Unicode (1200) code page, specified in section 2.2.4.1.1.1.
                        if (rec.PropList.PropValues[i].PropType - 0x8000 >= 0 && rec.PropList.PropValues[i].PropType != 0x84b0)
                        {
                            abstractMessageContent.StringPropertiesInOtherCodePage = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                abstractMessageContent.StringPropertiesInUnicode = false;
            }

            if (this.MessageChildren != null)
            {
                // If MessageChildren contains attachments check whether attachments Preceded By PidTagFXDelProp.
                if (this.MessageChildren.Attachments != null && this.MessageChildren.Attachments.Count > 0)
                {
                    abstractMessageContent.AbsMessageChildren.AttachmentPrecededByPidTagFXDelProp
                         = this.MessageChildren.FXDelPropsBeforeAttachment != null
                         && this.MessageChildren.FXDelPropsBeforeAttachment.Count > 0;
                }

                if (this.MessageChildren.Recipients != null && this.MessageChildren.Recipients.Count > 0)
                {
                    abstractMessageContent.AbsMessageChildren.RecipientPrecededByPidTagFXDelProp
                        = this.MessageChildren.FXDelPropsBeforeRecipient != null
                        && this.MessageChildren.FXDelPropsBeforeRecipient.Count > 0;
                }
            }

            abstractFastTransferStream.AbstractMessageContent = abstractMessageContent;
            return abstractFastTransferStream;
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.propList = new PropList(stream);
            this.messageChildren = new MessageChildren(stream);
        }
    }
}