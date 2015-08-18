namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// The MessageList element contains a list of messages, 
    /// which is determined by the scope of the operation.
    /// MessageList          = 1*( [PidTagEcWarning] message )
    /// </summary>
    public class MessageList : SyntacticalBase
    {
        /// <summary>
        /// A list of message objects.
        /// </summary>
        private List<Message> messages;

        /// <summary>
        /// A list of uint32 values. Each represents an error code.
        /// </summary>
        private List<uint> errorCodeList;

        /// <summary>
        /// Initializes a new instance of the MessageList class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public MessageList(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the error code list.
        /// </summary>
        public List<uint> ErrorCodeList
        {
            get { return this.errorCodeList; }
            set { this.errorCodeList = value; }
        }

        /// <summary>
        /// Gets or sets  the message list.
        /// </summary>
        public List<Message> Messages
        {
            get { return this.messages; }
            set { this.messages = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized MessageList.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized MessageList, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return (!stream.IsEndOfStream
                && stream.VerifyUInt32() == (uint)MetaProperties.PidTagEcWarning)
                || Message.Verify(stream);
        }

        /// <summary>
        /// Get the corresponding AbstractFastTransferStream.
        /// </summary>
        /// <returns>The corresponding AbstractFastTransferStream.</returns>
        public AbstractFastTransferStream GetAbstractFastTransferStream()
        {
            AbstractFastTransferStream abstractFastTransferStream = new AbstractFastTransferStream
            {
                StreamType = FastTransferStreamType.MessageList
            };
            AbstractMessageList abstractMessageList = new AbstractMessageList
            {
                IsPidTagEcWarningOut = this.ErrorCodeList.Count != 0
            };

            // If ErrorCodeList contains values, it means PidTagEcWaring is out in model level.
            // Since ErrorCodeList contains PidTagEcWaring values, each value is after a PidTagEcWaring.

            // Beside checking permission, each MessageList contains at least 1 message.
            abstractMessageList.AbsMessage.AbsMessageContent.IsNoPermissionMessageNotOut = !(this.Messages != null && this.Messages.Count > 0);
            abstractFastTransferStream.AbstractMessageList = abstractMessageList;
            return abstractFastTransferStream;
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            int count = 0;
            this.messages = new List<Message>();
            this.errorCodeList = new List<uint>();
            do
            {
                if (stream.VerifyMetaProperty(MetaProperties.PidTagEcWarning))
                {
                    stream.ReadMarker();
                    this.errorCodeList.Add(stream.ReadUInt32());
                    return;
                }

                if (Message.Verify(stream))
                {
                    this.messages.Add(new Message(stream));
                }
                else if (count > 0)
                {
                    return;
                }
                else
                {
                    AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
                }

                count++;
            } 
            while (Verify(stream));
        }
    }
}