//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The attachmentContent element contains the properties and the embedded message
    /// of an Attachment object. If present,
    /// attachmentContent = propList [EmbeddedMessage]
    /// </summary>
    public class AttachmentContent : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// An EmbeddedMessage value.
        /// </summary>
        private EmbeddedMessage embeddedMessage;

        /// <summary>
        /// Initializes a new instance of the AttachmentContent class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public AttachmentContent(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets or sets the EmbeddedMessage.
        /// </summary>
        public EmbeddedMessage EmbeddedMessage
        {
            get { return this.embeddedMessage; }
            set { this.embeddedMessage = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized attachmentContent.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized attachmentContent, return true, else false.</returns>
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
                StreamType = FastTransferStreamType.attachmentContent
            };

            return abstractFastTransferStream;
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.propList = new PropList(stream);
            if (EmbeddedMessage.Verify(stream))
            {
                this.embeddedMessage = new EmbeddedMessage(stream);
            }
        }
    }
}