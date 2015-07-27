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
    /// Contain a MessageContent.
    /// EmbeddedMessage      = StartEmbed MessageContent EndEmbed
    /// </summary>
    public class EmbeddedMessage : SyntacticalBase
    {
        /// <summary>
        /// The start marker of the EmbeddedMessage.
        /// </summary>
        public const Markers StartMarker = Markers.PidTagStartEmbed;

        /// <summary>
        /// The end marker of the EmbeddedMessage.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagEndEmbed;

        /// <summary>
        /// A MessageContent value represents the content of a message: its properties, the recipients, and the attachments.
        /// </summary>
        private MessageContent messageContent;

        /// <summary>
        /// Initializes a new instance of the EmbeddedMessage class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public EmbeddedMessage(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the messageContent.
        /// </summary>
        public MessageContent MessageContent
        {
            get { return this.messageContent; }
            set { this.messageContent = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized EmbeddedMessage.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized EmbeddedMessage, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(StartMarker);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.Deserialize<MessageContent>(stream, StartMarker, EndMarker, out this.messageContent);
        }
    }
}