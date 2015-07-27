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
    /// The message element represents a Message object.
    /// message              = ( StartMessage / StartFAIMsg ) 
    ///          MessageContent 
    ///          EndMessage
    /// </summary>
    public class Message : SyntacticalBase
    {
        /// <summary>
        /// The start marker of message.
        /// </summary>
        public const Markers StartMarker1 = Markers.PidTagStartMessage;

        /// <summary>
        /// The start marker of message.
        /// </summary>
        public const Markers StartMarker2 = Markers.PidTagStartFAIMsg;

        /// <summary>
        /// The end marker of message.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagEndMessage;

        /// <summary>
        /// A MessageContent value.
        /// Represents the content of a message: 
        /// its properties, the recipients, and the attachments.
        /// </summary>
        private MessageContent content;

        /// <summary>
        /// Initializes a new instance of the Message class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public Message(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets MessageContent.
        /// </summary>
        public MessageContent Content
        {
            get { return this.content; }
            set { this.content = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized message.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized message, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(StartMarker1) ||
                stream.VerifyMarker(StartMarker2);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            Markers marker = stream.ReadMarker();
            if (marker == Markers.PidTagStartMessage
                || marker == Markers.PidTagStartFAIMsg)
            {
                this.content = new MessageContent(stream);
                if (stream.ReadMarker(Markers.PidTagEndMessage))
                {
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}