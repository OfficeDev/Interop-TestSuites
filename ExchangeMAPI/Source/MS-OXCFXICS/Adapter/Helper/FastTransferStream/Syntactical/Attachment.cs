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
    /// Contains an attachmentContent.
    /// Attachment = NewAttach attachmentContent EndAttach
    /// </summary>
    public class Attachment : SyntacticalBase
    {
        /// <summary>
        /// The  start marker of an attachment object.
        /// </summary>
        public const Markers StartMarker = Markers.PidTagNewAttach;

        /// <summary>
        /// The end marker of an attachment object.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagEndAttach;

        /// <summary>
        /// Attachment content.
        /// </summary>
        private AttachmentContent attachmentContent;

        /// <summary>
        /// Initializes a new instance of the Attachment class.
        /// </summary>
        /// <param name="stream">a FastTransferStream</param>
        public Attachment(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the attachmentContent.
        /// </summary>
        public AttachmentContent AttachmentContent
        {
            get { return this.attachmentContent; }
            set { this.attachmentContent = value; }
        }

        /// <summary>
        /// Gets a value indicating whether there is an rtf body
        /// propList containing PidTagRtfCompressed.
        /// </summary>
        public bool IsRTFFormat
        {
            get
            {
                if (this.AttachmentContent != null
                    && this.AttachmentContent.PropList != null)
                {
                    return this.AttachmentContent.PropList.HasPropertyTag(0x1009, 0x0102);
                }

                return false;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized attachment.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized attachment, return true, else false.</returns>
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
            if (stream.ReadMarker(StartMarker))
            {
                this.attachmentContent = new AttachmentContent(stream);
                if (stream.ReadMarker(EndMarker))
                {
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}