//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    /// <summary>
    /// An enumeration identifies the attachment type.
    /// </summary>
    public enum AttachmentTypeValue
    {
        /// <summary>
        ///  Identifies the attachment is a file attachment.
        /// </summary>
        FileAttachment,

        /// <summary>
        /// Identifies the attachment is an item type attachment.
        /// </summary>
        ItemAttachment,

        /// <summary>
        /// Identifies the attachment is a message type item attachment.
        /// </summary>
        MessageAttachment,

        /// <summary>
        /// Identifies the attachment is a calendar type item attachment.
        /// </summary>
        CalendarAttachment,

        /// <summary>
        /// Identifies the attachment is a task type item attachment.
        /// </summary>
        TaskAttachment,

        /// <summary>
        /// Identifies the attachment is a contact type item attachment.
        /// </summary>
        ContactAttachment,

        /// <summary>
        /// Identifies the attachment is a post type item attachment.
        /// </summary>
        PostAttachment
    }
}