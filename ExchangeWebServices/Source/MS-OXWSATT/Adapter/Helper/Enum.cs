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
        PostAttachment,

        /// <summary>
        ///  Identifies the attachment is a MeetingMessage type item attachment.
        /// </summary>
        MeetingMessageAttachemnt,

        /// <summary>
        /// Identifies the attachment is a MeetingRequest type item attachment.
        /// </summary>
        MeetingRequestAttachment,

        /// <summary>
        /// Identifies the attachment is a MeetingResponse type item attachment.
        /// </summary>
        MeetingResponseAttachment,

        /// <summary>
        /// Identifies the attachment is a MeetingCancellation type item attachment.
        /// </summary>
        MeetingCancellationAttachment,

        /// <summary>
        /// Identifies the attachment is a Person type item attachment.
        /// </summary>
        PersonAttachment,

        /// <summary>
        /// Identifies the attachment is a reference attachment.
        /// </summary>
        ReferenceAttachment,

        /// <summary>
        /// Identifies the none clild element in item attachment.
        /// </summary>
        NoneChildAttachment,
    }
}