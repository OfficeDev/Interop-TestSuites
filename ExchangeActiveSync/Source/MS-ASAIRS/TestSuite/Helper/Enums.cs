namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    /// <summary>
    /// The type of email.
    /// </summary>
    public enum EmailType
    {
        /// <summary>
        /// The email type is plain text
        /// </summary>
        Plaintext,

        /// <summary>
        /// The email type is html
        /// </summary>
        HTML,

        /// <summary>
        /// The email attachment is a normal attachment
        /// </summary>
        NormalAttachment,

        /// <summary>
        /// The email attachment is an e-mail message
        /// </summary>
        EmbeddedAttachment,

        /// <summary>
        /// The email attachment is an embedded Object Linking and Embedding (OLE) object
        /// </summary>
        AttachOLE
    }
}