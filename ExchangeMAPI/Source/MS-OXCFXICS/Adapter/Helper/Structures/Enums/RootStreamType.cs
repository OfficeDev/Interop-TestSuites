namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Represents the type of FastTransfer stream.
    /// </summary>
    public enum FastTransferStreamType
    {
        /// <summary>
        /// The contentsSync.
        /// </summary>
        contentsSync = 0,

        /// <summary>
        /// The hierarchySync.
        /// </summary>
        hierarchySync = 1,

        /// <summary>
        /// The current state.
        /// </summary>
        state = 2,

        /// <summary>
        /// The folderContent.
        /// </summary>
        folderContent = 3,

        /// <summary>
        /// The MessageContent.
        /// </summary>
        MessageContent = 4,

        /// <summary>
        /// The attachmentContent.
        /// </summary>
        attachmentContent = 5,

        /// <summary>
        /// The MessageList.
        /// </summary>
        MessageList = 6,

        /// <summary>
        /// The TopFolder.
        /// </summary>
        TopFolder = 7
    }
}