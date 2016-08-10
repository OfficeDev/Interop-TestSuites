namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The fields TableRowDataSize,TotalMessageCount,UnreadMessageCount,FolderIDNumber,
    /// TableRowFolderID,TableRowMessageID,TableRowPreviousInstance,TableRowOldFolderID,TableRowOldMessageID
    /// </summary>
    internal class AvailableFieldsType
    {
        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowDataSize is available or not
        /// </summary>
        public bool IsTableRowDataSizeAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TotalMessageCount is available or not
        /// </summary>
        public bool IsTotalMessageCountAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field UnreadMessageCount is available or not
        /// </summary>
        public bool IsUnreadMessageCountAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field FolderIDNumber is available or not
        /// </summary>
        public bool IsFolderIDNumberAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowFolderID is available or not
        /// </summary>
        public bool IsTableRowFolderIDAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowMessageID is available or not
        /// </summary>
        public bool IsTableRowMessageIDAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowPreviousInstance is available or not
        /// </summary>
        public bool IsTableRowPreviousInstanceAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowOldFolderID is available or not
        /// </summary>
        public bool IsTableRowOldFolderIDAvailable { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field TableRowOldMessageID is available or not
        /// </summary>
        public bool IsTableRowOldMessageIDAvailable { get; set; }
    }
}
