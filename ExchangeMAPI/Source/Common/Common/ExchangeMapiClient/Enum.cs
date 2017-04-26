namespace Microsoft.Protocols.TestSuites.Common
{
    using System;

    #region RopId definition
    /// <summary>
    /// Each remote operation (ROP) is identified by a one-byte value, which is contained in the RopId field of ROP request buffers and ROP response buffers. 
    /// The ROPs that a client is allowed to use are listed in the following table. A ROP that is specified as "Reserved" is not used in the communication between 
    /// the client and server. Therefore, the client MUST NOT use the reserved ROPs.
    /// </summary>
    public enum RopId : byte
    {
        /// <summary>
        /// Reserved rop id
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Rop Release id
        /// </summary>
        RopRelease = 0x01,

        /// <summary>
        /// Rop OpenFolder id
        /// </summary>
        RopOpenFolder = 0x02,

        /// <summary>
        /// Rop OpenMessage id
        /// </summary>
        RopOpenMessage = 0x03,

        /// <summary>
        /// Rop GetHierarchyTable id
        /// </summary>
        RopGetHierarchyTable = 0x04,

        /// <summary>
        /// Rop GetContentsTable id
        /// </summary>
        RopGetContentsTable = 0x05,

        /// <summary>
        /// Rop CreateMessage id
        /// </summary>
        RopCreateMessage = 0x06,

        /// <summary>
        /// Rop GetPropertiesSpecific id
        /// </summary>
        RopGetPropertiesSpecific = 0x07,

        /// <summary>
        /// Rop GetPropertiesAll id
        /// </summary>
        RopGetPropertiesAll = 0x08,

        /// <summary>
        /// Rop GetPropertiesList id 
        /// </summary>
        RopGetPropertiesList = 0x09,

        /// <summary>
        /// Rop SetProperties id
        /// </summary>
        RopSetProperties = 0x0A,

        /// <summary>
        /// Rop DeleteProperties id
        /// </summary>
        RopDeleteProperties = 0x0B,

        /// <summary>
        /// Rop SaveChangesMessage id
        /// </summary>
        RopSaveChangesMessage = 0x0C,

        /// <summary>
        /// Rop RemoveAllRecipients id
        /// </summary>
        RopRemoveAllRecipients = 0x0D,

        /// <summary>
        /// Rop ModifyRecipients id
        /// </summary>
        RopModifyRecipients = 0x0E,

        /// <summary>
        /// Rop ReadRecipients id
        /// </summary>
        RopReadRecipients = 0x0F,

        /// <summary>
        /// Rop ReloadCachedInformation id
        /// </summary>
        RopReloadCachedInformation = 0x10,

        /// <summary>
        /// Rop SetMessageReadFlag id
        /// </summary>
        RopSetMessageReadFlag = 0x11,

        /// <summary>
        /// Rop SetColumns id
        /// </summary>
        RopSetColumns = 0x12,

        /// <summary>
        /// Rop SortTable id
        /// </summary>
        RopSortTable = 0x13,

        /// <summary>
        /// Rop Restrict id
        /// </summary>
        RopRestrict = 0x14,

        /// <summary>
        /// Rop QueryRows id
        /// </summary>
        RopQueryRows = 0x15,

        /// <summary>
        /// Rop GetStatus id
        /// </summary>
        RopGetStatus = 0x16,

        /// <summary>
        /// Rop QueryPosition id
        /// </summary>
        RopQueryPosition = 0x17,

        /// <summary>
        /// Rop SeekRow id
        /// </summary>
        RopSeekRow = 0x18,

        /// <summary>
        /// Rop SeekRowBookmark id
        /// </summary>
        RopSeekRowBookmark = 0x19,

        /// <summary>
        /// Rop SeekRowFractional id
        /// </summary>
        RopSeekRowFractional = 0x1A,

        /// <summary>
        /// Rop CreateBookmark id
        /// </summary>
        RopCreateBookmark = 0x1B,

        /// <summary>
        /// Rop CreateFolder id
        /// </summary>
        RopCreateFolder = 0x1C,

        /// <summary>
        /// Rop DeleteFolder id
        /// </summary>
        RopDeleteFolder = 0x1D,

        /// <summary>
        /// Rop DeleteMessages id
        /// </summary>
        RopDeleteMessages = 0x1E,

        /// <summary>
        /// Rop GetMessageStatus id
        /// </summary>
        RopGetMessageStatus = 0x1F,

        /// <summary>
        /// Rop SetMessageStatus id
        /// </summary>
        RopSetMessageStatus = 0x20,

        /// <summary>
        /// Rop GetAttachmentTable id
        /// </summary>
        RopGetAttachmentTable = 0x21,

        /// <summary>
        /// Rop GetValidAttachments id
        /// </summary>
        RopGetValidAttachments = 0x52,

        /// <summary>
        /// Rop OpenAttachment id
        /// </summary>
        RopOpenAttachment = 0x22,

        /// <summary>
        /// Rop CreateAttachment id
        /// </summary>
        RopCreateAttachment = 0x23,

        /// <summary>
        /// Rop DeleteAttachment id
        /// </summary>
        RopDeleteAttachment = 0x24,

        /// <summary>
        /// Rop SaveChangesAttachment id
        /// </summary>
        RopSaveChangesAttachment = 0x25,

        /// <summary>
        /// Rop SetReceiveFolder id
        /// </summary>
        RopSetReceiveFolder = 0x26,

        /// <summary>
        /// Rop GetReceiveFolder id
        /// </summary>
        RopGetReceiveFolder = 0x27,

        /// <summary>
        /// Rop RegisterNotification id
        /// </summary>
        RopRegisterNotification = 0x29,

        /// <summary>
        /// Rop Notify id
        /// </summary>
        RopNotify = 0x2A,

        /// <summary>
        /// Rop OpenStream id
        /// </summary>
        RopOpenStream = 0x2B,

        /// <summary>
        /// Rop ReadStream id
        /// </summary>
        RopReadStream = 0x2C,

        /// <summary>
        /// Rop WriteStream id
        /// </summary>
        RopWriteStream = 0x2D,

        /// <summary>
        /// Rop SeekStream id
        /// </summary>
        RopSeekStream = 0x2E,

        /// <summary>
        /// Rop SetStreamSize id
        /// </summary>
        RopSetStreamSize = 0x2F,

        /// <summary>
        /// Rop SetSearchCriteria id
        /// </summary>
        RopSetSearchCriteria = 0x30,

        /// <summary>
        /// Rop GetSearchCriteria id
        /// </summary>
        RopGetSearchCriteria = 0x31,

        /// <summary>
        /// Rop SubmitMessage id
        /// </summary>
        RopSubmitMessage = 0x32,

        /// <summary>
        /// Rop MoveCopyMessages id
        /// </summary>
        RopMoveCopyMessages = 0x33,

        /// <summary>
        /// Rop AbortSubmit id
        /// </summary>
        RopAbortSubmit = 0x34,

        /// <summary>
        /// Rop MoveFolder id
        /// </summary>
        RopMoveFolder = 0x35,

        /// <summary>
        /// Rop CopyFolder id
        /// </summary>
        RopCopyFolder = 0x36,

        /// <summary>
        /// Rop QueryColumnsAll id
        /// </summary>
        RopQueryColumnsAll = 0x37,

        /// <summary>
        /// Rop Abort id
        /// </summary>
        RopAbort = 0x38,

        /// <summary>
        /// Rop CopyTo id
        /// </summary>
        RopCopyTo = 0x39,

        /// <summary>
        /// Rop CopyToStream id
        /// </summary>
        RopCopyToStream = 0x3A,

        /// <summary>
        /// Rop CloneStream id
        /// </summary>
        RopCloneStream = 0x3B,

        /// <summary>
        /// Rop GetPermissionsTable id
        /// </summary>
        RopGetPermissionsTable = 0x3E,

        /// <summary>
        /// Rop GetRulesTable id
        /// </summary>
        RopGetRulesTable = 0x3F,

        /// <summary>
        /// Rop ModifyPermissions id
        /// </summary>
        RopModifyPermissions = 0x40,

        /// <summary>
        /// Rop ModifyRules id
        /// </summary>
        RopModifyRules = 0x41,

        /// <summary>
        /// Rop GetOwningServers id
        /// </summary>
        RopGetOwningServers = 0x42,

        /// <summary>
        /// Rop LongTermIdFromId id
        /// </summary>
        RopLongTermIdFromId = 0x43,

        /// <summary>
        /// Rop IdFromLongTermId id
        /// </summary>
        RopIdFromLongTermId = 0x44,

        /// <summary>
        /// Rop PublicFolderIsGhosted id
        /// </summary>
        RopPublicFolderIsGhosted = 0x45,

        /// <summary>
        /// Rop OpenEmbeddedMessage id
        /// </summary>
        RopOpenEmbeddedMessage = 0x46,

        /// <summary>
        /// Rop SetSpooler id
        /// </summary>
        RopSetSpooler = 0x47,

        /// <summary>
        /// Rop SpoolerLockMessage
        /// </summary>
        RopSpoolerLockMessage = 0x48,

        /// <summary>
        /// Rop GetAddressTypes id
        /// </summary>
        RopGetAddressTypes = 0x49,

        /// <summary>
        /// Rop TransportSend id
        /// </summary>
        RopTransportSend = 0x4A,

        /// <summary>
        /// Rop FastTransferSourceCopyMessages id
        /// </summary>
        RopFastTransferSourceCopyMessages = 0x4B,

        /// <summary>
        /// Rop FastTransferSourceCopyFolder id
        /// </summary>
        RopFastTransferSourceCopyFolder = 0x4C,

        /// <summary>
        /// Rop FastTransferSourceCopyTo id
        /// </summary>
        RopFastTransferSourceCopyTo = 0x4D,

        /// <summary>
        /// Rop FastTransferSourceGetBuffer id
        /// </summary>
        RopFastTransferSourceGetBuffer = 0x4E,

        /// <summary>
        /// Rop FindRow id
        /// </summary>
        RopFindRow = 0x4F,

        /// <summary>
        /// Rop Progress id
        /// </summary>
        RopProgress = 0x50,

        /// <summary>
        /// Rop TransportNewMail id
        /// </summary>
        RopTransportNewMail = 0x51,

        /// <summary>
        /// Rop FastTransferDestinationConfigure id
        /// </summary>
        RopFastTransferDestinationConfigure = 0x53,

        /// <summary>
        /// Rop FastTransferDestinationPutBuffer id
        /// </summary>
        RopFastTransferDestinationPutBuffer = 0x54,

        /// <summary>
        /// Rop GetNamesFromPropertyIds id
        /// </summary>
        RopGetNamesFromPropertyIds = 0x55,

        /// <summary>
        /// Rop GetPropertyIdsFromNames id
        /// </summary>
        RopGetPropertyIdsFromNames = 0x56,

        /// <summary>
        /// Rop UpdateDeferredActionMessages id
        /// </summary>
        RopUpdateDeferredActionMessages = 0x57,

        /// <summary>
        /// Rop EmptyFolder id
        /// </summary>
        RopEmptyFolder = 0x58,

        /// <summary>
        /// Rop ExpandRow id
        /// </summary>
        RopExpandRow = 0x59,

        /// <summary>
        /// Rop CollapseRow id
        /// </summary>
        RopCollapseRow = 0x5A,

        /// <summary>
        /// Rop LockRegionStream id
        /// </summary>
        RopLockRegionStream = 0x5B,

        /// <summary>
        /// Rop UnlockRegionStream id
        /// </summary>
        RopUnlockRegionStream = 0x5C,

        /// <summary>
        /// Rop CommitStream id
        /// </summary>
        RopCommitStream = 0x5D,

        /// <summary>
        /// Rop GetStreamSize id
        /// </summary>
        RopGetStreamSize = 0x5E,

        /// <summary>
        /// Rop QueryNamedProperties id
        /// </summary>
        RopQueryNamedProperties = 0x5F,

        /// <summary>
        /// Rop GetPerUserLongTermIds id
        /// </summary>
        RopGetPerUserLongTermIds = 0x60,

        /// <summary>
        /// Rop GetPerUserGuid id
        /// </summary>
        RopGetPerUserGuid = 0x61,

        /// <summary>
        /// Rop ReadPerUserInformation id
        /// </summary>
        RopReadPerUserInformation = 0x63,

        /// <summary>
        /// Rop WritePerUserInformation id
        /// </summary>
        RopWritePerUserInformation = 0x64,

        /// <summary>
        /// Rop SetReadFlags id
        /// </summary>
        RopSetReadFlags = 0x66,

        /// <summary>
        /// Rop CopyProperties id
        /// </summary>
        RopCopyProperties = 0x67,

        /// <summary>
        /// Rop GetReceiveFolderTable id
        /// </summary>
        RopGetReceiveFolderTable = 0x68,

        /// <summary>
        /// Rop FastTransferSourceCopyProperties id
        /// </summary>
        RopFastTransferSourceCopyProperties = 0x69,

        /// <summary>
        /// Rop GetCollapseState id
        /// </summary>
        RopGetCollapseState = 0x6B,

        /// <summary>
        /// Rop SetCollapseState id
        /// </summary>
        RopSetCollapseState = 0x6C,

        /// <summary>
        /// Rop GetTransportFolder id
        /// </summary>
        RopGetTransportFolder = 0x6D,

        /// <summary>
        /// Rop Pending id
        /// </summary>
        RopPending = 0x6E,

        /// <summary>
        /// Rop OptionsData id
        /// </summary>
        RopOptionsData = 0x6F,

        /// <summary>
        /// Rop SynchronizationConfigure id
        /// </summary>
        RopSynchronizationConfigure = 0x70,

        /// <summary>
        /// Rop SynchronizationImportMessageChange id
        /// </summary>
        RopSynchronizationImportMessageChange = 0x72,

        /// <summary>
        /// Rop SynchronizationImportHierarchyChange id
        /// </summary>
        RopSynchronizationImportHierarchyChange = 0x73,

        /// <summary>
        /// Rop SynchronizationImportDeletes id
        /// </summary>
        RopSynchronizationImportDeletes = 0x74,

        /// <summary>
        /// Rop SynchronizationUploadStateStreamBegin id
        /// </summary>
        RopSynchronizationUploadStateStreamBegin = 0x75,

        /// <summary>
        /// Rop SynchronizationUploadStateStreamContinue id
        /// </summary>
        RopSynchronizationUploadStateStreamContinue = 0x76,

        /// <summary>
        /// Rop SynchronizationUploadStateStreamEnd  id
        /// </summary>
        RopSynchronizationUploadStateStreamEnd = 0x77,

        /// <summary>
        /// Rop SynchronizationImportMessageMove id
        /// </summary>
        RopSynchronizationImportMessageMove = 0x78,

        /// <summary>
        /// Rop SetPropertiesNoReplicate id
        /// </summary>
        RopSetPropertiesNoReplicate = 0x79,

        /// <summary>
        /// Rop DeletePropertiesNoReplicate id
        /// </summary>
        RopDeletePropertiesNoReplicate = 0x7A,

        /// <summary>
        /// Rop GetStoreState id
        /// </summary>
        RopGetStoreState = 0x7B,

        /// <summary>
        /// Rop SynchronizationOpenCollector id
        /// </summary>
        RopSynchronizationOpenCollector = 0x7E,

        /// <summary>
        /// Rop GetLocalReplicaIds id
        /// </summary>
        RopGetLocalReplicaIds = 0x7F,

        /// <summary>
        /// Rop SynchronizationImportReadStateChanges id
        /// </summary>
        RopSynchronizationImportReadStateChanges = 0x80,

        /// <summary>
        /// Rop ResetTable id
        /// </summary>
        RopResetTable = 0x81,

        /// <summary>
        /// Rop SynchronizationGetTransferState id
        /// </summary>
        RopSynchronizationGetTransferState = 0x82,

        /// <summary>
        /// Rop TellVersion id
        /// </summary>
        RopTellVersion = 0x86,

        /// <summary>
        /// Rop FreeBookmark id
        /// </summary>
        RopFreeBookmark = 0x89,

        /// <summary>
        /// Rop WriteAndCommitStream id
        /// </summary>
        RopWriteAndCommitStream = 0x90,

        /// <summary>
        /// Rop HardDeleteMessages id
        /// </summary>
        RopHardDeleteMessages = 0x91,

        /// <summary>
        /// Rop HardDeleteMessagesAndSubfolders id
        /// </summary>
        RopHardDeleteMessagesAndSubfolders = 0x92,

        /// <summary>
        /// Rop SetLocalReplicaMidsetDeleted id
        /// </summary>
        RopSetLocalReplicaMidsetDeleted = 0x93,

        /// <summary>
        /// Rop Backoff id
        /// </summary>
        RopBackoff = 0xF9,

        /// <summary>
        /// Rop Logon id
        /// </summary>
        RopLogon = 0xFE,

        /// <summary>
        /// Rop BufferTooSmall id
        /// </summary>
        RopBufferTooSmall = 0xFF
    }
    #endregion

    #region Property
    #region PropertyType
    /// <summary>
    /// Property types.
    /// </summary>
    public enum PropertyType : ushort
    {
        /// <summary>
        /// The property value is compatible with the property types defined in [MS-OXCDATA].
        /// </summary>
        PtypUnspecified = 0x0000,

        /// <summary>
        /// Special type ID for int32.
        /// </summary>
        PtypInteger32 = 0x0003,

        /// <summary>
        /// Special type ID for String8
        /// </summary>
        PtypString8 = 0x001E,

        /// <summary>
        /// Special type ID for String
        /// </summary>
        PtypString = 0x001F,

        /// <summary>
        /// Special type ID for binary
        /// </summary>
        PtypBinary = 0x0102,

        /// <summary>
        /// Special type ID for MultiBinary
        /// </summary>
        PtypMultipleBinary = 0x1102,

        /// <summary>
        /// Special type ID for MultiString
        /// </summary>
        PtypMultipleString = 0x101F,

        /// <summary>
        /// Special type ID for MultiString8
        /// </summary>
        PtypMultipleString8 = 0x101E,

        /// <summary>
        /// Special type ID for BOOLEAN
        /// </summary>
        PtypBoolean = 0x000B,

        /// <summary>
        /// Special type ID for time
        /// </summary>
        PtypTime = 0x0040,

        /// <summary>
        /// Special type ID for OBJECT
        /// </summary>
        PtypComObject = 0x000D,

        /// <summary>
        /// Special type ID for Int64
        /// </summary>
        PtypInteger64 = 0x0014,

        /// <summary>
        /// Variable size, a 16-bit COUNT followed a structure.
        /// </summary>
        PtypServerId = 0x00FB,

        /// <summary>
        /// Variable size, a 16-bit COUNT of actions (not bytes) followed by that many Rule Action structures.
        /// </summary>
        PtypRuleAction = 0x00FE,

        /// <summary>
        /// Variable size, a byte array representing one or more Restriction structures.
        /// </summary>
        PtypRestriction = 0x00FD,

        /// <summary>
        /// 2 bytes, a 16-bit integer 
        /// </summary>
        PtypInteger16 = 0x0002,

        /// <summary>
        /// 4 bytes, a 32-bit floating point number
        /// </summary>
        PtypFloating32 = 0x0004,

        /// <summary>
        /// 8 bytes, a 64-bit floating point number
        /// </summary>
        PtypFloating64 = 0x0005,

        /// <summary>
        /// 8 bytes, a 64-bit signed, scaled integer representation of a decimal currency value, 
        /// with 4 places to the right of the decimal point.
        /// </summary>
        PtypCurrency = 0x0006,

        /// <summary>
        /// 8 bytes, a 64-bit floating point number in which the whole number part represents the number of days 
        /// since December 30, 1899, and the fractional part represents the fraction of a day since midnight
        /// </summary>
        PtypFloatingTime = 0x0007,

        /// <summary>
        /// 4 bytes, a 32-bit integer encoding error information
        /// </summary>
        PtypErrorCode = 0x000A,

        /// <summary>
        /// 16 bytes, a GUID with Data1, Data2, and Data3 fields in little-endian format.
        /// </summary>
        PtypGuid = 0x0048,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger16 values.
        /// </summary>
        PtypMultipleInteger16 = 0x1002,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger32 values.
        /// </summary>
        PtypMultipleInteger32 = 0x1003,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypFloating32 values.
        /// </summary>
        PtypMultipleFloating32 = 0x1004,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypFloating64 values.
        /// </summary>
        PtypMultipleFloating64 = 0x1005,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypCurrency values.
        /// </summary>
        PtypMultipleCurrency = 0x1006,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypFloatingTime values.
        /// </summary>
        PtypMultipleFloatingTime = 0x1007,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger64 values.
        /// </summary>
        PtypMultipleInteger64 = 0x1014,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypTime values.
        /// </summary>
        PtypMultipleTime = 0x1040,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypGuid values.
        /// </summary>
        PtypMultipleGuid = 0x1048
    }
    #endregion

    #region Kind
    /// <summary>
    /// This is the kind field enumeration for propertyName.
    /// </summary>
    public enum Kind : byte
    {
        /// <summary>
        /// The property is identified by the LID field. 
        /// </summary>
        LidField = 0x00,

        /// <summary>
        /// The property is identified by the Name field.
        /// </summary>
        NameField = 0x01,

        /// <summary>
        /// The property does not have an associated PropertyName.
        /// </summary>
        NoAssociated = 0xFF
    }
    #endregion

    #region PropertyNames
    /// <summary>
    /// This is an enumeration for property name definition.
    /// </summary>
    public enum PropertyNames
    {
        /// <summary>
        /// Define a property for default.
        /// </summary>
        None,

        /// <summary>
        /// Define a new property name for user specification.
        /// </summary>
        UserSpecified,

        /// <summary>
        /// Indicates whether the Message object contains at least one attachment. 
        /// </summary>       
        PidTagHasAttachments,

        /// <summary>
        /// Denotes the specific type of the Message object.  
        /// </summary>       
        PidTagMessageClass,

        /// <summary>
        /// Specifies the code page used to encode the non-Unicode string properties on   
        /// </summary>       
        PidTagMessageCodepage,

        /// <summary>
        /// Contains the Windows LCID of the end-user who created this message. 
        /// </summary>       
        PidTagMessageLocaleId,

        /// <summary>
        /// Contains the size in bytes consumed by the Message object on the server. 
        /// </summary>       
        PidTagMessageSize,

        /// <summary>
        /// Specifies the status of a message in a contents table. 
        /// </summary>       
        PidTagMessageStatus,

        /// <summary>
        /// Contains the prefix for the subject of the message. 
        /// </summary>       
        PidTagSubjectPrefix,

        /// <summary>
        /// Contains the normalized subject of the message. 
        /// </summary>       
        PidTagNormalizedSubject,

        /// <summary>
        /// Indicates the level of importance assigned by the end user to the Message object. 
        /// </summary>       
        PidTagImportance,

        /// <summary>
        /// Indicates the client's request for the priority at which the message is to be sent by the messaging system. 
        /// </summary>       
        PidTagPriority,

        /// <summary>
        /// Indicates the sender's assessment of the sensitivity of the Message object. 
        /// </summary>       
        PidTagSensitivity,

        /// <summary>
        /// Indicates whether the Message object has no end-user visible attachments. 
        /// </summary>       
        PidLidSmartNoAttach,

        /// <summary>
        /// Indicates whether the end-user wishes for this Message object to be hidden from other users who have access to the Message object. 
        /// </summary>       
        PidLidPrivate,

        /// <summary>
        /// Specifies how a Message object is handled by the client when acting on end-user input. 
        /// </summary>       
        PidLidSideEffects,

        /// <summary>
        /// Contains keywords or categories for the Message object. The length of each string within the multi-value string is less than 256
        /// </summary>
        PidNameKeywords,

        /// <summary>
        /// Indicates the start time for the Message object.
        /// </summary>
        PidLidCommonStart,

        /// <summary>
        /// Indicates the end time for the Message object. MUST be greater than or equal to the value of PidLidCommonStart. 
        /// </summary>
        PidLidCommonEnd,

        /// <summary>
        /// Indicates that this message has been automatically generated or automatically forwarded. 
        /// If this property is unset, then a default value of "0x00" is assumed
        /// </summary>
        PidTagAutoForwarded,

        /// <summary>
        /// Contains a comment added by the auto-forwarding agent.
        /// </summary>
        PidTagAutoForwardComment,

        /// <summary>
        /// Contains the unformatted text analogous to the text/plain body of [RFC2822]
        /// </summary>
        PidTagBody,

        /// <summary>
        /// Indicates the best available format for storing the message body
        /// </summary>
        PidTagNativeBody,

        /// <summary>
        /// Contains the HTML body as specified in [RFC2822]
        /// </summary>
        PidTagBodyHtml,

        /// <summary>
        /// Contains a Rich Text Format (RTF) body compressed as specified 
        /// </summary>
        PidTagRtfCompressed,

        /// <summary>
        /// Indicates whether the RTF body has been synchronized with the contents in PidTagBody
        /// </summary>
        PidTagRtfInSync,

        /// <summary>
        /// Indicates the code page used for PidTagBody or PidTagBodyHtml
        /// </summary>
        PidTagInternetCodepage,

        /// <summary>
        /// Contains the list of address book EntryIDs linked to by this Message object
        /// </summary>
        PidLidContactLinkEntry,

        /// <summary>
        /// Contains the PidTagDisplayName of each address book EntryID referenced in the value of PidLidContactLinkEntry. 
        /// </summary>
        PidLidContacts,

        /// <summary>
        /// Contains the elements of PidLidContacts, separated by a semicolon and a space ("; ").
        /// </summary>
        PidLidContactLinkName,

        /// <summary>
        /// Contains the list of search keys for the Contact object linked to by this Message object
        /// </summary>
        PidLidContactLinkSearchKey,

        /// <summary>
        /// Specifies the GUID of an archive tag. 
        /// </summary>
        PidTagArchiveTag,

        /// <summary>
        /// Specifies the GUID of a retention tag
        /// </summary>
        PidTagPolicyTag,

        /// <summary>
        /// Specifies the number of days that a Message object can be retained
        /// </summary>
        PidTagRetentionPeriod,

        /// <summary>
        /// A composite property that holds two pieces of information
        /// </summary>
        PidTagStartDateEtc,

        /// <summary>
        /// Specifies the date, in UTC, after which a Message object is expired by the server. 
        /// </summary>
        PidTagRetentionDate,

        /// <summary>
        /// Contains flags that specify the status or nature of an item's retention tag or archive tag
        /// </summary>
        PidTagRetentionFlags,

        /// <summary>
        /// Specifies the number of days that a Message object can remain un-archived
        /// </summary>
        PidTagArchivePeriod,

        /// <summary>
        /// Specifies the date, in UTC, after which a Message object is moved to archive by the server
        /// </summary>
        PidTagArchiveDate,

        /// <summary>
        /// Indicates the last time the file referenced by the Attachment object was modified, 
        /// or the last time the Attachment object itself was modified.
        /// </summary>
        PidTagLastModificationTime,

        /// <summary>
        /// Indicates the time the file referenced by the Attachment object was created
        /// </summary>
        PidTagCreationTime,

        /// <summary>
        /// Contains the name of the attachment as input by the end user
        /// </summary>
        PidTagDisplayName,

        /// <summary>
        /// Contains the size in bytes consumed by the Attachment object on the server
        /// </summary>
        PidTagAttachSize,

        /// <summary>
        /// Identifies the Attachment object within its Message object
        /// </summary>
        PidTagAttachNumber,

        /// <summary>
        /// Represents the way the contents of an attachment are accessed
        /// </summary>
        PidTagAttachMethod,

        /// <summary>
        /// Contains the full filename and extension of the Attachment object
        /// </summary>
        PidTagAttachLongFilename,

        /// <summary>
        /// Contains the 8.3 name of PidTagAttachLongFilename
        /// </summary>
        PidTagAttachFilename,

        /// <summary>
        /// Contains a filename extension that indicates the document type of an attachment
        /// </summary>
        PidTagAttachExtension,

        /// <summary>
        /// Contains the fully qualified path and filename with extension.
        /// </summary>
        PidTagAttachLongPathname,

        /// <summary>
        /// Contains the 8.3 name of PidTagAttachLongPathname
        /// </summary>
        PidTagAttachPathname,

        /// <summary>
        /// Contains the identifier information for the application which supplied the Attachment object's data
        /// </summary>
        PidTagAttachTag,

        /// <summary>
        /// Represents an offset, in rendered characters, to use when rendering an attachment within the main message text
        /// </summary>
        PidTagRenderingPosition,

        /// <summary>
        /// Contains a Windows metafile as specified in [MS-WMF] for the Attachment object
        /// </summary>
        PidTagAttachRendering,

        /// <summary>
        /// Indicates which body formats might reference this attachment when rendering data
        /// </summary>
        PidTagAttachFlags,

        /// <summary>
        /// Contains the name of an attachment file, modified so that it can be correlated with TNEF messages, see [MS-OXTNEF].
        /// </summary>
        PidTagAttachTransportName,

        /// <summary>
        /// Contains encoding information about the Attachment object
        /// </summary>
        PidTagAttachEncoding,

        /// <summary>
        /// MUST be unset if PidTagAttachEncoding is unset
        /// </summary>
        PidTagAttachAdditionalInformation,

        /// <summary>
        /// The type of Message object to which this attachment is linked
        /// </summary>
        PidTagAttachmentLinkId,

        /// <summary>
        /// Indicates special handling for this Attachment object
        /// </summary>
        PidTagAttachmentFlags,

        /// <summary>
        /// Indicates whether this Attachment object is hidden from the end user
        /// </summary>
        PidTagAttachmentHidden,

        /// <summary>
        /// The content-type MIME header.
        /// </summary>
        PidTagAttachMimeTag,

        /// <summary>
        /// A content identifier unique to this Message object that matches a corresponding 
        /// "cid:" Uniform Resource Identifier (URI) scheme reference in the HTML body of the Message object.
        /// </summary>
        PidTagAttachContentId,

        /// <summary>
        /// A relative or full URI that matches a corresponding reference in the HTML body of the Message object
        /// </summary>
        PidTagAttachContentLocation,

        /// <summary>
        /// The base of a relative URI. MUST be set if PidTagAttachContentLocation contains a relative URI.
        /// </summary>
        PidTagAttachContentBase,

        /// <summary>
        /// Indicates the client's access level to the object.
        /// </summary>
        PidTagAccessLevel,

        /// <summary>
        /// Contains the binary representation of the Attachment object in an application-specific format.
        /// </summary>
        PidTagAttachDataObject,

        /// <summary>
        /// Contains the contents of the file to be attached
        /// </summary>
        PidTagAttachDataBinary,

        /// <summary>
        /// Specifies the status of the Message object
        /// </summary>
        PidTagMessageFlags,

        /// <summary>
        /// Contains a list of blind carbon copy (Bcc) recipient display names
        /// </summary>
        PidTagDisplayBcc,

        /// <summary>
        /// Contains list of carbon copy (Cc) recipient display names
        /// </summary>
        PidTagDisplayCc,

        /// <summary>
        /// Contains a list of the primary recipient display names, separated by semicolons, if an e-mail message has primary recipient.
        /// </summary>
        PidTagDisplayTo,

        /// <summary>
        /// The description for security
        /// </summary>
        PidTagSecurityDescriptor,

        /// <summary>
        /// Setting the Url name
        /// </summary>
        PidTagUrlCompNameSet,

        /// <summary>
        /// Setting the sender is trusted
        /// </summary>
        PidTagTrustSender,

        /// <summary>
        /// The Url name
        /// </summary>
        PidTagUrlCompName,

        /// <summary>
        /// Contains a unique binary-comparable key that identifies an object for a search
        /// </summary>
        PidTagSearchKey,

        /// <summary>
        /// Indicates the operations available to the client for the object
        /// </summary>
        PidTagAccess,

        /// <summary>
        /// Contains the name of a Message object.
        /// </summary>
        PidTagCreatorName,

        /// <summary>
        /// The id for creator
        /// </summary>
        PidTagCreatorEntryId,

        /// <summary>
        /// Contains the name of the last mail user to modify the object.
        /// </summary>
        PidTagLastModifierName,

        /// <summary>
        /// The id for last modifier
        /// </summary>
        PidTagLastModifierEntryId,

        /// <summary>
        /// Have name or not for properties
        /// </summary>
        PidTagHasNamedProperties,

        /// <summary>
        /// Contains the Logon object LocaleID.
        /// </summary>
        PidTagLocaleId,

        /// <summary>
        /// Contains a global identifier (GID) indicating the last change to the object.
        /// </summary>
        PidTagChangeKey,

        /// <summary>
        /// Indicates the type of Server object.
        /// </summary>
        PidTagObjectType,

        /// <summary>
        /// Contains a unique binary-comparable identifier for a specific object.
        /// </summary>
        PidTagRecordKey,

        /// <summary>
        /// Contains the time a RopCreateMessage ([MS-OXCROPS] section 2.2.6.2) was processed.
        /// </summary>
        PidTagLocalCommitTime,

        /// <summary>
        /// Contains a value that indicates how to display an Address Book 
        /// object in a table or as an addressee on a message.
        /// </summary>
        PidTagDisplayType,

        /// <summary>
        /// PidTagAddressBookDisplayNamePrintable property
        /// </summary>
        PidTagAddressBookDisplayNamePrintable,

        /// <summary>
        /// Contains the AddressBook object's SMTP address.
        /// </summary>
        PidTagSmtpAddress,

        /// <summary>
        /// Contains a bitmask of message encoding preferences for mail sent to an 
        /// e-mail-enabled entity that is represented by this Address Book object.
        /// </summary>
        PidTagSendInternetEncoding,

        /// <summary>
        /// Contains a value that indicates how to display an Address 
        /// Book object in a table or as a recipient on a message.
        /// </summary>
        PidTagDisplayTypeEx,

        /// <summary>
        /// PidTagRecipientDisplayName property
        /// </summary>
        PidTagRecipientDisplayName,

        /// <summary>
        /// PidTagRecipientFlags property
        /// </summary>
        PidTagRecipientFlags,

        /// <summary>
        /// PidTagRecipientTrackStatus property
        /// </summary>
        PidTagRecipientTrackStatus,

        /// <summary>
        /// PidTagRecipientResourceState property
        /// </summary>
        PidTagRecipientResourceState,

        /// <summary>
        /// PidTagRecipientOrder property
        /// </summary>
        PidTagRecipientOrder,

        /// <summary>
        /// PidTagRecipientEntryId property
        /// </summary>
        PidTagRecipientEntryId,

        /// <summary>
        /// Contains the FID of the folder.
        /// </summary>
        PidTagFolderId,

        /// <summary>
        /// Contains a value that contains the MID of the message currently being synchronized.
        /// </summary>
        PidTagMid,

        /// <summary>
        /// Contains an identifier for all instances of a row in the table.
        /// </summary>
        PidTagInstID,

        /// <summary>
        /// PidTagInstanceNum property
        /// </summary>
        PidTagInstanceNum,

        /// <summary>
        /// PidTagSubject property
        /// </summary>
        PidTagSubject,

        /// <summary>
        /// PidTagMessageDeliveryTime property
        /// </summary>
        PidTagMessageDeliveryTime,

        /// <summary>
        /// PidTagRowType property
        /// </summary>
        PidTagRowType,

        /// <summary>
        /// PidTagContentCount property
        /// </summary>
        PidTagContentCount,

        /// <summary>
        /// PidTagOfflineAddressBookName property
        /// </summary>
        PidTagOfflineAddressBookName,

        /// <summary>
        /// PidTagOfflineAddressBookSequence property
        /// </summary>
        PidTagOfflineAddressBookSequence,

        /// <summary>
        /// PidTagOfflineAddressBookContainerGuid property
        /// </summary>
        PidTagOfflineAddressBookContainerGuid,

        /// <summary>
        /// PidTagOfflineAddressBookMessageClass property
        /// </summary>
        PidTagOfflineAddressBookMessageClass,

        /// <summary>
        /// PidTagOfflineAddressBookDistinguishedName property
        /// </summary>
        PidTagOfflineAddressBookDistinguishedName,

        /// <summary>
        /// PidTagSortLocaleId property
        /// </summary>
        PidTagSortLocaleId,

        /// <summary>
        /// PidTagEntryId property
        /// </summary>
        PidTagEntryId,

        /// <summary>
        /// PidTagMemberId property
        /// </summary>
        PidTagMemberId,

        /// <summary>
        /// PidTagMemberName property
        /// </summary>
        PidTagMemberName,

        /// <summary>
        /// PidTagMemberRights property
        /// </summary>
        PidTagMemberRights,

        /// <summary>
        /// PidTagRuleSequence property
        /// </summary>
        PidTagRuleSequence,

        /// <summary>
        /// PidTagRuleCondition property
        /// </summary>
        PidTagRuleCondition,

        /// <summary>
        /// PidTagRuleActions property
        /// </summary>
        PidTagRuleActions,

        /// <summary>
        /// PidTagRuleProvider property
        /// </summary>
        PidTagRuleProvider,

        /// <summary>
        /// Contains an IDSET of CNs for folders or normal messages in the current synchronization scope that 
        /// have been previously communicated to a client, and are reflected in its local replica.
        /// </summary>
        PidTagCnsetSeen,

        /// <summary>
        /// Contains a value that contains an internal identifier (GID) for this folder or message.
        /// </summary>
        PidTagSourceKey,

        /// <summary>
        /// Contains a value that contains a serialized representation of a PredecessorChangeList structure.
        /// </summary>
        PidTagPredecessorChangeList,

        /// <summary>
        /// PidTagParentSourceKey property
        /// </summary>
        PidTagParentSourceKey,

        /// <summary>
        /// PidTagFolderType property
        /// </summary>
        PidTagFolderType,

        /// <summary>
        /// PidTagTemplateData property
        /// </summary>
        PidTagTemplateData,

        /// <summary> 
        /// PidTagRowid property 
        /// </summary> 
        PidTagRowid,

        /// <summary> 
        /// PidTagTextAttachmentCharset property
        /// </summary> 
        PidTagTextAttachmentCharset,

        /// <summary>
        /// PidLidClassified property
        /// </summary>
        PidLidClassified,

        /// <summary>
        /// PidLidCategories property
        /// </summary>
        PidLidCategories,

        /// <summary>
        /// PidTagInternetReferences property
        /// </summary>
        PidTagInternetReferences,

        /// <summary>
        /// PidLidInfoPathFormName property
        /// </summary>
        PidLidInfoPathFormName,

        /// <summary>
        /// PidTagMimeSkeleton property
        /// </summary>
        PidTagMimeSkeleton,

        /// <summary>
        /// PidTagTnefCorrelationKey property
        /// </summary>
        PidTagTnefCorrelationKey,

        /// <summary>
        /// PidLidAgingDontAgeMe property
        /// </summary>
        PidLidAgingDontAgeMe,

        /// <summary>
        /// PidLidCurrentVersion property
        /// </summary>
        PidLidCurrentVersion,

        /// <summary>
        /// PidLidCurrentVersionName property
        /// </summary>
        PidLidCurrentVersionName,

        /// <summary>
        /// PidTagAlternateRecipientAllowed property
        /// </summary>
        PidTagAlternateRecipientAllowed,

        /// <summary>
        /// PidTagResponsibility property
        /// </summary>
        PidTagResponsibility,

        /// <summary>
        /// PidNameContentBase property
        /// </summary>
        PidNameContentBase,

        /// <summary>
        /// PidNameAcceptLanguage property
        /// </summary>
        PidNameAcceptLanguage,

        /// <summary>
        /// PidTagPurportedSenderDomain property
        /// </summary>
        PidTagPurportedSenderDomain,

        /// <summary>
        /// PidTagStoreEntryId property
        /// </summary>
        PidTagStoreEntryId,

        /// <summary>
        /// PidNameContentClass property
        /// </summary>
        PidNameContentClass,

        /// <summary>
        /// PidTagMessageRecipients property
        /// </summary>
        PidTagMessageRecipients,

        /// <summary>
        /// PidTagBodyContentId property
        /// </summary>
        PidTagBodyContentId,

        /// <summary>
        /// PidTagBodyContentLocation property
        /// </summary>
        PidTagBodyContentLocation,

        /// <summary>
        /// PidTagHtml property
        /// </summary>
        PidTagHtml,

        /// <summary>
        /// PidTagAttachPayloadClass property
        /// </summary>
        PidTagAttachPayloadClass,

        /// <summary>
        /// PidTagAttachPayloadProviderGuidString property
        /// </summary>
        PidTagAttachPayloadProviderGuidString,

        /// <summary>
        /// PidLidClassification property
        /// </summary>
        PidLidClassification,

        /// <summary>
        /// PidLidClassificationDescription property
        /// </summary>
        PidLidClassificationDescription,

        /// <summary>
        /// PidNameContentType property.
        /// </summary>
        PidNameContentType,

        /// <summary>
        /// PidTagRead property.
        /// </summary>
        PidTagRead
    }
    #endregion
    #endregion

    #region LogonFlags
    /// <summary>
    /// Contains additional flags that control the behavior of the logon. Individual flag values and their meanings are specified in the following table.
    /// </summary>
    [FlagsAttribute]
    public enum LogonFlags : byte
    {
        /// <summary>
        /// This is set for logon to public folders.
        /// </summary>
        PublicFolder = 0x00,

        /// <summary>
        /// This bit is set for logon to a private mailbox and is not set for logon to public folders.
        /// </summary>
        Private = 0x01,

        /// <summary>
        /// This bit is ignored by the server and is returned to the client in the response. 
        /// </summary>
        Undercover = 0x02,

        /// <summary>
        /// If the Private bit is set, this bit MUST NOT be set by the client and MUST be ignored by the server.
        /// If this bit is not set and the OpenFlags field does not have any of the following bits set
        /// ALTERNATE_SERVER
        /// IGNORE_HOME_MDB,
        /// the server will use the global directory to find the default public folder database to log on to. 
        /// If this server does not host that database, the ReturnValue will be ecWrongServer.
        /// Otherwise, the server will log on to the public folder database present on the server, if there is one. 
        /// If there is no public folder database on the server, ReturnValue will be ecLoginFailure. 
        /// </summary>
        Ghosted = 0x04,
    }
    #endregion

    #region ResponseFlags
    /// <summary>
    /// Contains flags that provide details about the state of the mailbox. Individual flag values and their meanings are specified in the following table.
    /// </summary>
    [FlagsAttribute]
    public enum ResponseFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// MUST be set.
        /// </summary>
        Reserved = 0x01,

        /// <summary>
        /// If set, the user has Full Owner or View Admin rights for the mailbox.
        /// </summary>
        OwnerRight = 0x02,

        /// <summary>
        /// If set, the user has the right to send mail from this mailbox.
        /// </summary>
        SendAsRight = 0x04,

        /// <summary>
        /// Indicates whether Out of Office (OOF) is set for the mailbox. 
        /// </summary>
        OOF = 0x10
    }
    #endregion

    #region OpenFlags
    /// <summary>
    /// Contains additional flags that control the behavior of the logon. Individual flag values and their meanings are specified in the following table.
    /// </summary>
    [FlagsAttribute]
    public enum OpenFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// When set, this bit indicates that the user is requesting administrative access to the mailbox. 
        /// To grant administrative access to the mailbox, the server MUST confirm that the user has the right to such access. 
        /// Confirmation is implementation-dependent.
        /// </summary>
        UseAdminPrivilege = 0x00000001,

        /// <summary>
        /// If set, RopLogon opens the public folders store. Otherwise, RopLogon opens a private user mailbox.
        /// </summary>
        Public = 0x00000002,

        /// <summary>
        /// This bit is ignored.
        /// </summary>
        HomeLogon = 0x00000004,

        /// <summary>
        /// This bit is ignored.
        /// </summary>
        TakeOwnerShip = 0x00000008,

        /// <summary>
        /// Requests a private server to provide an alternate public server.
        /// </summary>
        AlternateServer = 0x00000100,

        /// <summary>
        /// This bit is used only for public logons.When set, this bit allows the client to log on to a public MDB that is not the user's default public MDB; 
        /// otherwise, attempts to log on to a public MDB that is not the user's default results in the client being redirected back to the user's default public MDB.
        /// </summary>
        IgnoreHomeMDB = 0x00000200,

        /// <summary>
        /// Requests a non-messaging logon session. Non-messaging sessions allow clients to access the store, but do not allow messages to be sent or received.
        /// </summary>
        NoMail = 0x00000400,

        /// <summary>
        /// For a private-mailbox logon, the client uses this bit to control server behavior as specified in section 2.2.1.1.1.2.1[MS-OXCSTOR]. 
        /// For logons to a public folder store, this bit is ignored.
        /// </summary>
        UsePerMDBReplipMapping = 0x01000000
    }
    #endregion

    #region FolderOpenModeFlags
    /// <summary>
    /// The OpenModeFlags contains a bitmask of flags that indicate the open folder mode.
    /// </summary>
    [FlagsAttribute]
    public enum FolderOpenModeFlags : byte
    {
        /// <summary>
        /// No specific settings.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If this bit is not set, then it indicates the opening of an existing folder. 
        /// If this bit is set, then it indicates the opening of either an existing or soft deleted folder.
        /// </summary>
        OpenSoftDeleted = 0x04
    }
    #endregion

    #region MessageOpenModeFlags
    /// <summary>
    /// OpenModeFlags for RopOpenMessage.
    /// </summary>
    public enum MessageOpenModeFlags : byte
    {
        /// <summary>
        /// Message will be opened as read-only.
        /// </summary>
        ReadOnly = 0x00,

        /// <summary>
        /// Message will be opened for both reading and writing.
        /// </summary>
        ReadWrite = 0x01,

        /// <summary>
        /// Open for read/write if possible, read-only if not.
        /// </summary>
        BestAccess = 0x03,

        /// <summary>
        /// Open a soft deleted Message object if available.
        /// </summary>
        OpenSoftDeleted = 0x04
    }
    #endregion

    #region EmbeddedMessageOpenModeFlags
    /// <summary>
    /// OpenModeFlags for RopOpenEmbeddedMessage .
    /// </summary>
    public enum EmbeddedMessageOpenModeFlags : byte
    {
        /// <summary>
        /// Message will be opened as read-only.
        /// </summary>
        ReadOnly = 0x00,

        /// <summary>
        /// Message will be opened for both reading and writing.
        /// </summary>
        ReadWrite = 0x01,

        /// <summary>
        /// Create the attachment if it does not already exist and open the message for both reading and writing.
        /// </summary>
        Create = 0x02,
    }
    #endregion

    #region StreamOpenModeFlags
    /// <summary>
    /// OpenModeFlags for RopOpenStream.
    /// </summary>
    public enum StreamOpenModeFlags : byte
    {
        /// <summary>
        /// Open stream for read-only access.
        /// </summary>
        ReadOnly = 0x00,

        /// <summary>
        /// Open stream for read/write access.
        /// </summary>
        ReadWrite = 0x01,

        /// <summary>
        /// Opens new stream, this will delete the current property value and open stream for read/write access. This is required to open a property that has not been set.
        /// </summary>
        Create = 0x02,

        /// <summary>
        /// If the object this ROP is acting on was opened with ReadWrite access, then the stream MUST be opened with ReadWrite access. 
        /// Otherwise, the stream MUST be opened with ReadOnly access.
        /// </summary>
        BestAccess = 0x03
    }
    #endregion

    #region SaveFlags
    /// <summary>
    /// 1 byte indicating the server save behavior; MUST be one value from the following table.
    /// </summary>
    public enum SaveFlags : byte
    {
        /// <summary>
        /// The client requests that the server commit the changes. The server either returns an error and leaves the Message object open with unchanged access level, 
        /// or returns a success code and keeps the Message object open with read-only access.
        /// </summary>
        KeepOpenReadOnly = 0x01,

        /// <summary>
        /// The client requests that the server commit the changes. The server either returns an error and leaves the Message object open with unchanged access level, 
        /// or returns a success code and keeps the Message object open with read/write access.
        /// </summary>
        KeepOpenReadWrite = 0x02,

        /// <summary>
        /// The client requests that the server commit the changes. 
        /// The server either returns an error and leaves the Message object open with unchanged access level, 
        /// or returns a success code and keeps the Message object open with read/write access. 
        /// The ecObjectModified error code is not valid when this flag is set: the server overwrites any changes instead.
        /// </summary>
        ForceSave = 0x0C
    }
    #endregion

    #region CopyFlags for RopCopyProperties
    /// <summary>
    /// CopyFlags for RopCopyProperties
    /// CopyFlags is a BYTE bit field. Any bits not specified below MUST be ignored by the server.
    /// </summary>
    [FlagsAttribute]
    public enum RopCopyPropertiesCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If set, makes the call a move operation rather than a copy operation.
        /// </summary>
        Move = 0x01,

        /// <summary>
        /// If set, any properties being set by RopCopyProperties that already have a value on the destination object will not be overwritten; otherwise, they are overwritten.
        /// </summary>
        NoOverwrite = 0x02,

        /// <summary>
        /// Both move and overwrite
        /// </summary>
        MoveAndOverwrite = 0x03
    }
    #endregion

    #region CopyFlags for RopCopyTo
    /// <summary>
    /// CopyFlags for RopCopyTo
    /// </summary>
    [FlagsAttribute]
    public enum RopCopyToCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If set, makes the call a move operation rather than a copy operation.
        /// </summary>
        Move = 0x01,

        /// <summary>
        /// If set, any properties being set by RopCopyTo that already have a value on the destination object will not be overwritten; otherwise, they are overwritten.
        /// </summary>
        NoOverwrite = 0x02,
    }
    #endregion

    #region CopyFlags for RopFastTransferSourceCopyTo
    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyTo
    /// </summary>
    [FlagsAttribute]
    public enum RopFastTransferSourceCopyToCopyFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.
        /// If this flag is set, the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is not set, the client is not identifying the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is specified for a download operation, the server SHOULD NOT output any objects in a FastTransfer stream that the client does not have permissions to delete. 
        /// </summary>
        Move = 0x00000001,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 0x00000002,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 0x00000004,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 0x00000008,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused4 = 0x00000200,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused5 = 0x00000400,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a message.
        /// If set, the server SHOULD output the message body, and the body of embedded messages, in their original format.
        /// If not set, the server MUST output message body in the compressed Rich Text Format (RTF).
        /// </summary>
        BestBody = 0x00002000
    }
    #endregion

    #region CopyFlags for RopFastTransferDestinationConfigure
    /// <summary>
    /// CopyFlags for RopFastTransferDestinationConfigure
    /// </summary>
    [FlagsAttribute]
    public enum RopFastTransferDestinationConfigureCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.
        /// If this flag is set, the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is not set, the client is not identifying the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// </summary>
        Move = 0x01,
    }
    #endregion

    #region CopyFlags for RopFastTransferSourceCopyFolder
    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyFolder
    /// </summary>
    [FlagsAttribute]
    public enum RopFastTransferSourceCopyFolderCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If this flag is set, the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is not set, the client is not identifying the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// </summary>
        Move = 0x01,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 0x02,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 0x04,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 0x08,

        /// <summary>
        /// If this flag is set, the server MUST recursively include the subfolders of the folder specified in the InputServerObject in the scope. 
        /// If this flag is not set, the server MUST NOT recursively include the subfolders of the folder specified in the InputServerObject in the scope.
        /// </summary>
        CopySubfolders = 0x10,
    }
    #endregion

    #region CopyFlags for RopFastTransferSourceCopyMessages
    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyMessages
    /// </summary>
    [FlagsAttribute]
    public enum RopFastTransferSourceCopyMessagesCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder.
        /// If this flag is set, the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is not set, the client is not identifying the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is specified for a download operation, the server SHOULD NOT output any objects in a FastTransfer stream that the client does not have permissions to delete. 
        /// </summary>
        Move = 0x01,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 0x02,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 0x04,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 0x08,

        /// <summary>
        /// If set, the server SHOULD output the message body, and the body of embedded messages, in their original format.
        /// If not set, the server MUST output message bodies in the compressed RTF.
        /// </summary>
        BestBody = 0x10,

        /// <summary>
        /// If this flag is set, message and change identification information is not removed from output.
        /// </summary>
        SendEntryId = 0x20
    }
    #endregion

    #region CopyFlags for RopFastTransferSourceCopyProperties
    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyProperties
    /// </summary>
    [FlagsAttribute]
    public enum RopFastTransferSourceCopyPropertiesCopyFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.
        /// If this flag is set, the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is not set, the client is not identifying the FastTransfer operation being configured as a logical part of a larger object move operation.
        /// If this flag is specified for a download operation, the server SHOULD NOT output any objects in a FastTransfer stream that the client does not have permissions to delete. 
        /// </summary>
        Move = 0x01,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 0x02,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 0x04,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 0x08,
    }
    #endregion

    #region Format of distinguished names for all Address Book objects
    /// <summary>
    /// The format of distinguished names for all Address Book objects.
    /// </summary>
    public enum DNFormat
    {
        /// <summary>
        /// Addresslist-dn format in ABNF definition.
        /// </summary>
        AddressListDn = 0,

        /// <summary>
        /// Gal-addrlist-dn format in ABNF definition.
        /// </summary>
        GalAddrlistDn = 1,

        /// <summary>
        /// X500-dn format in ABNF definition.
        /// </summary>
        X500Dn,

        /// <summary>
        /// X500-dn with no container-rdn format in ABNF definition.
        /// </summary>
        X500DnWithNoContainerRdn,

        /// <summary>
        /// DN format in ABNF definition.
        /// </summary>
        Dn,
    }
    #endregion

    #region SendOptions
    /// <summary>
    /// Send option settings.
    /// </summary>
    [FlagsAttribute]
    public enum SendOptions : byte
    {
        /// <summary>
        /// When used on RopSynchronizationConfigure, MUST match the value of the Unicode SynchronizationFlag 
        /// </summary>
        Unicode = 0x01,

        /// <summary>
        /// If this flag is set, the Unicode flag MUST also be set.
        /// </summary>
        UseCpid = 0x02,

        /// <summary>
        /// Used in FastTransfer operations only when the client requests a FastTransfer stream with the intent of uploading it immediately to another destination server.
        /// The ROP that uses this flag MUST be followed by RopTellVersion.
        /// </summary>
        ForUpload = 0x03,

        /// <summary>
        /// Used when a client supports recovery mode and requests that a server MUST attempt to recover from failures to download changes for individual messages.
        /// MUST NOT be set when ForUpload flag is set. 
        /// </summary>
        RecoverMode = 0x04,

        /// <summary>
        /// See the following table for all possible combinations of encoding flags (Unicode and ForceUnicode).
        /// The following table lists all valid combinations of the Unicode | ForceUnicode flags. 
        /// Neither    String properties MUST be output in the code page set on connection. 
        /// Unicode    String properties MUST be output either in Unicode, or in the code page set on the current connection, with Unicode being preferred.
        /// Unicode | ForceUnicode    String properties MUST be output in Unicode.
        /// </summary>
        ForceUnicode = 0x08,

        /// <summary>
        /// MUST NOT be passed for anything but contents synchronization download.
        /// This flag is set if a client supports partial message downloads. 
        /// If a server supports this mode, it SHOULD output partial message changes if it reduces the size of the produced stream. 
        /// If a server does not support this mode, it does not output partial message changes and this flag is ignored 
        /// </summary>
        PartialItem = 0x10
    }
    #endregion

    #region FolderType
    /// <summary>
    /// The FolderType parameter contains the type of folder to be created. One of the values specified in the following table MUST be used.
    /// </summary>
    public enum FolderType : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Generic folder
        /// </summary>
        Genericfolder = 0x01,

        /// <summary>
        /// Search folder
        /// </summary>
        Searchfolder = 0x02
    }
    #endregion

    #region SourceOperation
    /// <summary>
    /// This enumeration is used to specify the type of data in a FastTransfer stream that would be uploaded by using RopFastTransferDestinationPutBuffer 
    /// on the FastTransfer upload context that is returned in the OutputServerObject field.
    /// </summary>
    public enum SourceOperation : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// RopFastTransferSourceCopyTo property
        /// </summary>
        CopyTo = 0x01,

        /// <summary>
        /// RopFastTransferSourceCopyProperties property
        /// </summary>
        CopyProperties = 0x02,

        /// <summary>
        /// RopFastTransferSourceCopyMessages property
        /// </summary>
        CopyMessages = 0x03,

        /// <summary>
        /// RopFastTransferSourceCopyFolder property
        /// </summary>
        CopyFolder = 0x04
    }
    #endregion

    #region PropertyRowFlag
    /// <summary>
    /// propertyRow Flag
    /// </summary>
    public enum PropertyRowFlag : byte
    {
        /// <summary>
        /// StandardPropertyRow property
        /// </summary>
        StandardPropertyRow = 0x00,

        /// <summary>
        /// FlaggedPropertyRow property
        /// </summary>
        FlaggedPropertyRow = 0x01
    }
    #endregion

    #region StoreState
    /// <summary>
    /// State information about the current mailbox.
    /// </summary>
    [FlagsAttribute]
    public enum StoreState
    {
        /// <summary>
        /// Default value without any specification.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// Indicates whether the mailbox has active search folders being populated. 
        /// </summary>
        StoreHasSearches = 0x01000000
    }
    #endregion

    #region TableStatus
    /// <summary>
    /// The table status refers to the status of any asynchronous operations being performed on the table. 
    /// The following values are used in the RopGetStatus, RopAbort, RopSetColumns, RopRestrict, and RopSortTable responses. 
    /// </summary>
    public enum TableStatus : byte
    {
        /// <summary>
        /// No operations are in progress.
        /// </summary>
        TblstatComplete = 0x00,

        /// <summary>
        /// A RopSortTable operation is in progress.
        /// </summary>
        TblstatSorting = 0x09,

        /// <summary>
        /// An error occurred during a RopSortTable operation.
        /// </summary>
        TblstatSortError = 0x0A,

        /// <summary>
        /// A RopSetColumns operation is in progress.
        /// </summary>
        TblstatSettingCols = 0x0B,

        /// <summary>
        /// An error occurred during a RopSetColumns operation. 
        /// </summary>
        TblstatSetColError = 0x0D,

        /// <summary>
        /// A RopRestrict operation is in progress.
        /// </summary>
        TblstatRestricting = 0x0E,

        /// <summary>
        /// An error occurred during a RopRestrict operation.
        /// </summary>
        TblstatRestrictError = 0x0F
    }
    #endregion

    #region Origin
    /// <summary>
    /// Specifies the location in stream.
    /// </summary>
    public enum Origin : byte
    {
        /// <summary>
        /// The new seek pointer is an offset relative to the beginning of the stream.
        /// </summary>
        Beginning = 0x00,

        /// <summary>
        /// The new seek pointer is an offset relative to the current seek pointer location.
        /// </summary>
        Current = 0x01,

        /// <summary>
        /// The new seek pointer is an offset relative to the end of the stream.
        /// </summary>
        End = 0x02,

        /// <summary>
        /// The invalid offset of the stream.
        /// </summary>
        Invalid = 0x03
    }
    #endregion

    #region RecipientType
    /// <summary>
    /// A Recipient Type is a bitwise OR of one value from the Types table with zero or more values from the flags table.
    /// </summary>
    [FlagsAttribute]
    public enum RecipientType : byte
    {
        /// <summary>
        /// Zero for bitwise.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Primary recipient.
        /// </summary>
        PrimaryRecipient = 0x01,

        /// <summary>
        /// Carbon copy (Cc) recipient.
        /// </summary>
        CcRecipient = 0x02,

        /// <summary>
        /// Bcc recipient.
        /// </summary>
        BccRecipient = 0x03,

        #region flags
        /// <summary>
        /// When resending a previous failure this flag indicates that this recipient did not successfully receive the message on the previous attempt.
        /// </summary>
        UnSuccess = 0x10,

        /// <summary>
        /// When resending a previous failure this flag indicates that this recipient did successfully receive the message on the previous attempt.
        /// </summary>
        Success = 0x80
        #endregion
    }
    #endregion

    #region MessageStatusFlags
    /// <summary>
    /// 4 bytes indicating the status flags that were set on the Message object prior to processing this request. 
    /// </summary>
    [FlagsAttribute]
    public enum MessageStatusFlags
    {
        /// <summary>
        /// Zero for bitwise.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// The message has been marked for downloading from the remote message store to the local client. 
        /// </summary>
        MsRemoteDownload = 0x00001000,

        /// <summary>
        /// This is a conflict resolve message as specified in [MS-OXCSYNC]. This is a read-only value for the client.
        /// </summary>
        MsInConflict = 0x00000800,

        /// <summary>
        /// The message has been marked for deletion at the remote message store without downloading to the local client. 
        /// </summary>
        MsRemoteDelete = 0x00002000
    }
    #endregion

    #region TransferStatus
    /// <summary>
    /// Represents the status of the download operation after producing data for the TransferBuffer field.
    /// </summary>
    public enum TransferStatus : ushort
    {
        /// <summary>
        /// The download stopped because a non-recoverable error has occurred when producing a FastTransfer stream. 
        /// The ReturnValue field of the ROP output buffer contains a code for that error.
        /// </summary>
        Error = 0x0000,

        /// <summary>
        /// The FastTransfer stream was split, and more data is available. TransferBuffer contains incomplete data. 
        /// </summary>
        Partial = 0x0001,

        /// <summary>
        /// This was the last portion of the FastTransfer stream.
        /// </summary>
        Done = 0x0003
    }
    #endregion

    #region DeleteFolderFlags
    /// <summary>
    /// The DeleteFolderFlags parameter contains a bitmask of flags that control the folder deletion operation.
    /// By default, RopDeleteFolder operates only on empty folders, but it can be used successfully on non-empty folders by setting two flags: DEL_FOLDERS and DEL_MESSAGES. 
    /// The DEL_FOLDERS flag enables all the folder's subfolders to be removed; the DEL_MESSAGES flag enables all the folder's messages to be removed. 
    /// RopDeleteFolder causes a hard delete of the folder if the DELETE_HARD_DELETE flag is set.
    /// </summary>
    [FlagsAttribute]
    public enum DeleteFolderFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If this bit is set, then delete all the messages in the folder.
        /// </summary>
        DelMessages = 0x01,

        /// <summary>
        /// If this bit is set, then delete the subfolder and all its subfolders.
        /// </summary>
        DelFolders = 0x04,

        /// <summary>
        /// If this bit is set, then the folder is hard deleted. If it is not set, the folder is soft deleted.
        /// </summary>
        DeleteHardDelete = 0x10
    }
    #endregion

    #region GetSearchFlags
    /// <summary>
    /// Contains the state of the current search
    /// </summary>
    [FlagsAttribute]
    public enum GetSearchFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// The search is running.
        /// </summary>
        Running = 0x00000001,

        /// <summary>
        /// The search is in the CPU-intensive mode of its operation, trying to locate messages that match the criteria. 
        /// If this flag is not set, the CPU-intensive part of the search's operation is over. 
        /// This flag only has meaning if the search is active (if the SEARCH_RUNNING flag is set).
        /// </summary>
        Rebuild = 0x00000002,

        /// <summary>
        /// The search is looking in specified search folder containers and all their child search folder containers for matching entries. 
        /// If this flag is not set, only the search folder containers that are explicitly included in the last call to the RopSetSearchCriteria are being searched.
        /// </summary>
        Recursive = 0x00000004,

        /// <summary>
        /// The search is running at a high priority relative to other searches. If this flag is not set, the search is running at a normal priority relative to other searches.
        /// </summary>
        ForGround = 0x00000008,

        /// <summary>
        /// The search results are complete.
        /// </summary>
        Complete = 0x00001000,

        /// <summary>
        /// The search results are not complete, as only some parts of messages were included.
        /// </summary>
        Partial = 0x00002000,

        /// <summary>
        /// The search is static.
        /// </summary>
        Static = 0x00010000,

        /// <summary>
        /// The search is still being evaluated.
        /// </summary>
        MaybeStatic = 0x00020000,

        /// <summary>
        /// The search is completely done using content indexing.
        /// </summary>
        CiTotally = 0x01000000,

        /// <summary>
        /// The search is mostly done using content indexing.
        /// </summary>
        CiWithTwirResidual = 0x02000000,

        /// <summary>
        /// The search is mostly done using store-only search.
        /// </summary>
        TwirMostly = 0x04000000,

        /// <summary>
        /// The search is completely done using store-only search.
        /// </summary>
        TwirTotally = 0x08000000
    }
    #endregion

    #region SetSearchFlags
    /// <summary>
    /// The SearchFlags parameter contains a bitmask of flags that control the search for a search folder.
    /// </summary>
    [FlagsAttribute]
    public enum SetSearchFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// Request server to abort the search. This flag cannot be set at the same time as the RESTART_SEARCH flag.
        /// </summary>
        StopSearch = 0x00000001,

        /// <summary>
        /// The search is initiated if this is the first call to RopSetSearchCriteria, or if the search is restarted, or if the search is inactive. 
        /// This flag cannot be set at the same time as the STOP_SEARCH flag.
        /// </summary>
        RestartSearch = 0x00000002,

        /// <summary>
        /// The search includes the search folder containers that are specified in the folder list in the request buffer and all their child folders. 
        /// This flag cannot be set at the same time as the SHALLOW_SEARCH flag.
        /// </summary>
        RecursiveSearch = 0x00000004,

        /// <summary>
        /// The search only looks in the search folder containers specified in the FolderIdList parameter for matching entries. 
        /// This flag cannot be set at the same time as the RECURSIVE_SEARCH flag. 
        /// Passing neither RECURSIVE_SEARCH nor SHALLOW_SEARCH indicates that the search will use the flag from the previous execution of RopSetSearchCriteria. 
        /// Also, passing neither RECURSIVE_SEARCH nor SHALLOW_SEARCH for the first search defaults to RECURSIVE_SHALLOW. 
        /// </summary>
        ShallowSearch = 0x00000008,

        /// <summary>
        /// Request the server to run this search at a high priority relative to other searches. This flag cannot be set at the same time as the BACKGROUND_SEARCH flag.
        /// </summary>
        ForGroundSearch = 0x00000010,

        /// <summary>
        /// Request the server to run this search at normal priority relative to other searches. This flag cannot be set at the same time as the FOREGROUND_SEARCH flag. 
        /// Passing neither FOREGROUND_SEARCH nor BACKGROUND_SEARCH indicates that the search will use the flag from the previous execution of RopSetSearchCriteria. 
        /// Passing neither FOREGROUND_SEARCH nor BACKGROUND_SEARCH on the first search defaults to BACKGROUND_SEARCH.
        /// </summary>
        BackGroundSearch = 0x00000020,

        /// <summary>
        /// Use content-indexed search exclusively.
        /// </summary>
        ContentIndexedSearch = 0x00010000,

        /// <summary>
        /// Never use content-indexed search.
        /// </summary>
        NonContentIndexedSearch = 0x00020000,

        /// <summary>
        /// Make the search static. 
        /// </summary>
        StaticSearch = 0x00040000
    }
    #endregion

    #region FolderTableFlags
    /// <summary>
    /// The TableFlags parameter contains a bitmask of flags that control how information is returned in the table on folder.
    /// </summary>
    [FlagsAttribute]
    public enum FolderTableFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Fills the hierarchy table with search folder containers from all levels. 
        /// If this flag is not set, the hierarchy table contains only the search folder container's immediate child search folder containers.
        /// </summary>
        Depth = 0x04,

        /// <summary>
        /// The ROP response can return immediately, possibly before the ROP execution is complete, 
        /// and in this case, the ReturnValue as well the RowCount fields in the return buffer might not be accurate. 
        /// Only ReturnValues reporting failure can be considered valid in this case.
        /// </summary>
        DeferredErrors = 0x08,

        /// <summary>
        /// Disables all notifications on this Table object.
        /// </summary>
        NoNotifications = 0x10,

        /// <summary>
        /// Enables the client to get a list of the soft-deleted folders.
        /// </summary>
        SoftDeletes = 0x20,

        /// <summary>
        /// Requests that the columns that contain string data be returned in Unicode format. 
        /// If UseUnicode is not present, then the string data will be encoded in the code page of the logon.
        /// </summary>
        UseUnicode = 0x40,

        /// <summary>
        /// Suppresses notifications generated by this client's actions on this Table object.
        /// </summary>
        SuppressesNotifications = 0x80
    }
    #endregion

    #region MsgTableFlags
    /// <summary>
    /// 8-bit flags structure. These flags control the type of table about message.
    /// </summary>
    [FlagsAttribute]
    public enum MsgTableFlags : byte
    {
        /// <summary>
        /// Open the table.
        /// </summary>
        Standard = 0x00,

        /// <summary>
        /// Open the table. Also requests that the columns containing string data be returned in Unicode format. 
        /// </summary>
        Unicode = 0x40
    }
    #endregion

    #region ModifyFlags
    /// <summary>
    /// 8-bit flags structure. These flags control behavior of RopModifyPermissions operation.
    /// </summary>
    [FlagsAttribute]
    public enum ModifyFlags : byte
    {
        /// <summary>
        /// Default value for no specification.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// This bit (bitmask 0x02) is the IncludeFreeBusy flag. If this bit is set, the server MUST use the values of the 
        /// FreeBusySimple and FreeBusyDetailed bits in the PidTagMemberRights property value when modifying the folder permissions. 
        /// If this bit is not set, the server MUST ignore the values of those bits. 
        /// The client SHOULD set this bit if the folder is the Calendar folder as specified in [MS-OXOSFLD] and the server version is greater 
        /// than or equal to 8.0.360.0, as specified in [MS-OXCRPC]. The client MUST NOT set this flag in any other circumstances.
        /// </summary>
        IncludeFreeBusy = 0x02,

        /// <summary>
        /// This bit (bitmask 0x01) is the ReplaceRows flag. If this bit is set, the server MUST replace any existing folder permissions, 
        /// and the client MUST NOT include any PermissionsDataFlags field values other than AddRow in this request. 
        /// If this bit is not set, the server MUST modify the existing folder permissions with the changes in this request (delete, modify, or add).
        /// </summary>
        ReplaceRows = 0x01
    }
    #endregion

    #region NotificationType
    /// <summary>
    /// A 12 bit enumeration defining the type of the notification
    /// </summary>
    [Flags]
    public enum NotificationType
    {
        /// <summary>
        /// The value of this field is 0x0
        /// </summary>
        NONE = 0x0,

        /// <summary>
        /// The notification is for NewMail events.
        /// </summary>
        NewMail = 0x0002,

        /// <summary>
        /// The notification is for ObjectCreated event.
        /// </summary>
        ObjectCreated = 0x0004,

        /// <summary>
        /// The notification is for ObjectDeleted event.
        /// </summary>
        ObjectDeleted = 0x0008,

        /// <summary>
        /// The notification is for ObjectModified event.
        /// </summary>
        ObjectModified = 0x0010,

        /// <summary>
        /// The notification is for ObjectMoved event.
        /// </summary>
        ObjectMoved = 0x0020,

        /// <summary>
        /// The notification is for ObjectCopied event.
        /// </summary>
        ObjectCopied = 0x0040,

        /// <summary>
        /// The notification is for SearchCompleted event.
        /// </summary>
        SearchCompleted = 0x0080,

        /// <summary>
        /// The notification is for TableModified events.
        /// </summary>
        TableModified = 0x0100,

        /// <summary>
        /// The notification is for StatusObjectModified event.
        /// </summary>
        StatusObjectModified = 0x0200,

        /// <summary>
        /// The value is reserved and MUST NOT be used.
        /// </summary>
        Reserved = 0x0400,

        /// <summary>
        /// Combination of all event:NewMail | ObjectCopied | ObjectCreated | ObjectDeleted | ObjectModified | ObjectMoved | SearchCompleted
        /// </summary>
        AllEvents = NewMail | ObjectCopied | ObjectCreated | ObjectDeleted | ObjectModified | ObjectMoved | SearchCompleted
    }

    /// <summary>
    /// Type of the notification for a TableModified event.
    /// </summary>
    public enum EventTypeOfTable
    {
        /// <summary>
        /// The value of this field is 0x0
        /// </summary>
        NONE = 0x0,

        /// <summary>
        /// The notification is for TableChanged event.
        /// </summary>
        TableChanged = 0x01,

        /// <summary>
        /// The notification is for TableRowAdded event.
        /// </summary>
        TableRowAdded = 0x03,

        /// <summary>
        /// The notification is for TableRowDeleted event.
        /// </summary>
        TableRowDeleted = 0x04,

        /// <summary>
        /// The notification is for TableRowModified event.
        /// </summary>
        TableRowModified = 0x05,

        /// <summary>
        /// The notification is for TableRestrictionChanged event.
        /// </summary>
        TableRestrictionChanged = 0x07
    }

    /// <summary>
    /// The flags bit in NotificationFlags
    /// </summary>
    [System.Flags]
    public enum FlagsBit
    {
        /// <summary>
        /// The value of this field is 0x0
        /// </summary>
        NONE = 0x0,

        /// <summary>
        /// The notification contains information about a change in total number of messages in a folder triggering the event.
        /// </summary>
        T = 0x1000,

        /// <summary>
        /// The notification contains information about a change in number of unread messages in a folder triggering the event. 
        /// </summary>
        U = 0x2000,

        /// <summary>
        /// The notification is caused by an event in a search folder. 
        /// </summary>
        S = 0x4000,

        /// <summary>
        /// The notification is caused by an event on a Message.
        /// </summary>
        M = 0x8000
    }
    #endregion

    #region PermTableFlags
    /// <summary>
    /// 8-bit flag structure. These flags control the type of table on permissions.
    /// </summary>
    [FlagsAttribute]
    public enum PermTableFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// This bit (bitmask 0x02) is the IncludeFreeBusy flag.
        /// </summary>
        IncludeFreeBusy = 0x02,
    }
    #endregion

    #region RPCAsyncStatus
    /// <summary>
    /// Represent the status of an asynchronous remote call
    /// </summary>
    public enum RPCAsyncStatus : uint
    {
        /// <summary>
        /// The call was completed successfully.
        /// </summary>
        RPC_S_OK = 0,

        /// <summary>
        /// The call has not yet completed.
        /// </summary>
        RPC_S_ASYNC_CALL_PENDING = 997,

        /// <summary>
        /// The asynchronous call handle is not valid.
        /// </summary>
        RPC_S_INVALID_ASYNC_HANDLE = 1914,
    }
    #endregion

    #region TableFlags
    /// <summary>
    /// 8-bit flags structure. These flags control the type of table on rules. 
    /// </summary>
    [FlagsAttribute]
    public enum TableFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// U (bit mask 0x40): Set if the client is requesting that string values in the table to be returned as Unicode strings.
        /// </summary>
        Normal = 0x40,

        /// <summary>
        /// An invalid tableFlag value
        /// </summary>
        Invalid = 0xab
    }
    #endregion

    #region Order
    /// <summary>
    /// The direction of the sort, these flags would be contained in SortOrder structure.
    /// </summary>
    [FlagsAttribute]
    public enum Order : byte
    {
        /// <summary>
        /// Sort by this column in ascending order.
        /// </summary>
        Ascending = 0x00,

        /// <summary>
        /// Sort by this column in descending order.
        /// </summary>
        Descending = 0x01,

        /// <summary>
        /// Indicates this is an aggregated column in a categorized sort, whose maximum value 
        /// (within the group of items with the same value of the previous category) is to be used as the sort key for the entire group.
        /// </summary>
        MaximumCategory = 0x04
    }
    #endregion

    #region AsynchronousFlags
    /// <summary>
    /// This is a BYTE field which contains an OR'ed combination of the asynchronous flags
    /// </summary>
    [FlagsAttribute]
    public enum AsynchronousFlags : byte
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// The server SHOULD perform the operation asynchronously. The server can perform the operation synchronously.
        /// </summary>
        TblAsync = 0x01,
    }
    #endregion

    #region QueryRowsFlags
    /// <summary>
    /// 8-bit flags structure. This field cannot contain both flags set simultaneously. 
    /// The field cannot have any of the other bits set.
    /// </summary>
    [FlagsAttribute]
    public enum QueryRowsFlags : byte
    {
        /// <summary>
        /// Advance the table cursor.
        /// </summary>
        Advance = 0x00,

        /// <summary>
        /// Do not advance the table cursor.
        /// </summary>
        NoAdvance = 0x01,

        /// <summary>
        /// Enable packed buffers for the response. To allow packed buffers to be used, 
        /// this flag is used in conjunction with the Chain flag (0x00000004) 
        /// that is passed in the pulFlags parameter of the EcDoRpcExt2 method. 
        /// </summary>
        EnablePackedBuffers = 0x02
    }
    #endregion

    #region FindRowFlags
    /// <summary>
    /// Byte for findRow direction settings.
    /// The field MUST NOT have any of the other bits set.
    /// </summary>
    public enum FindRowFlags : byte
    {
        /// <summary>
        /// Perform the find forwards.
        /// </summary>
        Forwards = 0x00,

        /// <summary>
        /// Perform the find backwards.
        /// </summary>
        Backwards = 0x01
    }
    #endregion

    #region RecipientFlags
    /// <summary>
    /// A flags field indicating which of several standard properties are present
    /// </summary>
    public enum RecipientFlags : ushort
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0x0000,

        /// <summary>
        /// 1-bit flag (mask 0x0080). If b'1', a different transport is responsible for delivery to this recipient.
        /// </summary>
        R = 0x0080,

        /// <summary>
        /// 1-bit flag (mask 0x0040). If b'1', the Transmittable Display Name is the same as the Display Name.
        /// </summary>
        S = 0x0040,

        /// <summary>
        /// 1-bit flag (mask 0x0020). If b'1', the TransmittableDisplayName field is included.
        /// </summary>
        T = 0x0020,

        /// <summary>
        /// 1-bit flag (mask 0x0010). If b'1', the DisplayName field is included.
        /// </summary>
        D = 0x0010,

        /// <summary>
        /// 1-bit flag (mask 0x0008). If b'1', the EmailAddress field is included.
        /// </summary>
        E = 0x0008,

        #region Type
        /// <summary>
        /// 3-bit enumeration (mask 0x0007). This enumeration specifies the type of address.
        /// X500DN 
        /// </summary>
        X500DN = 0x0001,

        /// <summary>
        /// MsMail property
        /// </summary>
        MsMail = 0x0002,

        /// <summary>
        /// SMTP property
        /// </summary>
        SMTP = 0x0003,

        /// <summary>
        /// Fax property
        /// </summary>
        Fax = 0x0004,

        /// <summary>
        /// ProfessionalOfficeSystem property
        /// </summary>
        ProfessionalOfficeSystem = 0x0005,

        /// <summary>
        /// PersonalDistributionList1 property
        /// </summary>
        PersonalDistributionList1 = 0x0006,

        /// <summary>
        /// PersonalDistributionList2 property
        /// </summary>
        PersonalDistributionList2 = 0x0007,
        #endregion

        /// <summary>
        /// 1-bit flag (mask 0x8000). If b'1', this recipient has a non-standard address type and the AddressType field is included.
        /// </summary>
        O = 0x8000,

        /// <summary>
        /// (4 bits):  (mask 0x7800) The server MUST set this to b'0000'.
        /// </summary>
        Reserved = 0x7800,

        /// <summary>
        /// 1-bit flag (mask 0x0400). If b'1', the SimpleDisplayName is included.
        /// </summary>
        I = 0x0400,

        /// <summary>
        /// 1-bit flag (mask 0x0200). If b'1', the associated string properties are in Unicode with a 2-byte null terminator; 
        /// if b'0', string properties are MBCS with a single null terminator, in the code page sent to the server in EcDoConnectEx
        /// </summary>
        U = 0x0200,

        /// <summary>
        /// 1-bit flag (mask 0x0100). This flag specifies that the recipient does not support receiving rich text messages.
        /// </summary>
        N = 0x0100
    }
    #endregion

    #region StringType
    /// <summary>
    /// 8-bit enumeration indicates the string type in ROP buffer.
    /// </summary>
    public enum StringType
    {
        /// <summary>
        /// There is no string present.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// The string is empty.
        /// </summary>
        Empty = 0x01,

        /// <summary>
        /// Null-terminated 8-bit character string. The null terminator is one zero byte.
        /// </summary>
        CharacterString = 0x02,

        /// <summary>
        /// Null-terminated Reduced Unicode character string. The null terminator is one zero byte.
        /// </summary>
        ReducedUnicodeCharacterString = 0x03,

        /// <summary>
        /// Null-terminated Unicode character string. The null terminator is 2 zero bytes.
        /// </summary>
        UnicodeCharacterString = 0x04
    }
    #endregion

    #region ReadFlags
    /// <summary>
    /// 1 byte containing a bitwise OR of zero or more values from the following table
    /// </summary>
    [FlagsAttribute]
    public enum ReadFlags : byte
    {
        /// <summary>
        /// The server sets the read flag and sends the receipt.
        /// </summary>
        Default = 0x00,

        /// <summary>
        /// The user requests that any pending read report be canceled; Server sets mfRead bit.
        /// </summary>
        SuppressReceipt = 0x01,

        /// <summary>
        /// Ignored by the server.
        /// </summary>
        Reserved = 0x0A,

        /// <summary>
        /// Server clears the mfRead bit; Client MUST include rfSuppressReceipt with this flag.
        /// </summary>
        ClearReadFlag = 0x04,

        /// <summary>
        /// The server sends a read report if one is pending, but does not change the mfRead bit.
        /// </summary>
        GenerateReceiptOnly = 0x10,

        /// <summary>
        /// The server clears the mfNotifyRead bit, but does not send a read report.
        /// </summary>
        ClearNotifyRead = 0x20,

        /// <summary>
        /// The server clears the mfNotifyUnread bit, but does not send a non-read report.
        /// </summary>
        ClearNotifyUnread = 0x40,
    }
    #endregion

    #region OpenAttachmentFlags
    /// <summary>
    /// 1 byte containing one of the following values.
    /// </summary>
    public enum OpenAttachmentFlags : byte
    {
        /// <summary>
        /// Message will be opened as read-only.
        /// </summary>
        ReadOnly = 0x00,

        /// <summary>
        /// Message will be opened for both reading and writing.
        /// </summary>
        ReadWrite = 0x01,

        /// <summary>
        /// Open for read/write if possible, read-only if not.
        /// </summary>
        BestAccess = 0x03
    }
    #endregion

    #region SubmitFlags
    /// <summary>
    /// When the client submits the message, the SubmitFlags value indicates how the message is to be delivered.
    /// The following table lists the possible values.
    /// </summary>
    public enum SubmitFlags : byte
    {
        /// <summary>
        /// The message without any flag
        /// </summary>
        None = 0x00,

        /// <summary>
        /// The message needs to be preprocessed by the server.
        /// </summary>
        PreProcess = 0x01,

        /// <summary>
        /// The message is to be processed by a client spooler.
        /// </summary>
        NeedsSpooler = 0x02
    }
    #endregion

    #region LockState
    /// <summary>
    /// Specifies a status to set on the message.
    /// </summary>
    public enum LockState : byte
    {
        /// <summary>
        /// Mark the message as locked.
        /// </summary>
        Lock = 0x00,

        /// <summary>
        /// Mark the message as unlocked.
        /// </summary>
        Unlock = 0x01,

        /// <summary>
        /// Mark the message as ready for processing by the server.
        /// </summary>
        Finished = 0x02
    }
    #endregion

    #region MessageFlags
    /// <summary>
    /// Specifies the status of the Message object. Set to a bitwise OR of zero or more of the values from the following tables.
    /// </summary>
    [FlagsAttribute]
    public enum MessageFlags
    {
        /// <summary>
        /// Default value for bitwise.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// The message is marked as having been read.
        /// </summary>
        MfRead = 0x00000001,

        /// <summary>
        /// The message is still being composed. This bit is cleared by the server when responding to RopSubmitMessage with a success code.
        /// </summary>
        MfUnsent = 0x00000008,

        /// <summary>
        /// The message includes a request for a resend operation with a non-delivery report.
        /// </summary>
        MfResend = 0x00000080,

        /// <summary>
        /// The message has not been modified since it was first saved (if unsent) or it was delivered (if sent).
        /// </summary>
        MfUnmodified = 0x00000002,

        /// <summary>
        /// The message is marked for sending as a result of a call to RopSubmitMessage 
        /// </summary>
        MfSubmitted = 0x00000004,

        /// <summary>
        /// The message has at least one attachment. This flag corresponds to the message's PidTagHasAttachments property. 
        /// </summary>
        MfHasAttach = 0x00000010,

        /// <summary>
        /// The user receiving the message was also the user who sent the message.
        /// </summary>
        MfFromMe = 0x00000020,

        /// <summary>
        /// The message is an FAI message.
        /// </summary>
        MfFAI = 0x00000040,

        /// <summary>
        /// The user who sent the message has requested notification when a recipient first reads it. 
        /// </summary>
        MfNotifyRead = 0x00000100,

        /// <summary>
        /// The user who sent the message has requested notification when a recipient deletes it before reading or the Message object expires as specified in [MS-OXOMSG].
        /// </summary>
        MfNotifyUnread = 0x00000200,

        /// <summary>
        /// The incoming message arrived over the Internet and originated either outside the organization or from a source the gateway does not consider trusted.
        /// </summary>
        MfInternet = 0x00002000,

        /// <summary>
        /// The incoming message arrived over an external link other than X.400 or the Internet. 
        /// It originated either outside the organization or from a source the gateway does not consider trusted.
        /// </summary>
        MfUntrusted = 0x00008000
    }
    #endregion

    #region Flags for RopGetPropertyIdsFromNames
    /// <summary>
    /// 8-bit flags structure. These flags control the behavior of RopGetPropertyIdsFromNames operation. 
    /// </summary>
    [FlagsAttribute]
    public enum GetPropertyIdsFromNamesFlags
    {
        /// <summary>
        /// Default value for bitwise.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// If set, indicates that the server MUST create new entries for any name parameters that are not found in the existing mapping set,
        /// and return existing entries for any name parameters that are found in the existing mapping set. 
        /// </summary>
        Create = 0x00000002
    }
    #endregion

    #region QueryFlags
    /// <summary>
    /// QueryFlags is a BYTE bit field.
    /// </summary>
    public enum QueryFlags : byte
    {
        /// <summary>
        /// Default value for bitwise.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Named properties with a Kind [MS-OXCDATA] of 0x1 MUST NOT be included in the response.
        /// </summary>
        NoStrings = 0x01,

        /// <summary>
        /// Named properties with a Kind [MS-OXCDATA] of 0x0 MUST NOT be included in the response.
        /// </summary>
        NoIds = 0x02
    }
    #endregion

    #region LockFlags
    /// <summary>
    /// LockFlags has only two signs: if OpenReadNoWriting or not.
    /// If any other value is granted, then reading and writing to the specified range of bytes is prohibited except by the owner that was granted this lock.
    /// </summary>
    public enum LockFlags
    {
        /// <summary>
        /// If this lock is granted, then the specified range of bytes can be opened and read any number of times, 
        /// but writing to the locked range is prohibited except for the owner that was granted this lock.
        /// </summary>
        OpenReadNoWriting = 0x00000001,

        /// <summary>
        /// If this lock is granted, then reading and writing to the specified range of bytes is prohibited except by the owner that was granted this lock.
        /// </summary>
        OtherValue = 0
    }
    #endregion

    #region PermissionDataFlags
    /// <summary>
    /// 8-bit flags structure. This field is used to specify the type of RopModifyPermissions operation.
    /// </summary>
    [FlagsAttribute]
    public enum PermissionDataFlags : byte
    {
        /// <summary>
        /// Default value for bitwise.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Adds new permissions that are specified in the PermissionData structure.
        /// </summary>
        AddRow = 0x01,

        /// <summary>
        /// Modifies the existing permissions for a user identified by the value of the PidTagMemberId property.
        /// </summary>
        ModifyRow = 0x02,

        /// <summary>
        /// Removes the existing permissions for a user identified by the value of the PidTagMemberId property.
        /// </summary>
        RemoveRow = 0x04
    }
    #endregion

    #region ModifyRulesFlags
    /// <summary>
    /// This is an 8-bit field with last bit used. 
    /// </summary>
    [FlagsAttribute]
    public enum ModifyRulesFlags : byte
    {
        /// <summary>
        /// Default value for bitwise.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// If this bit (bitmask 0x01) is set, the rules in this request are to replace existing rules in the folder; 
        /// in this case, all subsequent RuleData structures (see section 2.2.1.3) MUST have ROW_ADD as the value of their RuleDataFlag field (see section 2.2.1.3.1). 
        /// If this bit is not set, the rules specified in this request represent changes (delete, modify, add) to the rules already existing in this folder.
        /// </summary>
        R = 0x01
    }
    #endregion

    #region RuleDataFlags
    /// <summary>
    /// The RuleDataFlags field in the RuleData structure MUST have one of the following values.
    /// </summary>
    [FlagsAttribute]
    public enum RuleDataFlags : byte
    {
        /// <summary>
        /// Default value for no specification.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Adds the data in the rule buffer to the rule set as a new rule.
        /// </summary>
        RowAdd = 0x01,

        /// <summary>
        /// Modifies the existing rule identified by the value of PidTagRuleId property.
        /// </summary>
        RowModify = 0x02,

        /// <summary>
        /// Removes from the rule set the rule that has the same value of the PidTagRuleId property.
        /// </summary>
        RowRemove = 0x04
    }
    #endregion

    #region NotificationTypes
    /// <summary>
    /// These flags specify the types of events to register for.
    /// </summary>
    [FlagsAttribute]
    public enum NotificationTypes : byte
    {
        /// <summary>
        /// The server sends notifications to the client when NewMail events occur within the scope of interest.
        /// </summary>
        NewMail = 0x02,

        /// <summary>
        /// The server sends notifications to the client when ObjectCreated events occur within the scope of interest.
        /// </summary>
        ObjectCreated = 0x04,

        /// <summary>
        /// The server sends notifications to the client when ObjectDeleted events occur within the scope of interest.
        /// </summary>
        ObjectDeleted = 0x08,

        /// <summary>
        /// The server sends notifications to the client when ObjectModified events occur within the scope of interest.
        /// </summary>
        ObjectModified = 0x10,

        /// <summary>
        /// The server sends notifications to the client when ObjectMoved events occur within the scope of interest.
        /// </summary>
        ObjectMoved = 0x20,

        /// <summary>
        /// The server sends notifications to the client when ObjectCopied events occur within the scope of interest.
        /// </summary>
        ObjectCopied = 0x40,

        /// <summary>
        /// The server sends notifications to the client when SearchCompleted events occur within the scope of interest.
        /// </summary>
        SearchCompleted = 0x80
    }
    #endregion

    #region ImportFlag
    /// <summary>
    /// 8-bit flags structure. These flags control the behavior of the synchronization.
    /// </summary>
    [FlagsAttribute]
    public enum ImportFlag : byte
    {
        /// <summary>
        /// The message being imported is a normal message.
        /// </summary>
        Normal = 0x00,

        /// <summary>
        /// If this flag is set, the message being imported is an FAI message.
        /// </summary>
        Associated = 0x10,

        /// <summary>
        /// Identifies whether the server accepts conflicting versions of message.
        /// </summary>
        FailOnConflict = 0x40,

        /// <summary>
        /// Invalid parameter
        /// </summary>
        InvalidParameter = 0xaa
    }
    #endregion

    #region SynchronizationType
    /// <summary>
    /// An 8-bit enumeration that defines the type of synchronization requested: contents or hierarchy. 
    /// This field contributes to the synchronization scope. 
    /// </summary>
    public enum SynchronizationType : byte
    {
        /// <summary>
        /// Indicates a contents synchronization.
        /// </summary>
        Contents = 0x01,

        /// <summary>
        /// Indicates a hierarchy synchronization.
        /// </summary>
        Hierarchy = 0x02
    }
    #endregion

    #region SynchronizationFlag
    /// <summary>
    /// 8-bit enumeration. This value controls the type of synchronization.
    /// </summary>
    [FlagsAttribute]
    public enum SynchronizationFlag : ushort
    {
        /// <summary>
        /// Default Zero for bitwise.
        /// </summary>
        None = 0x0000,

        /// <summary>
        /// If this flag is set, the client supports Unicode and the server MUST output values of string 
        /// properties as they are stored, whether in Unicode or non-Unicode format.
        /// If this flag is not set, the client does not support Unicode and the server MUST output values 
        /// of string properties in the in the code page set on connection.
        /// This flag MUST match the value of the Unicode flag from SendOptions field.
        /// </summary>
        Unicode = 0x0001,

        /// <summary>
        /// If this flag is set, the server MUST NOT download information about item deletions 
        /// and the server MUST behave as if IgnoreNoLongerInScope was set.
        /// If this flag is not set, the server MUST download information about item deletions.
        /// The client MAY implement this flag
        /// </summary>
        NoDeletions = 0x0002,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.If this flag is set, 
        /// the server MUST NOT download information about messages that went out of scope as deletions.
        /// If this flag is not set, the server MUST download information about messages that went out of scope as deletions.
        /// The client MAY implement this flag.
        /// </summary>
        IgnoreNoLongerInScope = 0x0004,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, the server MUST also download information about changes to the read state of messages.
        /// If this flag is not set, the server MUST NOT download information about changes to the read state of messages.
        /// </summary>
        ReadState = 0x0008,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, the server MUST download information about changes to FAI messages.
        /// If this flag is not set, the server MUST NOT download information about changes to FAI messages.
        /// </summary>
        FAI = 0x0010,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, the server MUST download information about changes to normal messages.
        /// If this flag is not set, the server MUST NOT download information about changes to normal messages. 
        /// </summary>
        Normal = 0x0020,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, the server SHOULD limit properties and subobjects output 
        /// for top-level messages to the properties listed in PropertyTags.
        /// If this flag is not set, the server SHOULD exclude properties and subobjects 
        /// output for folders and top-level messages, if they are listed in PropertyTags.
        /// </summary>
        OnlySpecifiedProperties = 0x0080,

        /// <summary>
        /// If this flag is set, the server MUST ignore any persisted values for the PidTagSourceKey 
        /// and PidTagParentSourceKey properties when producing output for folder and message changes.
        /// If this flag is not set, the server MUST NOT ignore any persisted values for the PidTagSourceKey 
        /// and PidTagParentSourceKey properties when producing output for folder and message changes.
        /// Clients SHOULD set this flag. For more details about possible issues if this flag is not set
        /// </summary>
        NoForeignIdentifiers = 0x0100,

        /// <summary>
        /// MUST be set to "0" when sending. Servers MUST fail the ROP request if this flag is set.
        /// The client MAY implement this flag
        /// </summary>
        Reserved = 0x1000,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, a server SHOULD output message bodies in their original format.
        /// If this flag is not set, a server MUST output message bodies in the compressed RTF format.
        /// </summary>
        BestBody = 0x2000,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, all properties and subobjects of FAI messages MUST be output.
        /// If this flag is not set, the server ignores properties and subobjects of FAI messages.
        /// </summary>
        IgnoreSpecifiedOnFAI = 0x4000,

        /// <summary>
        /// MUST NOT be passed for anything but contents synchronization download.
        /// If this flag is set, the server SHOULD inject progress information into the output FastTransfer stream. 
        /// If this flag is not set, the server MUST not inject progress information into the output FastTransfer stream.
        /// This flag is in addition to the means of progress reporting available through the RopFastTransferSourceGetBuffer results.
        /// </summary>
        Progress = 0x8000
    }
    #endregion

    #region SynchronizationExtraFlag
    /// <summary>
    /// 32-bit flags structure. These flags control the additional behavior of the synchronization.
    /// </summary>
    [FlagsAttribute]
    public enum SynchronizationExtraFlag
    {
        /// <summary>
        /// A server MUST include PidTagFolderId (for hierarchy synchronization) or PidTagMid 
        /// (for contents synchronization) into a folder change or message change header IFF this flag is set.
        /// </summary>
        Eid = 0x00000001,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// A server MUST include the PidTagMessageSize property into a message change header IFF this flag is set.
        /// </summary>
        MessageSize = 0x00000002,

        /// <summary>
        /// A server MUST include the PidTagChangeNumber property into a message change header IFF this flag is set.
        /// </summary>
        CN = 0x00000004,

        /// <summary>
        /// MUST NOT be passed for anything but a contents synchronization download.
        /// If this flag is set, the server MUST sort messages by the value of their PidTagMessageDeliveryTime property ([MS-OXOMSG] section 2.2.3.9), 
        /// or by PidTagLastModificationTime ([MS-OXCMSG] section 2.2.2.2) if the former is missing, when generating 
        /// a sequence of messageChange elements for the FastTransfer stream, as specified in section 2.2.4.2.
        /// If this flag is not set, there is no requirement on the server to return items in a specific order.
        /// </summary>
        OrderByDeliveryTime = 0x00000008
    }
    #endregion

    #region Restrictions
    /// <summary>
    /// Restrictions describe a filter for limiting the view of a table to particular set of rows. 
    /// This filter represents a Boolean expression that is evaluated against each item of the table. 
    /// The item will be included as a row of the restricted table if and only if the value of the Boolean expression evaluates to TRUE.
    /// </summary>
    public enum Restrictions : byte
    {
        /// <summary>
        /// Logical AND operation applied to a list of subrestrictions.
        /// </summary>
        AndRestriction = 0x00,

        /// <summary>
        /// Logical OR operation applied to a list of subrestrictions.
        /// </summary>
        OrRestriction = 0x01,

        /// <summary>
        /// Logical NOT applied to a subrestriction.
        /// </summary>
        NotRestriction = 0x02,

        /// <summary>
        /// Search a property value for specific content.
        /// </summary>
        ContentRestriction = 0x03,

        /// <summary>
        /// Compare a property value to a particular value.
        /// </summary>
        PropertyRestriction = 0x04,

        /// <summary>
        /// Compare the values of two properties.
        /// </summary>
        ComparePropertiesRestriction = 0x05,

        /// <summary>
        /// Perform bitwise AND of a property value with a mask and compare to zero.
        /// </summary>
        BitMaskRestriction = 0x06,

        /// <summary>
        /// Compare the size of a property value to a particular figure.
        /// </summary>
        SizeRestriction = 0x07,

        /// <summary>
        /// Test whether a property has a value.
        /// </summary>
        ExistRestriction = 0x08,

        /// <summary>
        /// Test whether any row of a message's attachment or recipient table satisfies a subrestriction.
        /// </summary>
        SubObjectRestriction = 0x09,

        /// <summary>
        /// Associates a comment with a subrestriction.
        /// </summary>
        CommentRestriction = 0x0A,

        /// <summary>
        /// Limits the number of matches returned from a subrestriction.
        /// </summary>
        CountRestriction = 0x0B
    }
    #endregion

    #region FlaggedPropertyValue flag
    /// <summary>
    /// An 8-bit unsigned integer. This flag MUST be set one of three possible values: 0x0, 0x1, or 0xA, 
    /// which determines what is conveyed in the PropertyValue field. 
    /// </summary>
    public enum FlaggedPropertyValueFlag : byte
    {
        /// <summary>
        /// The PropertyValue field will be PropertyValue structure containing a value compatible with the property type implied the context.
        /// </summary>
        Present = 0x00,

        /// <summary>
        /// The PropertyValue field is not present. 
        /// </summary>
        NotPresent = 0x01,

        /// <summary>
        /// The PropertyValue field will be a PropertyValue structure containing an unsigned 32-bit integer. 
        /// This value is a property error code (see section 2.4.2) indicating why the property value is not present.
        /// </summary>
        Error = 0x0A
    }
    #endregion

    #region ErrorCodeValue flag
    /// <summary>
    /// The error code values are used to specify status from an NSPI method. 
    /// </summary>
    public enum ErrorCodeValue : uint
    {
        /// <summary>
        /// The operation succeeds.
        /// </summary>
        Success = 0x00000000,

        /// <summary>
        /// A request involving multiple properties that fails for one or more individual properties, 
        /// while succeeding overall.
        /// </summary>
        ErrorsReturned = 0x00040380,

        /// <summary>
        /// The operation fails for an unspecified reason.
        /// </summary>
        GeneralFailure = 0x80004005,

        /// <summary>
        /// The server does not support this method call.
        /// </summary>
        NotSupported = 0x80040102,

        /// <summary>
        /// A method call is made using a reference to an object that has been destroyed or is not in a viable state.
        /// </summary>
        InvalidObject = 0x80040108,

        /// <summary>
        /// Not enough of an unspecified resource is available to complete the operation.
        /// </summary>
        OutOfResources = 0x8004010E,

        /// <summary>
        /// The requested object cannot be found on the server.
        /// </summary>
        NotFound = 0x8004010F,

        /// <summary>
        /// A client is unable to log on to the server.
        /// </summary>
        LogonFailed = 0x80040111,

        /// <summary>
        /// The operation requested is too complex for the server to handle; 
        /// often applied to restrictions.
        /// </summary>
        TooComplex = 0x80040117,

        /// <summary>
        /// The server is not configured to support the code page requested by the client.
        /// </summary>
        InvalidCodepage = 0x8004011E,

        /// <summary>
        /// The server is not configured to support the locale requested by the client.
        /// </summary>
        InvalidLocale = 0x8004011F,

        /// <summary>
        /// The table is too big for the requested operation to complete.
        /// </summary>
        TableTooBig = 0x80040403,

        /// <summary>
        /// The bookmark passed to a table operation is not created on the same table.
        /// </summary>
        InvalidBookmark = 0x80040405,

        /// <summary>
        /// An unresolved recipient matches more than one entry in the directory.
        /// </summary>
        AmbiguousRecipient = 0x80040700,

        /// <summary>
        /// The caller does not have sufficient access rights to perform the operation.
        /// </summary>
        AccessDenied = 0x80070005,

        /// <summary>
        /// On get, indicates that the property or column value is too large to be retrieved by the request, 
        /// and the property value needs to be accessed with RopOpenStream. 
        /// </summary>
        NotEnoughMemory = 0x8007000E,

        /// <summary>
        /// An invalid parameter is passed to a remote procedure call.
        /// </summary>
        InvalidParameter = 0x80070057,
    }
    #endregion
}