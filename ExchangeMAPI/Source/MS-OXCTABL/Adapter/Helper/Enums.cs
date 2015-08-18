namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    /// <summary>
    /// The enum value for the table type
    /// </summary>
    public enum TableType : uint
    {
        /// <summary>
        /// Content Table
        /// </summary>
        CONTENT_TABLE = 0x00000000,

        /// <summary>
        /// Hierarchy Table
        /// </summary>
        HIERARCHY_TABLE = 0x00000001,

        /// <summary>
        /// Attachments Table
        /// </summary>
        ATTACHMENTS_TABLE = 0x00000002,

        /// <summary>
        /// Permissions Table
        /// </summary>
        PERMISSIONS_TABLE = 0x00000004,

        /// <summary>
        /// Rules Table
        /// </summary>
        RULES_TABLE = 0x00000010,

        /// <summary>
        /// Rules Table
        /// </summary>
        INVALID_TABLE = 0xFFFFFFFF,
    }

    /// <summary>
    /// The enum value for the Asynchronous ROP SetColumn
    /// </summary>
    public enum SetColumnsFlag : uint
    {
        /// <summary>
        /// This flag is used to identify no set columns
        /// is sent
        /// </summary>
        SETCOLUMNS_NOTSENT = 0x00000000,

        /// <summary>
        /// This flag is used to send the set columns synchronously
        /// </summary>
        SETCOLUMNS_SYNC = 0x00000011,

        /// <summary>
        /// This flag is used to send the set columns asynchronously
        /// </summary>
        SETCOLUMNS_ASYNC = 0x00000101,

        /// <summary>
        /// This flag is used to identify the set columns is sent
        /// </summary>
        SETCOLUMNS_SENT = 0x00000001,

        /// <summary>
        /// This flag is used to identify successful set columns
        /// response is received
        /// </summary>
        SETCOLUMNS_RESPONSESUCCESS = 0x00110001,

        /// <summary>
        /// This flag is used to identify failed set columns
        /// response is received
        /// </summary>
        SETCOLUMNS_RESPONSEFAIL = 0x01010001
    }

    /// <summary>
    /// The enum value for the Asynchronous ROP SortTable
    /// </summary>
    public enum SortTableFlag : uint
    {
        /// <summary>
        /// This flag is used to identify no sort table
        /// is sent
        /// </summary>
        SORTTABLE_NOTSENT = 0x00000000,

        /// <summary>
        /// This flag is used to send the sort table synchronously
        /// </summary>
        SORTTABLE_SYNC = 0x00000011,

        /// <summary>
        /// This flag is used to send the sort table asynchronously
        /// </summary>
        SORTTABLE_ASYNC = 0x00000101,

        /// <summary>
        /// This flag is used to identify the sort table is sent
        /// </summary>
        SORTTABLE_SENT = 0x00000001,

        /// <summary>
        /// This flag is used to identify successful sort table
        /// response is received
        /// </summary>
        SORTTABLE_RESPONSESUCCESS = 0x00110001,

        /// <summary>
        /// This flag is used to identify failed sort table
        /// response is received
        /// </summary>
        SORTTABLE_RESPONSEFAIL = 0x01010001,
    }

    /// <summary>
    /// The enum value for the Asynchronous ROP Restrict
    /// </summary>
    public enum RestrictTableFlag : uint
    {
        /// <summary>
        /// This flag is used to identify no set column
        /// is sent
        /// </summary>
        RESTRICT_NOTSENT = 0x00000000,

        /// <summary>
        /// This flag is used to send the restrict synchronously
        /// </summary>
        RESTRICT_SYNC = 0x00000011,

        /// <summary>
        /// This flag is used to send the restrict asynchronously
        /// </summary>
        RESTRICT_ASYNC = 0x00000101,

        /// <summary>
        /// This flag is used to identify the restrict is sent
        /// </summary>
        RESTRICT_SENT = 0x00000001,

        /// <summary>
        /// This flag is used to identify successful restrict
        /// response is received
        /// </summary>
        RESTRICT_RESPONSESUCCESS = 0x00110001,

        /// <summary>
        /// This flag is used to identify failed restrict
        /// response is received
        /// </summary>
        RESTRICT_RESPONSEFAIL = 0x01010001,
    }

    /// <summary>
    /// The enum value for Response types.
    /// </summary>
    public enum RopResponseType
    {
        /// <summary>
        /// Success response.
        /// </summary>
        SuccessResponse,

        /// <summary>
        /// Failure response.
        /// </summary>
        FailureResponse,

        /// <summary>
        /// Ordinary response
        /// </summary>
        Response,

        /// <summary>
        /// Null destination failure response.
        /// </summary>
        NullDestinationFailureResponse,

        /// <summary>
        /// Redirect response.
        /// </summary>
        RedirectResponse
    }

    /// <summary>
    /// The enum value for the return value
    /// of MS-OXCTABL protocol methods
    /// </summary>
    public enum TableRopReturnValues : uint
    {
        /// <summary>
        /// 0x00000000 will be returned in success.
        /// </summary>
        success = 0x00000000,

        /// <summary>
        /// This error will be returned when a property tag in the column 
        /// array is of type PT_UNSPECIFIED, PT_ERROR,  or an invalid type.
        /// </summary>
        ecInvalidParam = 0x80070057,

        /// <summary>
        /// This error will be returned if the object on which this ROP was 
        /// sent is not of type table.
        /// </summary>
        ecNotSupported = 0x80040102,

        /// <summary>
        /// This error will be returned if RopSetColumns has not been sent on 
        /// this table.
        /// </summary>
        ecNullObject = 0x000004B9,

        /// <summary>
        /// This error will be returned if the space allocated in the return 
        /// buffer is insufficient to fit at least one row of data.
        /// </summary>
        ecBufferTooSmall = 0x0000047D,

        /// <summary>
        /// This error will be returned if there were no asynchronous operations
        /// to abort, or the server was unable to abort the operations.
        /// </summary>
        ecUnableToAbort = 0x80040114,

        /// <summary>
        /// The error code will be returned if the bookmark sent in the request 
        /// is no longer valid
        /// </summary>
        ecNotFound = 0x8004010F,

        /// <summary>
        /// This error code will be returned if attempted to use the bookmark 
        /// after it was released
        /// </summary>
        ecInvalidBookmark = 0x80040405,

        /// <summary>
        /// This error code will be returned if the row 
        /// specified by the CategoryId field was not collapsed.
        /// </summary>
        ecNotCollapsed = 0x000004F8,

        /// <summary>
        /// This error code will be returned if
        /// The row specified by the CategoryId field was not expanded
        /// </summary>
        ecNotExpanded = 0x000004F7,

        /// <summary>
        /// This error code will be returned if
        /// the server does not implement this method call
        /// </summary>
        NotImplemented = 0x80040FFF,

        /// <summary>
        /// This error code will be returned by the server, but used to identify whether 
        /// the enabled property of the requirement implementation in the PtfConfig file is following one of Microsoft product
        /// </summary>
        unexpected = 0xffffffff
    }

    /// <summary>
    /// The enum value for the bookmark type
    /// </summary>
    public enum BookmarkType : byte
    {
        /// <summary>
        /// Points to the beginning position of the table, or the first row.
        /// </summary>
        BOOKMARK_BEGINNING = 0x00,

        /// <summary>
        /// Points to the current position of the table, or the current row. 
        /// </summary>
        BOOKMARK_CURRENT = 0x01,

        /// <summary>
        /// Points to the ending position of the table, or the location after the last row.
        /// </summary>
        BOOKMARK_END = 0x02,

        /// <summary>
        /// Points to the custom position in the table. Used with the BookmarkSize
        /// and bookmark fields.
        /// </summary>
        BOOKMARK_CUSTOM = 0x03
    }

    /// <summary>
    /// The enum value for the type of table Rops
    /// </summary>
    public enum TableRopType : byte
    {
        /// <summary>
        /// Perform set columns.
        /// </summary>
        SETCOLUMNS = 0x00,

        /// <summary>
        /// Perform sort table.  
        /// </summary>
        SORTTABLE = 0x01,

        /// <summary>
        /// Perform restrict table.
        /// </summary>
        RESTRICT = 0x02,

        /// <summary>
        /// Perform cursor move
        /// </summary>
        MOVECURSOR = 0x03,

        /// <summary>
        /// Perform bookmark creation
        /// </summary>
        CREATEBOOKMARK = 0x04,

        /// <summary>
        /// Reset table
        /// </summary>
        RESETTABLE = 0x05
    }

    /// <summary>
    /// The enum value for the cursor position
    /// </summary>
    public enum CursorPosition : byte
    {
        /// <summary>
        /// The cursor points to the begin row.
        /// </summary>
        BEGIN = 0x00,

        /// <summary>
        /// The cursor points to the end row.
        /// </summary>
        END = 0x01,

        /// <summary>
        /// The cursor points to the common right current row.
        /// </summary>
        CURRENT = 0x02
    }

    /// <summary>
    /// The value identify the order at RopSortTable
    /// </summary>
    public enum SortOrderFlag
    {
        /// <summary>
        /// Not sort table
        /// </summary>
        NotSort,

        /// <summary>
        /// The sort order is asc
        /// </summary>
        SortOrderASC,

        /// <summary>
        /// The sort order is desc
        /// </summary>
        SortOrderDESC
    }

    /// <summary>
    /// The enum value for the restrict type
    /// </summary>
    public enum RestrictFlag
    {
        /// <summary>
        /// Not restrict
        /// </summary>
        NotRestrict,

        /// <summary>
        /// Indicate a restrict that sender is Test1
        /// </summary>
        SenderIsTest1Restrict,

        /// <summary>
        /// Indicate a restrict that sender is Test2
        /// </summary>
        SenderIsTest2Restriction
    }

    /// <summary>
    /// ActionTypes used in ActionData of a RuleAction
    /// </summary>
    public enum ActionTypes : byte
    {
        /// <summary>
        /// Moves the message to a folder. MUST NOT be used in a public folder rule.
        /// </summary>
        OP_MOVE = 0x01,

        /// <summary>
        /// Copies the message to a folder. MUST NOT be used in a public folder rule.
        /// </summary>
        OP_COPY = 0x02,

        /// <summary>
        /// Replies to the message.
        /// </summary>
        OP_REPLY = 0x03,

        /// <summary>
        /// Sends an Out of Office (OOF) reply to the message.
        /// </summary>
        OP_OOF_REPLY = 0x04,

        /// <summary>
        /// Used for actions that cannot be executed by the server (like playing a sound). MUST NOT be used in a public folder rule.
        /// </summary>
        OP_DEFER_ACTION = 0x05,

        /// <summary>
        /// Rejects the message back to the sender.
        /// </summary>
        OP_BOUNCE = 0x06,

        /// <summary>
        /// Forwards the message to a recipient address.
        /// </summary>
        OP_FORWARD = 0x07,

        /// <summary>
        /// Assigns the message to another recipient.
        /// </summary>
        OP_DELEGATE = 0x08,

        /// <summary>
        /// Adds or changes a property on the message.
        /// </summary>
        OP_TAG = 0x09,

        /// <summary>
        /// Deletes the message.
        /// </summary>
        OP_DELETE = 0x0A,

        /// <summary>
        /// Sets the MSGFLAG_READ in the PidTagMessageFlags property on the message (see [MS-OXCMSG] section 2.2.1.6).
        /// </summary>
        OP_MARK_AS_READ = 0x0B
    }

    /// <summary>
    /// RuleState in PidTagRuleState of RuleData
    /// </summary>
    public enum RuleState : uint
    {
        /// <summary>
        /// The rule is enabled for execution. If this flag is not set, the server MUST skip this rule when evaluating rules.
        /// </summary>
        ST_ENABLED = 0x00000001,

        /// <summary>
        /// The server has encountered any non-parsing error processing the rule. The client SHOULD NOT set this flag. The server SHOULD ignore this flag if it is set by the client. 
        /// </summary>
        ST_ERROR = 0x00000002,

        /// <summary>
        /// The rule is executed only when the user sets the Out of Office state on the mailbox (see [MS-OXWOOF] section 2.2.5.2). This flag MUST NOT be set in a public folder rule. For details on this flag, see section 3.2.4.1.1.1.
        /// </summary>
        ST_ONLY_WHEN_OOF = 0x00000004,

        /// <summary>
        /// For details, see Out of Office Rule Processing in section 3.2. This flag MUST NOT be set in a public folder rule.
        /// </summary>
        ST_KEEP_OOF_HIST = 0x00000008,

        /// <summary>
        /// Rule evaluation will terminate after executing this rule, except for evaluation of Out of Office rules. For details, see Out of Office Rule Processing in section 3.2.4.1.1.1.
        /// </summary>
        ST_EXIT_LEVEL = 0x00000010,

        /// <summary>
        /// Evaluation of this rule will be skipped if the delivered message's PidTagContentFilterSpamConfidenceLevel property has a value of 0xFFFFFFFF.
        /// </summary>
        ST_SKIP_IF_SCL_IS_SAFE = 0x00000020,

        /// <summary>
        /// The server has encountered rule data from the client that is in an incorrect format, which caused an error parsing the rule data. The client SHOULD NOT set this flag. The server SHOULD ignore this flag if it is set by the client.
        /// </summary>
        ST_RULE_PARSE_ERROR = 0x00000040,

        /// <summary>
        /// Unused by this protocol. Bit locations marked with x are to be set to 0, SHOULD NOT be modified by the client and are ignored by the server.
        /// </summary>
        X = 0xFFFFFF80,

        /// <summary>
        /// Bit flag 0x00000080 is used to disable a specific Out of Office rule on Exchange 2007.
        /// </summary>
        X_DisableSpecificOOFRule = 0x00000080,

        /// <summary>
        /// Bit flag 0x00000100 has the same semantics as the ST_ONLY_WHEN_OOF bit flag on Exchange 2007.
        /// </summary>
        X_Same_Semantic_ST_ONLY_WHEN_OOF = 0x00000100
    }

    /// <summary>
    /// Specifies different RuleData generated from ptfconfig for different test purposes
    /// </summary>
    public enum TestRuleDataType : byte
    {
        /// <summary>
        /// Test data which will be used to test adding a new rule
        /// </summary>
        ForAdd = 0x01,

        /// <summary>
        /// Test data which will be used to test modifying an existing rule
        /// </summary>
        ForModify = 0x02,

        /// <summary>
        /// Test data which will be used to test removing an existing rule
        /// </summary>
        ForRemove = 0x04,
    }

    /// <summary>
    /// Property name for test purpose.
    /// E.g. PidTagRuleName
    /// </summary>
    public enum TestPropertyName : ushort
    {
        /// <summary>
        /// PropertyId of PidTagReceivedByEmailAddress
        /// </summary>
        PidTagReceivedByEmailAddress = 0x0076,

        /// <summary>
        /// PropertyId of PidTagContentUnreadCount
        /// </summary>
        PidTagContentUnreadCount = 0x3603,

        /// <summary>
        /// PropertyId of PidTagSearchKey
        /// </summary>
        PidTagSearchKey = 0x300B,

        /// <summary>
        /// PropertyId of PidTagAddressType
        /// </summary>
        PidTagAddressType = 0x3002,

        /// <summary>
        /// PropertyId of PidTagRuleActionNumber
        /// </summary>
        PidTagRuleActionNumber = 0x6650,

        /// <summary>
        /// PropertyId of PidTagRuleActionType
        /// </summary>
        PidTagRuleActionType = 0x6649,

        /// <summary>
        /// PropertyId of PidTagRuleError
        /// </summary>
        PidTagRuleError = 0x6648,

        /// <summary>
        /// PropertyId of PidTagMid
        /// </summary>
        PidTagMid = 0x674A,

        /// <summary>
        /// PropertyId of PidTagDeferredActionMessageOriginalEntryId
        /// </summary>
        PidTagDeferredActionMessageOriginalEntryId = 0x6741,

        /// <summary>
        /// PropertyId of PidTagFolderId
        /// </summary>
        PidTagFolderId = 0x6748,

        /// <summary>
        /// PropertyId of PidTagReplyTemplateId
        /// </summary>
        PidTagReplyTemplateId = 0x65C2,

        /// <summary>
        /// PropertyId of PidTagSmtpAddress
        /// </summary>
        PidTagSmtpAddress = 0x39FE,

        /// <summary>
        /// PropertyId of PidTagRecipientType
        /// </summary>
        PidTagRecipientType = 0x0C15,

        /// <summary>
        /// PropertyId of PidTagEmailAddress
        /// </summary>
        PidTagEmailAddress = 0x3003,

        /// <summary>
        /// PropertyId of PidTagDisplayName
        /// </summary>
        PidTagDisplayName = 0x3001,

        /// <summary>
        /// PropertyId of PidTagSubject
        /// </summary>
        PidTagSubject = 0x0037,

        /// <summary>
        /// PropertyId of PidTagRuleFolderEntryId
        /// </summary>
        PidTagRuleFolderEntryId = 0x6651,

        /// <summary>
        /// PropertyId of PidTagDamOriginalEntryId
        /// </summary>
        PidTagDamOriginalEntryId = 0x6646,

        /// <summary>
        /// PropertyId of PidTagDamBackPatched
        /// </summary>
        PidTagDamBackPatched = 0x6647,

        /// <summary>
        /// PropertyId of PidTagRuleIds
        /// </summary>
        PidTagRuleIds = 0x6675,

        /// <summary>
        /// PropertyId of PidTagClientActions
        /// </summary>
        PidTagClientActions = 0x6645,

        /// <summary>
        /// PropertyId of PidTagEntryId
        /// </summary>
        PidTagEntryId = 0x0FFF,

        /// <summary>
        /// PropertyId of PidTagRuleUserFlags
        /// </summary>
        PidTagRuleUserFlags = 0x6678,

        /// <summary>
        /// PropertyId of PidTagRuleProviderData
        /// </summary>
        PidTagRuleProviderData = 0x6684,

        /// <summary>
        /// PropertyId of PidTagHasRules
        /// </summary>
        PidTagHasRules = 0x663A,

        /// <summary>
        /// PropertyId of PidTagExtendedRuleSizeLimit
        /// </summary>
        PidTagExtendedRuleSizeLimit = 0x0E9B,

        /// <summary>
        /// PropertyId of PidTagRuleId
        /// </summary>
        PidTagRuleId = 0x6674,

        /// <summary>
        ///  PropertyId of PidTagRuleSequence
        /// </summary>
        PidTagRuleSequence = 0x6676,

        /// <summary>
        /// PropertyId of PidTagRuleState
        /// </summary>
        PidTagRuleState = 0x6677,

        /// <summary>
        /// PropertyId of PidTagRuleName
        /// </summary>
        PidTagRuleName = 0x6682,

        /// <summary>
        /// PropertyId of PidTagRuleProvider
        /// </summary>
        PidTagRuleProvider = 0x6681,

        /// <summary>
        /// PropertyId of PidTagRuleLevel
        /// </summary>
        PidTagRuleLevel = 0x6683,

        /// <summary>
        /// PropertyId of PidTagRuleCondition
        /// </summary>
        PidTagRuleCondition = 0x6679,

        /// <summary>
        /// PropertyId of PidTagRuleActions
        /// </summary>
        PidTagRuleActions = 0x6680,

        /// <summary>
        /// PropertyId of PidTagHasDeferredActionMessages
        /// </summary>
        PidTagHasDeferredActionMessages = 0x3FEA,

        /// <summary>
        /// PropertyId of PidTagMessageFlags
        /// </summary>
        PidTagMessageFlags = 0x0E07,

        /// <summary>
        /// PropertyId of PidTagRuleMessageName
        /// </summary>
        PidTagRuleMessageName = 0x65EC,

        /// <summary>
        /// PropertyId of PidTagMessageClass
        /// </summary>
        PidTagMessageClass = 0x001a,

        /// <summary>
        /// PropertyId of PidTagRuleMessageSequence
        /// </summary>
        PidTagRuleMessageSequence = 0x65F3,

        /// <summary>
        ///  PropertyId of PidTagRuleMessageState
        /// </summary>
        PidTagRuleMessageState = 0x65e9,

        /// <summary>
        ///  PropertyId of PidTagRuleMessageLevel
        /// </summary>
        PidTagRuleMessageLevel = 0x65ed,

        /// <summary>
        /// PropertyId of PidTagRuleMessageProvider
        /// </summary>
        PidTagRuleMessageProvider = 0x65eb,

        /// <summary>
        /// PropertyId of PidTagRuleMessageUserFlags
        /// </summary>
        PidTagRuleMessageUserFlags = 0x65ea,

        /// <summary>
        /// PropertyId of PidTagRuleMessageProviderData
        /// </summary>
        PidTagRuleMessageProviderData = 0x65ee,

        /// <summary>
        /// PropertyId of PidTagExtendedRuleMessageActions 
        /// </summary>
        PidTagExtendedRuleMessageActions = 0x0e99,

        /// <summary>
        /// PropertyId of PidTagExtendedRuleMessageCondition 
        /// </summary>
        PidTagExtendedRuleMessageCondition = 0x0e9a,

        /// <summary>
        /// PropertyId of PidTagObjectType  
        /// </summary>
        PidTagObjectType = 0x0ffe,

        /// <summary>
        /// PropertyId of PidTagDisplayType 
        /// </summary>
        PidTagDisplayType = 0x3900,

        /// <summary>
        /// PropertyId of PidTagAddressBookDisplayNamePrintable 
        /// </summary>
        PidTagAddressBookDisplayNamePrintable = 0x39ff,

        /// <summary>
        /// PropertyId of PidTagSendInternetEncoding 
        /// </summary>
        PidTagSendInternetEncoding = 0x3a71,

        /// <summary>
        /// PropertyId of PidTagDisplayTypeEx 
        /// </summary>
        PidTagDisplayTypeEx = 0x3905,

        /// <summary>
        /// PropertyId of PidTagRecipientDisplayName 
        /// </summary>
        PidTagRecipientDisplayName = 0x5ff6,

        /// <summary>
        /// PropertyId of PidTagRecipientFlags 
        /// </summary>
        PidTagRecipientFlags = 0x5ffd,

        /// <summary>
        /// PropertyId of PidTagRecipientTrackStatus 
        /// </summary>
        PidTagRecipientTrackStatus = 0x5fff,

        /// <summary>
        /// PropertyId of PidTagRecipientResourceState 
        /// </summary>
        PidTagRecipientResourceState = 0x5fde,

        /// <summary>
        /// PropertyId of PidTagRecipientOrder 
        /// </summary>
        PidTagRecipientOrder = 0x5fdf,

        /// <summary>
        /// PropertyId of PidTagRecipientEntryId 
        /// </summary>
        PidTagRecipientEntryId = 0x5ff7,

        /// <summary>
        /// PropertyId of PidTagSentRepresentingEmailAddress
        /// </summary>
        PidTagSentRepresentingEmailAddress = 0x0065,

        /// <summary>
        /// PropertyId of PidTagBody
        /// </summary>
        PidTagBody = 0x1000,

        /// <summary>
        /// PropertyId of PidTagAutoResponseSuppress
        /// </summary>
        PidTagAutoResponseSuppress = 0x3fdf,

        /// <summary>
        /// PropertyId of PidTagAutoForwarded
        /// </summary>
        PidTagAutoForwarded = 0x0005,

        /// <summary>
        ///  PropertyId of PidTagReceivedRepresentingEntryId
        /// </summary>
        PidTagReceivedRepresentingEntryId = 0x0043,

        /// <summary>
        /// PropertyId of PidTagReceivedRepresentingAddressType
        /// </summary>
        PidTagReceivedRepresentingAddressType = 0x0077,

        /// <summary>
        /// PropertyId of PidTagReceivedRepresentingEmailAddress
        /// </summary>
        PidTagReceivedRepresentingEmailAddress = 0x0078,

        /// <summary>
        /// PropertyId of PidTagReceivedRepresentingName
        /// </summary>
        PidTagReceivedRepresentingName = 0x0044,

        /// <summary>
        /// PropertyId of PidTagReceivedRepresentingSearchKey
        /// </summary>
        PidTagReceivedRepresentingSearchKey = 0x0052,

        /// <summary>
        /// PropertyId of PidTagDelegatedByRule
        /// </summary>
        PidTagDelegatedByRule = 0x3fe3,

        /// <summary>
        /// PropertyId of PidTagImportance
        /// </summary>
        PidTagImportance = 0x0017
    }

    /// <summary>
    /// The enumeration of PropertyTypeName
    /// </summary>
    public enum PropertyTypeName : ushort
    {
        /// <summary>
        /// PropertyType of PtypString
        /// </summary>
        PtypString = 0x001F,

        /// <summary>
        /// PropertyType of PtypString8
        /// </summary>
        PtypString8 = 0x001E,

        /// <summary>
        /// PropertyType of PtypInteger32
        /// </summary>
        PtypInteger32 = 0x0003,

        /// <summary>
        /// PropertyType of PtypBinary
        /// </summary>
        PtypBinary = 0x0102,

        /// <summary>
        /// PropertyType of PtypBoolean
        /// </summary>
        PtypBoolean = 0x000B,

        /// <summary>
        /// PropertyType of PtypServerId
        /// </summary>
        PtypServerId = 0x00FB,

        /// <summary>
        /// PropertyType of PtypRuleAction
        /// </summary>
        PtypRuleAction = 0x000FE,

        /// <summary>
        /// PropertyType of PtypInteger64
        /// </summary>
        PtypInteger64 = 0x0014,

        /// <summary>
        /// PropertyType of PtypRestriction
        /// </summary>
        PtypRestriction = 0x00FD
    }

    /// <summary>
    /// This enum is used to specify the modify rule flag used in modify rule rop operation.
    /// </summary>
    public enum ModifyRuleFlag : byte
    {
        /// <summary>
        /// If this bit is set, the rules in this request are to replace the existing set of rules in the folder;
        /// in this case, all subsequent RuleData structures MUST have ROW_ADD as the value of their RuleDataFlag field.
        /// </summary>
        Modify_ReplaceAll = (byte)0x01,

        /// <summary>
        /// The rules specified in this request represent changes (delete, modify, and add) to the set of rules already existing in this folder.
        /// </summary>
        Modify_OnExisting = (byte)0x00,

        /// <summary>
        /// Beside the last bit, all other bit is unused, they are ignored by the server if set.
        /// </summary>
        Modify_Unused = (byte)0xFE
    }

    /// <summary>
    /// This enumeration is used to set the length of Count field.
    /// </summary>
    public enum Count
    {
        /// <summary>
        /// COUNT is 2-byte
        /// </summary>
        TwoBytesCount,

        /// <summary>
        /// COUNT is 4-byte
        /// </summary>
        FourBytesCount
    }

    /// <summary>
    /// Enum of RestrictType used to create different Restriction
    /// [MS-OXCDATA] 2.12
    /// </summary>
    public enum RestrictionType : byte
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
        PropertyRestriction = 0X04,

        /// <summary>
        /// Logical NOT applied to a subrestriction.
        /// </summary>
        ComparePropertiesRestriction = 0x05,

        /// <summary>
        /// Perform bitwise AND of a property value with a mask and compare to zero.
        /// </summary>
        BitMaskRestriction = 0x06,

        /// <summary>
        /// Compare the Size of a property value to a particular figure.
        /// </summary>
        SizeRestriction = 0X07,

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

    /// <summary>
    /// Relational operator that is used to compare value.
    /// </summary>
    public enum RelationalOperater : byte
    {
        /// <summary>
        /// TRUE if the value of the object's property is less than the given value.
        /// </summary>
        RELOP_LT,

        /// <summary>
        /// TRUE if the value of the object's property is less than or equal to the given value.
        /// </summary>
        RELOP_LE,

        /// <summary>
        /// TRUE if the value of the object's property value is greater than the given value.
        /// </summary>
        RELOP_GT,

        /// <summary>
        /// TRUE if the value of the object's property value is greater than or equal to the given value.
        /// </summary>
        RELOP_GE,

        /// <summary>
        /// TRUE if the object's property value equals the given value.
        /// </summary>
        RELOP_EQ,

        /// <summary>
        /// TRUE if the object's property value does not equal the given value.
        /// </summary>
        RELOP_NE,

        /// <summary>
        /// TRUE if the value of the object's property is in the DL membership of the specified property value. The value of the object's property MUST be an EntryID of a mail-enabled object in the address book. The specified property value MUST be an EntryID of a distribution list object in the address book.
        /// </summary>
        RELOP_MEMBER_OF_DL
    }

    /// <summary>
    /// Rule properties structure.
    /// </summary>
    public struct RuleProperties
    {
        /// <summary>
        /// Name of the rule
        /// </summary>
        public string Name;

        /// <summary>
        /// Rule provider
        /// </summary>
        public string Provider;

        /// <summary>
        /// Rule user flag
        /// </summary>
        public string UserFlag;

        /// <summary>
        /// Rule provider data
        /// </summary>
        public string ProviderData;

        /// <summary>
        /// Subject name used in rule condition
        /// </summary>
        public string ConditionSubjectName;
    }
}