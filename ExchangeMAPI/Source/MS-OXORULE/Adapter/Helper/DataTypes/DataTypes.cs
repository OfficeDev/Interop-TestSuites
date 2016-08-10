namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;

    #region Enums
    /// <summary>
    /// ActionType which used in ActionData of a RuleAction
    /// </summary>
    public enum ActionType : byte
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
    /// BounceCode (4 bytes): Specifies a bounce code
    /// The BounceCode field MUST have one of the following values
    /// [MS-OXORULE] 2.2.5.1.3.5
    /// </summary>
    public enum BounceCode : uint
    {
        /// <summary>
        /// The message was refused because it was too large
        /// </summary>
        TooLarge = 0x0000000D,

        /// <summary>
        /// The message was refused because it cannot be displayed to the user
        /// </summary>
        CanNotDisplay = 0x0000001F,

        /// <summary>
        /// The message delivery was denied for other reasons
        /// </summary>
        Denied = 0x00000026
    }

    /// <summary>
    /// Type of Store object
    /// </summary>
    public enum StoreObjectType
    {
        /// <summary>
        /// Store is mailbox Store
        /// </summary>
        Mailbox,

        /// <summary>
        /// Store is public folder Store
        /// </summary>
        PublicFolder
    }

    /// <summary>
    /// This enumeration is used to set the length of COUNT field
    /// </summary>
    public enum CountByte
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
    /// RuleState in PidTagRuleState of RuleData
    /// </summary>
    [FlagsAttribute]
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
    /// Specifies test account in system under test
    /// </summary>
    public enum TestUser
    {
        /// <summary>
        /// TestUser1 account of test server which has owner level permission of the public folder 
        /// </summary>
        TestUser1,

        /// <summary>
        /// A TestUser2 that can deliver messages
        /// </summary>
        TestUser2
    }

    /// <summary>
    /// Specifies different RuleData which generated from ptf config for different test purpose
    /// </summary>
    public enum TestRuleDataType : byte
    {
        /// <summary>
        /// Test data which will be used for testing add a new rule
        /// </summary>
        ForAdd = 0x01,

        /// <summary>
        /// Test data which will be used for testing modify an existing rule
        /// </summary>
        ForModify = 0x02,

        /// <summary>
        /// Test data which will be used for testing remove an existing rule
        /// </summary>
        ForRemove = 0x04,
    }

    /// <summary>
    /// The Enumeration of PidTagMessageFlags
    /// </summary>
    public enum PidTagMessageFlag : uint
    {
        /// <summary>
        /// It is a FAI message
        /// </summary>
        mfFAI = 0x00000040
    }

    /// <summary>
    /// Property IDs for some property.
    /// </summary>
    public enum PropertyId : ushort
    {
        /// <summary>
        /// PropertyId of PidTagHasAttachment
        /// </summary>
        pidTagHasAttachment = 0x0E1B,

        /// <summary>
        /// PropertyId of PidTagReceivedByEmailAddress
        /// </summary>
        PidTagReceivedByEmailAddress = 0x0076,

        /// <summary>
        /// PropertyId of PidTagContentFilterSpamConfidenceLevel
        /// </summary>
        PidTagContentFilterSpamConfidenceLevel = 0x4076,

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
        /// PropertyId of PidTagRwRulesStream
        /// </summary>
        PidTagRwRulesStream = 0x6802,

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
    /// Action Flavor values of OP_FORWARD
    /// </summary>
    [FlagsAttribute]
    public enum ActionFlavorsForward : uint
    {
        /// <summary>
        /// Preserves the sender information and indicates that the message was auto-forwarded. Can be combined with the NC ActionFlavor flag.
        /// </summary>
        PR = 0x00000001,

        /// <summary>
        /// Preserves the sender information and indicates that the message was auto-forwarded. Can be combined with the NC ActionFlavor flag.
        /// </summary>
        NC = 0x00000002,

        /// <summary>
        /// Makes the message an attachment to the forwarded message. This value MUST NOT be combined with other ActionFlavor Flags.
        /// </summary>
        AT = 0x00000004,

        /// <summary>
        /// Indicates that the message SHOULD be forwarded as a Short Message Service (SMS) text message. This value MUST NOT be combined with other ActionFlavor Flags.
        /// </summary>
        TM = 0x00000008,

        /// <summary>
        /// Unused. This bit MUST be set to 0 by the client and ignored by the server.
        /// </summary>
        x = 0xfffffff0
    }

    /// <summary>
    /// Action Flavor values of OP_REPLY or OP_OOF_REPLY
    /// </summary>
    public enum ActionFlavorsReply
    {
        /// <summary>
        /// Do not send the message to the message sender (the reply template MUST contain recipients in this case).
        /// </summary>
        NS = 0x00000001,

        /// <summary>
        /// Server will use fixed, server-defined text in the reply message and ignore the text in the reply template. This text is an implementation detail.
        /// </summary>
        ST = 0x00000002
    }

    /// <summary>
    /// This enum is used to specify the ROP operation is performed on  ExtendedRule, DAM, DEM or StandardRule
    /// </summary>
    public enum TargetOfRop : uint
    {
        /// <summary>
        /// ROP operation for other
        /// </summary>
        OtherTarget,

        /// <summary>
        /// ROP operation for ExtendedRules
        /// </summary>
        ForExtendedRules,

        /// <summary>
        /// ROP operation for DAM
        /// </summary>
        ForDAM,

        /// <summary>
        /// ROP operation for DEM
        /// </summary>
        ForDEM,

        /// <summary>
        /// ROP operation for StandardRules
        /// </summary>
        ForStandardRules,

        /// <summary>
        /// ROP operation for Template messages.
        /// </summary>
        ForTemplateMessage
    }

    /// <summary>
    /// This enum used to specify the modify rule flag used in modify rule ROP operation.
    /// </summary>
    public enum ModifyRuleFlag : byte
    {
        /// <summary>
        /// If this bit is set, the rules  in this request are to replace the existing set of rules  in the folder;
        /// in this case, all subsequent RuleData structures  MUST have ROW_ADD as the value of their RuleDataFlag field.
        /// </summary>
        Modify_ReplaceAll = (byte)0x01,

        /// <summary>
        /// The rules specified in this request represent changes (delete, modify, and add) to the set of rules already existing in this folder.
        /// </summary>
        Modify_OnExisting = (byte)0x00,

        /// <summary>
        /// Beside the last bit, all other bit is unused, they are ignored by the server if set
        /// </summary>
        Modify_Unused = (byte)0xFE
    }

    /// <summary>
    /// This enum is used in ROP submit message operation. 
    /// </summary>
    public enum SubmitFlag : uint
    {
        /// <summary>
        /// The value of none.
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

    /// <summary>
    /// TableFlags parameter contains a bitmask of Flags that control how information is returned in the table.
    /// </summary>
    public enum ContentTableFlag : uint
    {
        /// <summary>
        /// No Flags enabled.
        /// </summary>
        None = 0x00,

        /// <summary>
        /// Requests an FAI table instead of a standard table.
        /// </summary>
        Associated = 0x02,

        /// <summary>
        /// The call can return immediately, possibly before the ROP execution is complete and in this case the ReturnValue and the RowCount fields 
        /// in the return buffer might not be accurate.
        /// </summary>
        DeferredErrors = 0x08,

        /// <summary>
        /// Disables table notifications to the client.
        /// </summary>
        NoNotifications = 0x10,

        /// <summary>
        /// Enables the client to get a list of the soft deleted messages in a folder.
        /// </summary>
        SoftDeletes = 0x20,

        /// <summary>
        /// Requests that the columns that contain string data of unspecified Type be returned in Unicode format.
        /// </summary>
        UseUnicode = 0x40,

        /// <summary>
        /// Retrieves all of the messages pertaining to a single conversation. One result row represents a single message.
        /// </summary>
        ConversationMembers = 0x80
    }

    /// <summary>
    /// This value specifies whether to return string properties in Unicode.
    /// </summary>
    public enum WantUnicode : ushort
    {
        /// <summary>
        /// Return string properties in unicode.
        /// </summary>
        Want = 0x01,

        /// <summary>
        ///  Return string properties not in unicode.
        /// </summary>
        NotWant = 0x00
    }

    /// <summary>
    /// The request type of MS-OXCMAPIHTTP
    /// </summary>
    public enum RequestType
    {
        /// <summary>
        /// The connect request type
        /// </summary>
        Connect,

        /// <summary>
        /// The Execute request type
        /// </summary>
        Execute,

        /// <summary>
        /// The Disconnect type
        /// </summary>
        Disconnect,

        /// <summary>
        /// The NotificationWait request type
        /// </summary>
        NotificationWait,

        /// <summary>
        /// The PING request type
        /// </summary>
        PING,

        /// <summary>
        /// The Bind request type
        /// </summary>
        Bind,

        /// <summary>
        /// The Unbind request type
        /// </summary>
        Unbind,

        /// <summary>
        /// The CompareMIds request type
        /// </summary>
        CompareMIds,

        /// <summary>
        /// The DNToMId request type
        /// </summary>
        DNToMId,

        /// <summary>
        /// The GetMatches request type
        /// </summary>
        GetMatches,

        /// <summary>
        /// The GetPropList request type
        /// </summary>
        GetPropList,

        /// <summary>
        /// The GetProps request type
        /// </summary>
        GetProps,

        /// <summary>
        /// The GetSpecialTable request type
        /// </summary>
        GetSpecialTable,

        /// <summary>
        /// The GetTemplateInfo request type
        /// </summary>
        GetTemplateInfo,

        /// <summary>
        /// The ModLinkAtt request type
        /// </summary>
        ModLinkAtt,

        /// <summary>
        /// The ModProps request type
        /// </summary>
        ModProps,

        /// <summary>
        /// The QueryColumns request type
        /// </summary>
        QueryColumns,

        /// <summary>
        /// The QueryRows request type
        /// </summary>
        QueryRows,

        /// <summary>
        /// The ResolveNames request type
        /// </summary>
        ResolveNames,

        /// <summary>
        /// The ResortRestriction request type
        /// </summary>
        ResortRestriction,

        /// <summary>
        /// The SeekEntries request type
        /// </summary>
        SeekEntries,

        /// <summary>
        /// The UpdateStat request type
        /// </summary>
        UpdateStat,

        /// <summary>
        /// The GetMailboxUrl request type
        /// </summary>
        GetMailboxUrl,

        /// <summary>
        /// The GetAddressBookUrl request type
        /// </summary>
        GetAddressBookUrl
    }

    /// <summary>
    /// The flag value of NspiBind method.
    /// </summary>
    public enum NspiBindFlag : uint
    {
        /// <summary>
        /// Indicate that the server does not validate that the client is an authenticated user. Now this value is defined in MS-NSPI.
        /// </summary>
        fAnonymousLogin = 0x00000020,
    }

    /// <summary>
    /// The property type values are used to specify property types.
    /// </summary>
    public enum PropertyTypeValue : uint
    {
        /// <summary>
        /// 2 bytes, a 16-bit integer.
        /// </summary>
        PtypInteger16 = 0x00000002,

        /// <summary>
        /// 4 bytes, a 32-bit integer.
        /// </summary>
        PtypInteger32 = 0x00000003,

        /// <summary>
        /// 1 byte, restricted to 1 or 0.
        /// </summary>
        PtypBoolean = 0x0000000B,

        /// <summary>
        /// Variable size, a string of multi-byte characters in externally specified encoding with terminating null character (single 0 byte).
        /// </summary>
        PtypString8 = 0x0000001E,

        /// <summary>
        /// Variable size, a COUNT followed by that many bytes.
        /// </summary>
        PtypBinary = 0x00000102,

        /// <summary>
        /// Variable size, a string of Unicode characters in UTF-16LE encoding with terminating null character (2 bytes of zero).
        /// </summary>
        PtypString = 0x0000001F,

        /// <summary>
        /// 16 bytes, a GUID with Data1, Data2, and Data3 fields in little-endian format.
        /// </summary>
        PtypGuid = 0x00000048,

        /// <summary>
        /// 8 bytes, a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601.
        /// </summary>
        PtypTime = 0x00000040,

        /// <summary>
        /// 4 bytes, a 32-bit integer encoding error information.
        /// </summary>
        PtypErrorCode = 0x0000000A,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger16 values.
        /// </summary>
        PtypMultipleInteger16 = 0x00001002,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypInteger32 values.
        /// </summary>
        PtypMultipleInteger32 = 0x00001003,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypString8 values.
        /// </summary>
        PtypMultipleString8 = 0x0000101E,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypBinary values.
        /// </summary>
        PtypMultipleBinary = 0x00001102,

        /// <summary>
        /// Variable size, a COUNT followed by that PtypString values.
        /// </summary>
        PtypMultipleString = 0x0000101F,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypGuid values.
        /// </summary>
        PtypMultipleGuid = 0x00001048,

        /// <summary>
        /// Variable size, a COUNT followed by that many PtypTime values.
        /// </summary>
        PtypMultipleTime = 0x00001040,

        /// <summary>
        /// Single 32-bit value, referencing an address list. 
        /// </summary>
        PtypEmbeddedTable = 0x0000000D,

        /// <summary>
        /// Clients MUST NOT specify this property type in any method's input parameters.
        /// The server MUST specify this property type in any method's output parameters to indicate that a property has a value that cannot be expressed in the Exchange Server NSPI Protocol.
        /// </summary>
        PtypNull = 0x00000001,

        /// <summary>
        /// Clients specify this property type in a method's input parameter to indicate that the client will accept any property type the server chooses when returning propvalues.
        /// Servers MUST NOT specify this property type in any method's output parameters except the method NspiGetIDsFromNames.
        /// </summary>
        PtypUnspecified = 0x00000000
    }

    /// <summary>
    /// The values are used to specify display types. 
    /// </summary>
    public enum DisplayTypeValue : uint
    {
        /// <summary>
        /// A typical messaging user.
        /// </summary>
        DT_MAILUSER = 0x00000000,

        /// <summary>
        /// A distribution list.
        /// </summary>
        DT_DISTLIST = 0x00000001,

        /// <summary>
        /// A forum, such as a bulletin board service or a public or shared folder.
        /// </summary>
        DT_FORUM = 0x00000002,

        /// <summary>
        /// An automated agent, such as Quote-Of-The-Day or a weather chart display.
        /// </summary>
        DT_AGENT = 0x00000003,

        /// <summary>
        /// An Address Book object defined for a large group, such as helpdesk, accounting, coordinator, 
        /// or department. Department objects usually have this display type.
        /// </summary>
        DT_ORGANIZATION = 0x00000004,

        /// <summary>
        /// A private, personally administered distribution list.
        /// </summary>
        DT_PRIVATE_DISTLIST = 0x00000005,

        /// <summary>
        /// An Address Book object known to be from a foreign or remote messaging system.
        /// </summary>
        DT_REMOTE_MAILUSER = 0x00000006,

        /// <summary>
        /// An address book hierarchy table container. 
        /// An Exchange NSPI server MUST NOT return this display type except as part of an EntryID of an object in the address book hierarchy table.
        /// </summary>
        DT_CONTAINER = 0x00000100,

        /// <summary>
        /// A display template object. An Exchange NSPI server MUST NOT return this display type.
        /// </summary>
        DT_TEMPLATE = 0x00000101,

        /// <summary>
        /// An address creation template. 
        /// An Exchange NSPI server MUST NOT return this display type except as part of an EntryID of an object in the Address Creation Table.
        /// </summary>
        DT_ADDRESS_TEMPLATE = 0x00000102,

        /// <summary>
        /// A search template. An Exchange NSPI server MUST NOT return this display type. 
        /// </summary>
        DT_SEARCH = 0x00000200
    }

    /// <summary>
    /// The language code identifier (LCID) specified in this section is associated with the minimal required sort order for Unicode strings. 
    /// </summary>
    public enum DefaultLCID
    {
        /// <summary>
        /// Represents the default LCID that is used for comparison of Unicode string representations.
        /// </summary>
        NSPI_DEFAULT_LOCALE = 0x00000409,
    }

    /// <summary>
    /// The required code pages listed in this section are associated with the string handling in the Exchange Server NSPI Protocol, 
    /// and they appear in input parameters to methods in the Exchange Server NSPI Protocol. 
    /// </summary>
    public enum RequiredCodePage : uint
    {
        /// <summary>
        /// Represents the Teletex code page.
        /// </summary>
        CP_TELETEX = 0x00004F25,

        /// <summary>
        /// Represents the Unicode code page.
        /// </summary>
        CP_WINUNICODE = 0x000004B0,
    }

    /// <summary>
    /// The positioning Minimal Entry IDs are used to specify objects in the address book as a function of their positions in tables.
    /// </summary>
    public enum MinimalEntryID
    {
        /// <summary>
        /// Specifies the position before the first row in the current address book container.
        /// </summary>
        MID_BEGINNING_OF_TABLE = 0x00000000,

        /// <summary>
        /// Specifies the position after the last row in the current address book container.
        /// </summary>
        MID_END_OF_TABLE = 0x00000002,

        /// <summary>
        /// Specifies the current position in a table. This Minimal Entry ID is only valid in the NspiUpdateStat method. 
        /// In all other cases, it is an invalid Minimal Entry ID, guaranteed to not specify any object in the address book.
        /// </summary>
        MID_CURRENT = 0x00000001,
    }

    /// <summary>
    /// Ambiguous name resolution (ANR) Minimal Entry IDs are used to specify the outcome of the ANR process. 
    /// </summary>
    public enum ANRMinEntryID
    {
        /// <summary>
        /// The ANR process is unable to map a string to any objects in the address book.
        /// </summary>
        MID_UNRESOLVED = 0x00000000,

        /// <summary>
        /// The ANR process maps a string to multiple objects in the address book.
        /// </summary>
        MID_AMBIGUOUS = 0x0000001,

        /// <summary>
        /// The ANR process maps a string to a single object in the address book.
        /// </summary>
        MID_RESOLVED = 0x0000002,
    }

    /// <summary>
    /// The values are used to specify a specific sort orders for tables. 
    /// </summary>
    public enum TableSortOrder
    {
        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property, as specified in [MS-OXCFOLD] section 2.2.2.2.2.3. 
        /// All Exchange NSPI servers MUST support this sort order for at least one LCID.
        /// </summary>
        SortTypeDisplayName = 0x00000000,

        /// <summary>
        /// The table is sorted ascending on the PidTagAddressBookPhoneticDisplayName property, as specified in [MS-OXOABK] section 2.2.3.9. 
        /// Exchange NSPI servers SHOULD support this sort order. Exchange NSPI servers MAY support this only for some LCIDs.
        /// </summary>
        SortTypePhoneticDisplayName = 0x00000003,

        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property. 
        /// The client MUST set this value only when using the NspiGetMatches method to open a non-writable table on an object-valued property.
        /// </summary>
        SortTypeDisplayName_RO = 0x000003E8,

        /// <summary>
        /// The table is sorted ascending on the PidTagDisplayName property. 
        /// The client MUST set this value only when using the NspiGetMatches method to open a writable table on an object-valued property.
        /// </summary>
        SortTypeDisplayName_W = 0x000003E9,
    }

    /// <summary>
    /// The property flag values that are used as bit flags in NspiGetPropList, NspiGetProps, and NspiQueryRows methods to specify optional behavior to a server.
    /// </summary>
    public enum RetrievePropertyFlag
    {
        /// <summary>
        /// Client requires that the server MUST NOT include proptags with the PtypEmbeddedTable property type 
        /// in any lists of proptags that the server creates on behalf of the client.
        /// </summary>
        fSkipObjects = 0x00000001,

        /// <summary>
        /// Client requires that the server MUST return Entry ID values in Ephemeral Entry ID form.
        /// </summary>
        fEphID = 0x00000002,
    }

    /// <summary>
    /// NspiGetSpecialTable flag values are used as bit flags in the NspiGetSpecialTable method to specify optional behavior to a server. 
    /// </summary>
    [FlagsAttribute]
    public enum NspiGetSpecialTableFlags
    {
        /// <summary>
        /// Specify none to 0.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// Specify that the server MUST return the table of the available address creation templates. 
        /// Specify that this flag causes the server to ignore the NspiUnicodeStrings flag.
        /// </summary>
        NspiAddressCreationTemplates = 0x00000002,

        /// <summary>
        /// Specifies that the server MUST return all strings as Unicode representations 
        /// rather than as multibyte strings in the client's code page. 
        /// </summary>
        NspiUnicodeStrings = 0x00000004,
    }

    /// <summary>
    /// The NspiQueryColumns flag value is used as a bit flag in the NspiQueryColumns method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiQueryColumnsFlag : uint
    {
        /// <summary>
        /// Specifies that the server MUST return all proptags that specify values with string 
        /// representations as having the PtypString property type.
        /// </summary>
        NspiUnicodeProptypes = 0x80000000,
    }

    /// <summary>
    /// The NspiGetTemplateInfo flag values are used as bit flags in the NspiGetTemplateInfo method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiGetTemplateInfoFlag
    {
        /// <summary>
        /// Specifies that the server is to return the value that represents a template.
        /// </summary>
        TI_TEMPLATE = 0x00000001,

        /// <summary>
        /// Specifies that the server is to return the value of the script that is associated with a template.
        /// </summary>
        TI_SCRIPT = 0x00000004,

        /// <summary>
        /// Specifies that the server is to return the e-mail type that is associated with a template.
        /// </summary>
        TI_EMT = 0x00000010,

        /// <summary>
        /// Specifies that the server is to return the name of the help file that is associated with a template.
        /// </summary>
        TI_HELPFILE_NAME = 0x00000020,

        /// <summary>
        /// Specifies that the server is to return the contents of the help file that is associated with a template.
        /// </summary>
        TI_HELPFILE_CONTENTS = 0x00000040,
    }

    /// <summary>
    /// The NspiModLinkAtt flag value is used as a bit flag in the NspiModLinkAtt method to specify optional behavior to a server. 
    /// </summary>
    public enum NspiModLinkAtFlag
    {
        /// <summary>
        /// Specify that the server is to remove values when modifying. 
        /// </summary>
        fDelete = 0x00000001,
    }

    /// <summary>
    /// The property tags with the property type.
    /// </summary>
    public enum AulProp : uint
    {
        /// <summary>
        /// The property tag of PidTagEntryId.
        /// </summary>
        PidTagEntryId = 0x0FFF0102,

        /// <summary>
        /// The property tag of PidTagAddressBookDisplayNamePrintable.
        /// </summary>
        PidTagAddressBookDisplayNamePrintable = 0x39FE001F,

        /// <summary>
        /// The property tag of PidTagTitle.
        /// </summary>
        PidTagTitle = 0x3A17001F,

        /// <summary>
        /// The property tag of PidTagTitle.
        /// </summary>
        PidTagAddressBookContainerId = 0xFFFD0003,

        /// <summary>
        /// The property tag of PidTagObjectType.
        /// </summary>
        PidTagObjectType = 0x0ffe0003,

        /// <summary>
        /// The property tag of PidTagDisplayType.
        /// </summary>
        PidTagDisplayType = 0x39000003,

        /// <summary>
        /// The property tag of PidTagDisplayName with the Property Type PtypString8.
        /// </summary>
        PidTagDisplayName = 0x3001001e,

        /// <summary>
        /// The property tag of PidTagPrimaryTelephoneNumber with the Property Type PtypString8.
        /// </summary>
        PidTagPrimaryTelephoneNumber = 0x3a1a001e,

        /// <summary>
        /// The property tag of PidTagDepartmentName with the Property Type PtypString8.
        /// </summary>
        PidTagDepartmentName = 0x3a18001e,

        /// <summary>
        /// The property tag of PidTagOfficeLocation with the Property Type PtypString8.
        /// </summary>
        PidTagOfficeLocation = 0x3a19001e,

        /// <summary>
        /// The property tag of PidTagUserX509Certificate with the Property Type PtypMultipleBinary.
        /// </summary>
        PidTagUserX509Certificate = 0x3a701102,

        /// <summary>
        /// The property tag of PidTagAddressBookX509Certificate with the Property Type PtypMultipleBinary
        /// </summary>
        PidTagAddressBookX509Certificate = 0x8c6a1102,

        /// <summary>
        /// The property tag of PidTagAddressBookMember with the Property Type PtypMultipleString8
        /// </summary>
        PidTagAddressBookMember = 0x8009101e,

        /// <summary>
        /// The property tag of PidTagAddressBookPublicDelegates with the Property Type PtypComObject.
        /// </summary>
        PidTagAddressBookPublicDelegates = 0x8015101e,

        /// <summary>
        /// The property tag of PidTagInstanceKey with the Property Type PtypBinary
        /// </summary>
        PidTagInstanceKey = 0x0FF60102,

        /// <summary>
        /// The property tag of PidTagAddressType with the Property Type PtypString8.
        /// </summary>
        PidTagAddressType = 0x3002001E,

        /// <summary>
        /// The property tag of PidTagDepth with the Property Type PtypInteger32.
        /// </summary>
        PidTagDepth = 0x30050003,

        /// <summary>
        /// The property tag of PidTagSelectable with the Property Type PtypBoolean.
        /// </summary>
        PidTagSelectable = 0x3609000B,

        /// <summary>
        /// The property tag of PidTagTemplateData with the Property Type PtypBinary.
        /// </summary>
        PidTagTemplateData = 0x00010102,

        /// <summary>
        /// The property tag of PidTagScriptData with the Property Type PtypBinary.
        /// </summary>
        PidTagScriptData = 0x00040102,

        /// <summary>
        /// The property tag of PidTagContainerContents with the Property Type PtypComObject.
        /// </summary>
        PidTagContainerContents = 0x360f000d,

        /// <summary>
        /// The property tag of PidTagContainerFlags with the Property Type PtypInteger32.
        /// </summary>
        PidTagContainerFlags = 0x36000003,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypInteger32.
        /// </summary>
        PidTagInitialDetailsPane = 0x3f080003,

        /// <summary>
        /// The property tag of PidTagSearchKey with the Property Type PtypBinary.
        /// </summary>
        PidTagSearchKey = 0x300b0102,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypBinary.
        /// </summary>
        PidTagRecordKey = 0xff90102,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypString.
        /// </summary>
        PidTagEmailAddress = 0x3003001f,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypBinary.
        /// </summary>
        PidTagTemplateid = 0x39020102,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypString.
        /// </summary>
        PidTagTransmittableDisplayName = 0x3a20001f,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypBinary.
        /// </summary>
        PidTagMappingSignature = 0x0ff80102,

        /// <summary>
        /// The property tag of PidTagInitialDetailsPane with the Property Type PtypString.
        /// </summary>
        PidTagAddressBookObjectDistinguishedName = 0x803c001f,

        /// <summary>
        /// The property tag of PidTagAddressBookPhoneticDisplayName with the Property Type PtypString.
        /// </summary>
        PidTagAddressBookPhoneticDisplayName = 0x8c92001f,

        /// <summary>
        /// The property tag of PidTagAddressBookPhoneticDisplayName with the Property Type PtypBoolean.
        /// </summary>
        PidTagAddressBookIsMaster = 0xFFFB000B,

        /// <summary>
        /// The property tag of PidTagAddressBookParentEntryId with the Property Type PtypBinary.
        /// </summary>
        PidTagAddressBookParentEntryId = 0xFFFC0102
    }
    #endregion

    /// <summary>
    /// Rule properties struct.
    /// </summary>
    public struct RuleProperties
    {
        /// <summary>
        /// Name of the rule.
        /// </summary>
        public string Name;
        
        /// <summary>
        /// Rule provider
        /// </summary>
        public string Provider;
        
        /// <summary>
        /// Rule user flag.
        /// </summary>
        public string UserFlag;
        
        /// <summary>
        /// Rule provider data.
        /// </summary>
        public string ProviderData;
        
        /// <summary>
        /// Subject name used in rule condition.
        /// </summary>
        public string ConditionSubjectName;
    }
}