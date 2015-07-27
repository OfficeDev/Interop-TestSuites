//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

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
    /// The enumeration of PropertyTypeName
    /// </summary>
    public enum PropertyType : ushort
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
        PtypRestriction = 0x00FD,

        /// <summary>
        /// 2 bytes, a 16-bit integer.
        /// </summary>
        PtypInteger16 = 0x00000002,

        /// <summary>
        /// 16 bytes, a GUID with Data1, Data2, and Data3 fields in little-endian format.
        /// </summary>
        PtypGuid = 0x00000048,
        
        /// <summary>
        /// 8 bytes, a 64-bit integer representing 
        /// the number of 100-nanosecond intervals since January 1, 1601.
        /// </summary>
        PtypTime = 0x00000040,
        
        /// <summary>
        /// 4 bytes, a 32-bit integer encoding error information.
        /// </summary>
        PtypErrorCode = 0x0000000A,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypInteger16 values.
        /// </summary>
        PtypMultipleInteger16 = 0x00001002,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypInteger32 values.
        /// </summary>
        PtypMultipleInteger32 = 0x00001003,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypString8 values.
        /// </summary>
        PtypMultipleString8 = 0x0000101E,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypBinary values.
        /// </summary>
        PtypMultipleBinary = 0x00001102,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that PtypString values.
        /// </summary>
        PtypMultipleString = 0x0000101F,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypGuid values.
        /// </summary>
        PtypMultipleGuid = 0x00001048,
        
        /// <summary>
        /// Variable Size, a COUNT followed by that many PtypTime values.
        /// </summary>
        PtypMultipleSystime = 0x00001040,
        
        /// <summary>
        /// Single 32-bit value, referencing an address list.
        /// </summary>
        PtypEmbeddedTable = 0x0000000D,
        
        /// <summary>
        /// Clients MUST NOT specify this Property Type in any method's input parameters.
        /// The server MUST specify this Property Type in any method's output parameters to 
        /// indicate that a property has a value that cannot be expressed in the NSPI Protocol.
        /// </summary>
        PtypNull = 0x00000001,
        
        /// <summary>
        /// Clients specify this Property Type in a method's input parameter to indicate that the
        /// client will accept any Property Type the server chooses when returning propvalues.
        /// </summary>
        PtypUnspecified = 0x00000000
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