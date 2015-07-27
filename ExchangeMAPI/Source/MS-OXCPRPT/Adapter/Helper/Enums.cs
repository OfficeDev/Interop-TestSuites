//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    /// <summary>
    /// Specifies the GUID in PropertyName.
    /// </summary>
    public enum SpecialGUID
    {
        /// <summary>
        /// GUID is equal to PS-MAPI
        /// </summary>
        PS_MAPI,

        /// <summary>
        /// GUID is non PS-MAPI.
        /// </summary>
        Other
    }

    /// <summary>
    /// Specifies the object type for ROP operations.
    /// </summary>
    public enum ServerObjectType
    {
        /// <summary>
        /// Logon Object
        /// </summary>
        Logon,

        /// <summary>
        /// Folder Object
        /// </summary>
        Folder,

        /// <summary>
        /// Message Object
        /// </summary>
        Message,

        /// <summary>
        /// Attachment Object
        /// </summary>
        Attachment
    }

    /// <summary>
    /// Specifies the SUT property value.
    /// </summary>
    public enum SutPropertyValue
    {
        /// <summary>
        /// First value
        /// </summary>
        First,

        /// <summary>
        /// Second value
        /// </summary>
        Second,

        /// <summary>
        /// Third value
        /// </summary>
        Third,

        /// <summary>
        /// Fourth value
        /// </summary>
        Fourth
    }

    /// <summary>
    /// Specifies the type of PropertyId.
    /// </summary>
    public enum PropertyIdType
    {
        /// <summary>
        /// Property id less than 0x8000
        /// </summary>
        LessThan0x8000,

        /// <summary>
        /// Property ID values that have an associated PropertyName
        /// </summary>
        HaveAssociatedName,

        /// <summary>
        /// Property ID values that do not have an associated PropertyName
        /// </summary>
        NoAssociatedName
    }

    /// <summary>
    /// Specifies which object will be operated.
    /// </summary>
    public enum ObjectToOperate
    {
        /// <summary>
        /// Indicates the first object
        /// </summary>
        FirstObject,

        /// <summary>
        /// Indicates the second object
        /// </summary>
        SecondObject,

        /// <summary>
        /// Indicates the second object
        /// </summary>
        ThirdObject,

        /// <summary>
        /// Indicates the fourth
        /// </summary>
        FourthObject,

        /// <summary>
        /// Indicates the fifth
        /// </summary>
        FifthObject
    }

    /// <summary>
    /// Specifies QueryFlags for [RopQueryNamedProperties] operation.
    /// QueryFlags is a BYTE bit field. Any bits not specified below SHOULD be ignored by the server.
    /// </summary>
    public enum QueryFlags : ushort
    {
        /// <summary>
        /// Values is 0x01, Named properties with a Kind [MS-OXCDATA] of 0x1 MUST NOT be included in the response
        /// </summary>
        NoStrings = 0x01,

        /// <summary>
        /// Values is 0x02, Named properties with a Kind [MS-OXCDATA] of 0x0 MUST NOT be included in the response.
        /// </summary>
        NoIds = 0x02,

        /// <summary>
        /// Other value
        /// </summary>
        OtherValue = 0x03
    }

    /// <summary>
    /// Specifies the CopyFlags parameter in request of [RopCopyProperties] or [RopCopyTo].
    /// </summary>
    [System.Flags]
    public enum CopyFlags
    {
        /// <summary>
        /// None: value is 0x00. If set, the operation should copy and overwrite the destination 
        /// </summary>
        None,

        /// <summary>
        /// Move: value is 0x01. If set, makes the call a move operation rather than a copy operation
        /// </summary>
        Move,

        /// <summary>
        /// NoOverwrite: value is 0x02
        /// If set, any properties being set by RopCopyProperties that 
        /// already have a value on the destination object will not be overwritten; 
        /// otherwise, they are overwritten
        /// </summary>
        NoOverWrite,

        /// <summary>
        /// Move and NoOverwrite are all set: value is 0x03
        /// If set, any properties being set by RopCopyProperties that 
        /// already have a value on the destination object will not be overwritten, 
        /// and properties will be moved from source object to definition object; 
        /// otherwise, they are overwritten
        /// </summary>
        MoveAndNoOverWrite,

        /// <summary>
        ///  NoOverwrite: value is 0x02
        /// If set, means that any properties being set by RopCopyProperties do not have 
        /// a destination properties.
        /// </summary>
        NoOverWriteAndDestPropNull,

        /// <summary>
        /// Other value other than 0x01, 0x02 and 0x03 MUST be ignored by the server
        /// </summary>
        Other
    }

    /// <summary>
    /// Specifies the out parameter for model action QueryNamedProperties,
    /// use to Indicates whether response contain named properties with specific Kind.
    /// </summary>
    public enum ResponseNotContainsSpecificKind
    {
        /// <summary>
        /// Named properties with a Kind [MS-OXCDATA] of 0x1 MUST NOT be included in the response
        /// </summary>
        Kind0x1,

        /// <summary>
        /// Named properties with a Kind [MS-OXCDATA] of 0x0 MUST NOT be included in the response.
        /// </summary>
        Kind0x0,

        /// <summary>
        /// Named properties with all Kind should be included in the response.
        /// </summary>
        KindAll
    }

    /// <summary>
    /// Specifies the particular properties.
    /// </summary>
    public enum SpecificProperty
    {
        /// <summary>
        /// Normal properties that include all information.
        /// </summary>
        Normal,

        /// <summary>
        /// Property without name
        /// </summary>
        WithoutName
    }

    /// <summary>
    /// Specifies the ErrorCode in the response.
    /// </summary>
    public enum CPRPTErrorCode : uint
    {
        /// <summary>
        /// ROPs call are successful
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// Returned in response when properties too large
        /// Indicated by PropertyLSizeLimit parameter or size of response buffer
        /// </summary>
        NotEnoughMemory = 0x8007000E,

        /// <summary>
        /// When processing [RopOpenStream], Indicates whether property tag exist
        /// value 0x8004010F means:
        /// the property tag does not exist for the object and it cannot be created because Create was not specified in OpenModeFlags.
        /// </summary>
        NotFound = 0x8004010F,

        /// <summary>
        /// A request involving multiple properties failed for one or more individual properties, while succeeding overall.
        /// </summary>
        ErrorsReturned = 0x00040380,

        /// <summary>
        /// [In Processing RopSeekStream] If the client requests the seek pointer be moved beyond 2^31 BYTES, 
        /// the server MUST return StreamSeekError
        /// </summary>
        StreamSeekError = 0x80030019,

        /// <summary>
        /// [In Processing RopLockRegionStream] If there are previous locks that are not expired, 
        /// the server MUST return an AccessDenied error.
        /// </summary>
        AccessDenied = 0x80070005,

        /// <summary>
        /// [In Processing RopLockRegionStream] If a session with an expired lock calls any ROP for this Stream object 
        /// that would encounter the locked region, the server MUST return a NetworkError
        /// </summary>        
        NetworkError = 0x80040115,

        /// <summary>
        /// This Error occurs when try to write or commit a stream that open with ReadOnly flag
        /// </summary>
        STG_E_ACCESSDENIED = 0x80030005,

        /// <summary>
        /// When [RopCopyToStream] and Destination Object is not exist
        /// Also possible error returned by RopCopyTo and RopCopyProperties 
        /// </summary>
        NullDestinationObject = 0x00000503,

        /// <summary>
        /// Returned when the source object and destination object are not compatible with each other for the copy operation.
        /// </summary>
        NotSupported = 0x80040102,

        /// <summary>
        /// Returned when If the server reaches this limit of at most 32,767
        /// </summary>
        OutOfMemory = 0x8007000E,

        /// <summary>
        /// When RopGetPropertyIdsFromNames, if the name could not be mapped, return this error
        /// </summary>
        ecWarnWithErrors = 0x00040380,

        /// <summary>
        /// When set read-only properties on Exchange Server 2010, the GeneralFailure will be returned.
        /// </summary>
        GeneralFailure = 0x80004005,

        /// <summary>
        /// If source object directly or indirect contains destination, RopCopyTo return this error
        /// </summary>
        MessageCycle = 0x00000504,

        /// <summary>
        /// The source folder contains the destination folder.
        /// </summary>
        FolderCycle = 0x8004060B,

        /// <summary>
        /// If there is already a sub-object existing in the destination object with the same display name (PidTagDisplayName), RopCopyTo return this error
        /// </summary>
        CollidingNames = 0x80040604,

        /// <summary>
        /// If CopyFlag is not 0x01(Move) or 0x02(NoOverwrite), RopCopyTo return this error
        /// </summary>
        InvalidParameter = 0x80070057,

        /// <summary>
        /// If Origin flag is invalid, RopSeekStream will return this error
        /// </summary>
        StreamInvalidParam = 0x80030057,

        /// <summary>
        /// If the write will exceed the maximum stream size, RopWriteStream will return this error
        /// </summary>
        StreamSizeError = 0x80030070,

        /// <summary>
        /// If the write will exceed the maximum stream size, RopWriteStream will return this error in Exchange 2010
        /// </summary>
        ecTooBig = 0x80040305,

        /// <summary>
        /// Undefined error code
        /// </summary>
        Other
    }

    /// <summary>
    /// Specifies the code page type in response of [RopGetPropertiesSpecific].
    /// </summary>
    public enum SpecificCodePage
    {
        /// <summary>
        /// For properties on Message objects the code page used for strings in MBCS format
        /// MUST be the code page set on the Message object when it was opened if any
        /// </summary>
        SameWithMessage,

        /// <summary>
        /// All other objects the code page used for strings in MBCS format MUST be the code page of the Logon object
        /// </summary>
        SameWithLogon,

        /// <summary>
        /// For properties on Attachment objects the code page used for strings in MBCS format
        /// MUST be the code page set on the parent Message object when it was opened if any
        /// </summary>
        SameWithParentMessage
    }

    /// <summary>
    /// Specifies OpenModeFlags in [RopOpenStream].
    /// </summary>
    public enum OpenModeFlags
    {
        /// <summary>
        /// Open stream for read-only access.
        /// </summary>
        ReadOnly,

        /// <summary>
        /// Open stream for read/write access
        /// </summary>
        ReadWrite,

        /// <summary>
        /// Opens new stream, this will delete the current property value and open stream for read/write access. 
        /// This is required to open a property that has not been set
        /// </summary>
        Create,

        /// <summary>
        /// If the Folder object, Attachment object, or Message object was opened with read/write access, 
        /// then the stream MUST be opened with read/write access. Otherwise, the stream MUST be opened with read-only access.
        /// </summary>
        BestAccess
    }

    /// <summary>
    /// Specifies a specific type of PropertyName.
    /// </summary>
    public enum SpecificPropertyName
    {
        /// <summary>
        /// GUID is equal to PS-MAPI and Kind is 0x00
        /// </summary>
        PS_MAPIAndKind0x00,

        /// <summary>
        /// Indicates an invalid PropertyId
        /// </summary>
        PS_MAPIAndKind0x01,

        /// <summary>
        /// Kind is 0x01 and GUID NOT equals PS-MAPI
        /// </summary>
        Kind0x01,

        /// <summary>
        /// No constraint for PropertyName
        /// </summary>
        NoConstraint
    }

    /// <summary>
    /// Specifies the nine Common Object Properties defined in section 2.2.1.
    /// </summary>
    public enum CommonObjectProperty
    {
        /// <summary>
        /// Indicates the operations available to the client for the object
        /// </summary>
        PidTagAccess,

        /// <summary>
        /// Indicates the client's access level to the object
        /// </summary>
        PidTagAccessLevel,

        /// <summary>
        /// Contains a global identifier (GID) indicating the last change to the object [MS-OXCFXICS]
        /// </summary>
        PidTagChangeKey,

        /// <summary>
        /// Contains the time the object was created in UTC.
        /// </summary>
        PidTagCreationTime,

        /// <summary>
        /// Contains the name of the last mail user to modify the object
        /// </summary>
        PidTagLastModifierName,

        /// <summary>
        /// Contains the time of the last modification to the object in UTC
        /// </summary>
        PidTagLastModificationTime,

        /// <summary>
        /// Indicates the type of Server object
        /// </summary>
        PidTagObjectType,

        /// <summary>
        /// Contains a unique binary-comparable identifier for a specific object
        /// </summary>
        PidTagRecordKey,

        /// <summary>
        /// Contains a unique binary-comparable key that identifies an object for a search
        /// </summary>
        PidTagSearchKey,

        /// <summary>
        /// Contains the name of the attachment as input by the end user
        /// </summary>
        PidTagDisplayName,

        /// <summary>
        /// Contains the FID of the folder.
        /// </summary>
        PidTagFolderId,
    }

    /// <summary>
    /// Specifies the pre-state before call [RopLockRegionStream].
    /// </summary>
    public enum PreStateBeforeLock
    {
        /// <summary>
        /// Special state that a session with an expired lock calls any ROP for this Stream object
        /// </summary>
        WithExpiredLock,

        /// <summary>
        /// Special state that previous locks that are not expired
        /// </summary>
        PreLockNotExpired,

        /// <summary>
        /// The normal state
        /// </summary>
        Normal
    }

    /// <summary>
    /// Specifies a specific state for RopCopyTo.
    /// </summary>
    public enum CopyToCondition
    {
        /// <summary>
        /// Source object and destination object are compatible
        /// </summary>
        SourceDestNotCompatible,

        /// <summary>
        /// Source object directly or indirectly contains destination object
        /// </summary>
        SourceContainsDest,

        /// <summary>
        /// source message directly contains destination message
        /// </summary>
        SourceMessageContainsDestMessage,

        /// <summary>
        /// source message indirectly contains destination message
        /// </summary>
        SourceMessageIndirectlyContainsDestMessage,
        
        /// <summary>
        /// There is already a sub-object existing in the destination object with the same display name (PidTagDisplayName) and CopyFlag is 0x01
        /// </summary>
        SourceDestHasSubObjWithSameDisplayName,

        /// <summary>
        /// Normal state for RopCopyTo
        /// </summary>
        Normal
    }
        
    /// <summary>
    /// Specifies particular scenario when RopSeekStream.
    /// </summary>
    public enum SeekStreamCondition
    {
        /// <summary>
        /// Seek to offset before the start, or beyond the max stream size of 2^31
        /// </summary>
        MovedBeyondMaxStreamSize,

        /// <summary>
        /// The seek pointer try to move beyond the end of the stream
        /// </summary>
        MovedBeyondEndOfStream,

        /// <summary>
        /// The value of Origin parameter is invalid
        /// </summary>
        OriginInvalid,

        /// <summary>
        /// Normal state for seek stream
        /// </summary>
        Normal
    }

    /// <summary>
    /// Specifies the TaggedProperty name.
    /// </summary>
    public enum TaggedPropertyName : ushort
    {
        /// <summary>
        /// PropertyId of PidTagDisplayName
        /// </summary>
        PidTagDisplayName = 0x3001,

        /// <summary>
        /// PropertyId of PidTagFolderId
        /// </summary>
        PidTagFolderId = 0x6748,

        /// <summary>
        /// PropertyId of PidTagAccess
        /// </summary>
        PidTagAccess = 0x0FF4,

        /// <summary>
        /// PropertyId of PidTagAccessLevel
        /// </summary>
        PidTagAccessLevel = 0x0FF7,

        /// <summary>
        ///  PropertyId of PidTagChangeKey
        /// </summary>
        PidTagChangeKey = 0x65E2,

        /// <summary>
        ///  PropertyId of PidTagCreationTime
        /// </summary>
        PidTagCreationTime = 0x3007,

        /// <summary>
        /// PropertyId of PidTagLastModifierName
        /// </summary>
        PidTagLastModifierName = 0x3FFA,

        /// <summary>
        /// PropertyId of PidTagLastModificationTime
        /// </summary>
        PidTagLastModificationTime = 0x3008,

        /// <summary>
        /// PropertyId of PidTagRecordKey
        /// </summary>
        PidTagRecordKey = 0x0FF9,

        /// <summary>
        /// PropertyId of PidTagSearchKey
        /// </summary>
        PidTagSearchKey = 0x300B,

        /// <summary>
        /// PropertyId of PidTagContentCount
        /// </summary>
        PidTagContentCount = 0x3602,

        /// <summary>
        /// PropertyId of PidTagEntryId
        /// </summary>
        PidTagEntryId = 0x0FFF,
 
        /// <summary>
        /// PropertyId of PidTagMessageFlags
        /// </summary>
        PidTagMessageFlags = 0x0E07,

        /// <summary>
        /// PropertyId of PidTagMessageClass
        /// </summary>
        PidTagMessageClass = 0x001a,

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
        /// PropertyId of PidTagAttachMethod
        /// </summary>
        PidTagAttachMethod = 0x3705,

        /// <summary>
        /// PropertyId of PidTagConversationId
        /// </summary>
        PidTagConversationId = 0x3013
    }

    /// <summary>
    /// Specifies the PropertyType name.
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
        PtypRestriction = 0x00FD,

        /// <summary>
        /// PropertyType of PtypTime
        /// </summary>
        PtypTime = 0x0040,

        /// <summary>
        /// PropertyType of PtypObject
        /// </summary>
        PtypObject = 0x000D
    }
}