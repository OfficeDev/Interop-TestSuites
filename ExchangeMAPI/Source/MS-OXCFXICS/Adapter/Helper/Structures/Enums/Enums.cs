//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// The final ICS state of the synchronization download operation.
    /// </summary>
    public enum ICSStateProperties 
    {
        /// <summary>
        /// The PidTagIdsetGiven.
        /// </summary>
        PidTagIdsetGiven,

        /// <summary>
        /// The PidTagCnsetSeen.
        /// </summary>
        PidTagCnsetSeen,

        /// <summary>
        /// The PidTagCnsetSeenFAI.
        /// </summary>
        PidTagCnsetSeenFAI,

        /// <summary>
        /// The PidTagCnsetRead.
        /// </summary>
        PidTagCnsetRead 
    }

    /// <summary>
    /// The ImportDeleteFlags.
    /// </summary>
    [Flags]
    public enum ImportDeleteFlags
    {
        /// <summary>
        /// If this flag is set, message deletions are being imported in server 2007. 
        /// </summary>
        delete = 0x00,

        /// <summary>
        /// If this flag is set, folder deletions are being imported.
        /// If this flag is not set, message deletions are being imported.
        /// </summary>
        Hierarchy = 0x01,

        /// <summary>
        /// If this flag is set, hard deletions are being imported.
        /// If this flag is not set, hard deletions are not being imported.
        /// It's not supported by Exchange 2007.
        /// </summary>
        HardDelete = 0x02,
    }

    /// <summary>
    /// The ROP operation's produce result.
    /// </summary>
    public enum RopResult : uint
    {
        /// <summary>
        /// The operation succeeded.
        /// </summary>
        Success = 0x00000000,

        /// <summary>
        /// An invalid parameter was passed to a remote procedure call (RPC).(ecInvalidParam)
        /// </summary>
        InvalidParameter = 0x80070057,

        /// <summary>
        /// The server does not implement this method call.
        /// </summary>
        NotImplemented = 0x80040FFF,

        /// <summary>
        /// In a change conflict, the client has the more recent change.
        /// </summary>
        NewerClientChange = 0x00040821,
        
        /// <summary>
        /// A badly formatted RPC buffer was detected.(ecRpcFormat)
        /// </summary>
        RpcFormat = 0x000004B6,

        /// <summary>
        /// A buffer passed to this function is not big enough.(ecBufferTooSmall)
        /// </summary>
        BufferTooSmall = 0x0000047D,

        /// <summary>
        /// The server does not support this method call. (ecNotSupported)
        /// </summary>
        NotSupported = 0x80040102,

        /// <summary>
        /// The parent folder could not be found.
        /// </summary>
        NoParentFolder = 0x80040803,

         /// <summary>
        /// The caller does not have sufficient access rights to perform the operation.
        /// </summary>
        AccessDenied = 0x80070005,
    }

    /// <summary>
    /// The status of the download operation after producing data for the TransferBuffer field.
    /// </summary>
    public enum TransferStatus
    {
        /// <summary>
        /// The download stopped because a non-recoverable error has occurred when producing a FastTransfer stream.
        /// </summary>
        Error = 0x0000,

        /// <summary>
        /// The FastTransfer stream was split, and more data is available.
        /// </summary>
        Partial = 0x0001,

        /// <summary>
        /// The FastTransfer stream was split, and more data is available. Only supported on 2007
        /// </summary>
        NoRoom = 0x0002,

        /// <summary>
        /// This is the last portion of the FastTransfer stream.
        /// </summary>
        Done = 0x0003,
    }

    /// <summary>
    /// Specific the amount of data in buffer.
    /// </summary>
    public enum BufferSize
    {
        /// <summary>
        /// The data size is 0xBABE.
        /// </summary>
        Normal = 0xBABE,

        /// <summary>
        /// The data size is less than 15480.
        /// </summary>
        TooSmall = 15479,

        /// <summary>
        /// The data size is greater than 32743.
        /// </summary>
        Greater = 32753,

        /// <summary>
        /// Runs into Partial status
        /// </summary>
        Partial = 1024,

        /// <summary>
        /// Runs into NoRoom status
        /// </summary>
        NoRoom = 10,

        /// <summary>
        /// Used for give an error value in the request in order to get an error return.
        /// </summary>
        Error = 1111
    }

    /// <summary>
    /// Property name enumeration.
    /// </summary>
    public enum PropertyTagName
    {
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
        /// Contains keywords or categories for the Message object. The length of each string within the multi-value string is less than 256.
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
        /// Indicates that this message has been automatically generated or automatically forwarded. If this property is unset, then a default value of "0x00" is assumed.
        /// </summary>
        PidTagAutoForwarded,

        /// <summary>
        /// Contains a comment added by the auto-forwarding agent.
        /// </summary>
        PidTagAutoForwardComment,

        /// <summary>
        /// Contains the unformatted text analogous to the text/plain body of [RFC2822].
        /// </summary>
        PidTagBody,

        /// <summary>
        /// Indicates the best available format for storing the message body.
        /// </summary>
        PidTagNativeBody,

        /// <summary>
        /// Contains the HTML body as specified in [RFC2822].
        /// </summary>
        PidTagBodyHtml,

        /// <summary>
        /// Contains a Rich Text Format (RTF) body compressed as specified .
        /// </summary>
        PidTagRtfCompressed,

        /// <summary>
        /// Indicates whether the RTF body has been synchronized with the contents in PidTagBody.
        /// </summary>
        PidTagRtfInSync,

        /// <summary>
        /// Indicates the code page used for PidTagBody or PidTagBodyHtml.
        /// </summary>
        PidTagInternetCodepage,

        /// <summary>
        /// Contains the list of address book EntryIDs linked to by this Message object.
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
        /// Contains the list of search keys for the Contact object linked to by this Message object.
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
        /// Specifies the number of days that a Message object can be retained.
        /// </summary>
        PidTagRetentionPeriod,

        /// <summary>
        /// A composite property that holds two pieces of information.
        /// </summary>
        PidTagStartDateEtc,

        /// <summary>
        /// Specifies the date, in UTC, after which a Message object is expired by the server. 
        /// </summary>
        PidTagRetentionDate,

        /// <summary>
        /// Contains flags that specify the status or nature of an item's retention tag or archive tag.
        /// </summary>
        PidTagRetentionFlags,

        /// <summary>
        /// Specifies the number of days that a Message object can remain un-archived.
        /// </summary>
        PidTagArchivePeriod,

        /// <summary>
        /// Specifies the date, in UTC, after which a Message object is moved to archive by the server.
        /// </summary>
        PidTagArchiveDate,

        /// <summary>
        /// Indicates the last time the file referenced by the Attachment object was modified, or the last time the Attachment object itself was modified.
        /// </summary>
        PidTagLastModificationTime,

        /// <summary>
        /// Indicates the time the file referenced by the Attachment object was created.
        /// </summary>
        PidTagCreationTime,

        /// <summary>
        /// Contains the name of the attachment as input by the end user.
        /// </summary>
        PidTagDisplayName,

        /// <summary>
        /// Contains the size in bytes consumed by the Attachment object on the server.
        /// </summary>
        PidTagAttachSize,

        /// <summary>
        /// Identifies the Attachment object within its Message object.
        /// </summary>
        PidTagAttachNumber,

        /// <summary>
        /// Represents the way the contents of an attachment are accessed.
        /// </summary>
        PidTagAttachMethod,

        /// <summary>
        /// Contains the full filename and extension of the Attachment object.
        /// </summary>
        PidTagAttachLongFilename,

        /// <summary>
        /// Contains the 8.3 name of PidTagAttachLongFilename.
        /// </summary>
        PidTagAttachFilename,

        /// <summary>
        /// Contains a filename extension that indicates the document type of an attachment.
        /// </summary>
        PidTagAttachExtension,

        /// <summary>
        /// Contains the fully qualified path and filename with extension.
        /// </summary>
        PidTagAttachLongPathname,

        /// <summary>
        /// Contains the 8.3 name of PidTagAttachLongPathname.
        /// </summary>
        PidTagAttachPathname,

        /// <summary>
        /// Contains the identifier information for the application which supplied the Attachment object's data.
        /// </summary>
        PidTagAttachTag,

        /// <summary>
        /// Represents an offset, in rendered characters, to use when rendering an attachment within the main message text.
        /// </summary>
        PidTagRenderingPosition,

        /// <summary>
        /// Contains a Windows metafile as specified in [MS-WMF] for the Attachment object.
        /// </summary>
        PidTagAttachRendering,

        /// <summary>
        /// Indicates which body formats might reference this attachment when rendering data.
        /// </summary>
        PidTagAttachFlags,

        /// <summary>
        /// Contains the name of an attachment file, modified so that it can be correlated with TNEF messages, see [MS-OXTNEF].
        /// </summary>
        PidTagAttachTransportName,

        /// <summary>
        /// Contains encoding information about the Attachment object.
        /// </summary>
        PidTagAttachEncoding,

        /// <summary>
        /// MUST be unset if PidTagAttachEncoding is unset.
        /// </summary>
        PidTagAttachAdditionalInformation,

        /// <summary>
        /// The type of Message object to which this attachment is linked.
        /// </summary>
        PidTagAttachmentLinkId,

        /// <summary>
        /// Indicates special handling for this Attachment object.
        /// </summary>
        PidTagAttachmentFlags,

        /// <summary>
        /// Indicates whether this Attachment object is hidden from the end user.
        /// </summary>
        PidTagAttachmentHidden,

        /// <summary>
        /// The content-type MIME header.
        /// </summary>
        PidTagAttachMimeTag,

        /// <summary>
        /// A content identifier unique to this Message object that matches a corresponding "cid:" Uniform Resource Identifier (URI) scheme reference in the HTML body of the Message object.
        /// </summary>
        PidTagAttachContentId,

        /// <summary>
        /// A relative or full URI that matches a corresponding reference in the HTML body of the Message object.
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
        /// Contains the contents of the file to be attached.
        /// </summary>
        PidTagAttachDataBinary,

        /// <summary>
        /// Specifies the status of the Message object.
        /// </summary>
        PidTagMessageFlags,

        /// <summary>
        /// Contains a list of blind carbon copy (Bcc) Recipient display names.
        /// </summary>
        PidTagDisplayBcc,

        /// <summary>
        /// Contains list of carbon copy (Cc) Recipient display names.
        /// </summary>
        PidTagDisplayCc,

        /// <summary>
        /// Contains a list of the primary Recipient display names, separated by semicolons, if an e-mail message has primary Recipient.
        /// </summary>
        PidTagDisplayTo,

        /// <summary>
        /// The description for security.
        /// </summary>
        PidTagSecurityDescriptor,

        /// <summary>
        /// Setting the url name.
        /// </summary>
        PidTagUrlCompNameSet,

        /// <summary>
        /// Setting the sender is trusted.
        /// </summary>
        PidTagTrustSender,

        /// <summary>
        /// The url name.
        /// </summary>
        PidTagUrlCompName,

        /// <summary>
        /// Contains a unique binary-comparable key that identifies an object for a search.
        /// </summary>
        PidTagSearchKey,

        /// <summary>
        /// Indicates the operations available to the client for the object.
        /// </summary>
        PidTagAccess,

        /// <summary>
        /// Contains the name of a Message object.
        /// </summary>
        PidTagCreatorName,

        /// <summary>
        /// The id for creator.
        /// </summary>
        PidTagCreatorEntryId,

        /// <summary>
        /// Contains the name of the last mail user to modify the object.
        /// </summary>
        PidTagLastModifierName,

        /// <summary>
        /// The id for last modifier.
        /// </summary>
        PidTagLastModifierEntryId,

        /// <summary>
        /// Have name or not for properties.
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
        /// Define a property for error result.
        /// </summary>
        unsigned,

        /// <summary>
        /// Contains the time a RopCreateMessage ([MS-OXCROPS] section 2.2.6.2) was processed.
        /// </summary>
        PidTagLocalCommitTime,
    }

    /// <summary>
    /// Contains the permissions for the specified user.
    /// </summary>
    [Flags]
    public enum PermissionLevels : uint
    {
        /// <summary>
        /// Has no permissions.
        /// </summary>
        None = 0x00000000,

        /// <summary>
        /// This flag indicates that the server MUST allow the specified user's client to retrieve detailed information about the appointments on the calendar through the Availability Web Service Protocol.
        /// </summary>
        FreeBusyDetailed = 0x00001000,

        /// <summary>
        /// Indicates that the server MUST allow the specified user's client to retrieve information through the Availability Web Service protocol.
        /// </summary>
        FreeBusySimple = 0x00000800,

        /// <summary>
        /// Indicates that the server MUST allow the specified user's client to see the folder in the folder hierarchy table and request a handle 
        /// for the folder by using a RopOpenFolder request.
        /// </summary>
        FolderVisible = 0x00000400,

        /// <summary>
        /// Indicates that the server MUST include the specified user in any list of administrative contacts associated with the folder.
        /// </summary>
        FolderContact = 0x00000200,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to modify properties set on the folder itself, including the folder permissions.
        /// </summary>
        FolderOwner = 0x00000100,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to create new folders within the folder.
        /// </summary>
        CreateSubFolder = 0x00000080,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to delete any Message object in the folder.
        /// </summary>
        DeleteAny = 0x00000040,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to modify any Message object in the folder.
        /// </summary>
        EditAny = 0x00000020,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to delete any Message object in the folder that was created by that user.
        /// </summary>
        DeleteOwned = 0x00000010,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to modify any Message object in the folder that was created by that user.
        /// </summary>
        EditOwned = 0x00000008,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to create new Message objects in the folder.
        /// </summary>
        Create = 0x00000002,

        /// <summary>
        /// If this flag is set, the server MUST allow the specified user's client to read any Message object in the folder.
        /// </summary>
        ReadAny = 0x00000001,

        /// <summary>
        /// Combine Owner role permissions.
        /// </summary>
        Owner = ReadAny | Create | EditOwned | DeleteOwned | EditAny | DeleteAny | CreateSubFolder | FolderOwner | FolderContact | FolderVisible
    }

    /// <summary>
    /// Flag indicate how the stream is created.
    /// </summary>
    public enum RopOfInitiatOperation
    {
        /// <summary>
        /// Created by RopSynchronizationConfigure operation and set SynchronizationType to Contents.
        /// </summary>
        RopSynchronizationContentConfigure,

        /// <summary>
        /// Created by RopSynchronizationConfigure operation and set SynchronizationType to Hierarchy.
        /// </summary>
        RopSynchronizationHierarchyConfigure,

        /// <summary>
        /// Created by RopSynchronizationGetTransferState operation.
        /// </summary>
        RopSynchronizationGetTransferState,

        /// <summary>
        /// Created by RopFastTranserSourceCopyTo or RopFastTranserSourceCopyProperties operation with folder as the input object.
        /// </summary>
        RopFastTranserSourceCopyToOrPropertiesFolder,

        /// <summary>
        /// Created by RopFastTranserSourceCopyTo or RopFastTranserSourceCopyProperties operation with message as the input object.
        /// </summary>
        RopFastTranserSourceCopyToOrPropertiesMessage,

        /// <summary>
        /// Created by RopFastTranserSourceCopyTo or RopFastTranserSourceCopyProperties operation with attachment as the input object.
        /// </summary>
        RopFastTranserSourceCopyToOrPropertiesAttachment,

        /// <summary>
        /// Created by RopFastTranserSourceCopyMessages operation.
        /// </summary>
        RopFastTranserSourceCopyMessages,

        /// <summary>
        /// Created by RopFastTranserSourceCopyFolder operation.
        /// </summary>
        RopFastTranserSourceCopyFolder,
    }

     /// <summary>
    /// Flags indicate which FastTransfer copy operation is called.
    /// </summary>
    public enum EnumFastTransferOperation
    {
        /// <summary>
        /// FastTransferSourceCopyTo operation is called.
        /// </summary>
        FastTransferSourceCopyTo,

        /// <summary>
        /// FastTransferSourceCopyProperties operation is called.
        /// </summary>
        FastTransferSourceCopyProperties,

        /// <summary>
        /// FastTransferSourceCopyMessage operation is called.
        /// </summary>
        FastTransferSourceCopyMessage,

        /// <summary>
        /// FastTransferSourceCopyFolder operation is called.
        /// </summary>
        FastTransferSourceCopyFolder,

        /// <summary>
        /// SynchronizationImportMessageChange operation is called.
        /// </summary>
        SynchronizationImportMessageChange,

        /// <summary>
        /// SynchronizationReadStateChanges operation is called.
        /// </summary>
        SynchronizationReadStateChanges,

        /// <summary>
        /// SynchronizationGetTransferState operation is called.
        /// </summary>
        SynchronizationGetTransferState,

        /// <summary>
        /// SynchronizationImportMessageMove operation is called.
        /// </summary>
        SynchronizationImportMessageMove,

        /// <summary>
        /// SynchronizationImportDeletes operation is called.
        /// </summary>
        SynchronizationImportDeletes 
    }

    /// <summary>
    /// Send option settings.
    /// </summary>
    [Flags]
    public enum SendOptionAlls : byte
    {
        /// <summary>
        /// When used on RopSynchronizationConfigure, MUST match the value of the Unicode
        /// SynchronizationFlag.
        /// </summary>
        Unicode = 0x01,
        
        /// <summary>
        /// If this flag is set, the Unicode flag MUST also be set.
        /// </summary>
        UseCpid = 0x02,
        
        /// <summary>
        /// Used in FastTransfer operations only when the client requests a FastTransfer
        /// stream with the intent of uploading it immediately to another destination
        /// server.  The ROP that uses this flag MUST be followed by RopTellVersion.
        /// </summary>
        ForUpload = 0x03,
        
        /// <summary>
        /// Used when a client supports recovery mode and requests that a server MUST
        /// attempt to recover from failures to download changes for individual messages.
        /// MUST NOT be set when ForUpload flag is set.
        /// </summary>
        RecoverMode = 0x04,
        
        /// <summary>
        /// See the following table for all possible combinations of encoding flags (Unicode
        /// and ForceUnicode).  The following table lists all valid combinations of the
        /// Unicode | ForceUnicode flags. Neither string properties MUST be output in
        /// the code page set on connection. Unicode string properties MUST be output
        /// either in Unicode, or in the code page set on the current connection, with
        /// Unicode being preferred.  Unicode | ForceUnicode string properties MUST be
        /// output in Unicode.
        /// </summary>
        ForceUnicode = 0x08,
        
        /// <summary>
        /// MUST NOT be passed for anything but contents synchronization download.  This
        /// flag is set if a client supports partial message downloads. If a server supports
        /// this mode, it SHOULD output partial message changes if it reduces the size
        /// of the produced stream. If a server does not support this mode, it does not
        /// output partial message changes and this flag is ignored
        /// </summary>
        PartialItem = 0x10,
        
        /// <summary>
        /// If is set,The server MUST fail the ROP. 
        /// </summary>
        Reserved1 = 0x20,
        
        /// <summary>
        /// If is set,The server MUST fail the ROP. 
        /// </summary>
        Reserved2 = 0x40,
        
        /// <summary>
        /// If is set,The server MUST fail the ROP.
        /// </summary>
        Invalid = 0x80,
    }

    /// <summary>
    /// CopyFlags for RopFastTransferDestinationConfigure.
    /// </summary>
    [Flags]
    public enum FastTransferDestinationConfigureCopyFlags
    {
        /// <summary>
        ///  Default with no specifications.
        /// </summary>
        None = 0,
        
        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.  If
        ///     this flag is set, the client identifies the FastTransfer operation being
        ///     configured as a logical part of a larger object move operation.  If this
        ///     flag is not set, the client is not identifying the FastTransfer operation
        ///     being configured as a logical part of a larger object move operation.
        /// </summary>
        Move = 1,

        /// <summary>
        /// Server should return 0x80070057.
        /// </summary>
        Invalid = 64,
    }

    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyTo.
    /// </summary>
    [Flags]
    public enum CopyToCopyFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.  If
        /// this flag is set, the client identifies the FastTransfer operation being
        /// configured as a logical part of a larger object move operation.  If this
        /// flag is not set, the client is not identifying the FastTransfer operation
        /// being configured as a logical part of a larger object move operation.  If
        /// this flag is specified for a download operation, the server SHOULD NOT output
        /// any objects in a FastTransfer stream that the client does not have permissions
        /// to delete.
        /// </summary>
        Move = 1,
       
        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 2,
        
        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 4,
        
        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 8,
        
        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused4 = 512,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused5 = 1024,
        
        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a message.  If set, the server
        /// SHOULD output the message body, and the body of embedded messages, in their
        /// original format.  If not set, the server MUST output message body in the
        /// compressed Rich Text Format (RTF).
        /// </summary>
        BestBody = 8192,
        
        /// <summary>
        /// Servers SHOULD return 0x80070057 
        /// </summary>
        Invalid = 0x10000000,
    }

    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyProperties
    /// </summary>
    [Flags]
    public enum CopyPropertiesCopyFlags
    {
        /// <summary>
        ///  Default with no specifications.
        /// </summary>
        None = 0,

        /// <summary>
        /// MUST NOT be passed if InputServerObject is not a folder or a message.  If
        /// this flag is set, the client identifies the FastTransfer operation being
        /// configured as a logical part of a larger object move operation.  If this
        /// flag is not set, the client is not identifying the FastTransfer operation
        /// being configured as a logical part of a larger object move operation.  If
        /// this flag is specified for a download operation, the server SHOULD NOT output
        /// any objects in a FastTransfer stream that the client does not have permissions
        /// to delete.
        /// </summary>
        Move = 1,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 2,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 4,
        
        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused3 = 8,
       
        /// <summary>
        /// Servers SHOULD return 0x80070057 
        /// </summary>
        Invalid = 64,
    }

    /// <summary>
    /// CopyFlags for RopFastTransferSourceCopyFolder
    /// </summary>
    [Flags]
    public enum CopyFolderCopyFlags
    {
        /// <summary>
        /// Default with no specifications.
        /// </summary>
        None = 0,

        /// <summary>
        /// If this flag is set, the client identifies the FastTransfer operation being
        /// configured as a logical part of a larger object move operation.  If this
        /// flag is not set, the client is not identifying the FastTransfer operation
        /// being configured as a logical part of a larger object move operation.
        /// </summary>
        Move = 1,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused1 = 2,

        /// <summary>
        /// MUST be ignored by the server.
        /// </summary>
        Unused2 = 4,

        /// <summary>
        ///  MUST be ignored by the server.
        /// </summary>
        Unused3 = 8,

        /// <summary>
        ///  If this flag is set, the server MUST recursively include the subfolders of
        ///     the folder specified in the InputServerObject in the scope. If this flag
        ///     is not set, the server MUST NOT recursively include the subfolders of the
        ///     folder specified in the InputServerObject in the scope.
        /// </summary>
        CopySubfolders = 16,

        /// <summary>
        ///  Servers SHOULD return 0x80070057 
        /// </summary>
        Invalid = 32,
    }

    /// <summary>
    /// An 8-bit enumeration that defines the type of synchronization 
    /// </summary>
    [Flags]
    public enum SynchronizationTypes
    {
        /// <summary>
        /// Indicates a contents synchronization.
        /// </summary>
        Contents = 0x01,

        /// <summary>
        /// Indicates a hierarchy synchronization.
        /// </summary>
        Hierarchy = 0x02,
        
        /// <summary>
        /// The InvalidParameter.
        /// </summary>
        InvalidParameter = 0x04,
    }

    /// <summary>
    /// Specifies the type of the object.
    /// </summary>
    [Flags]
    public enum ObjectType
    {
        /// <summary>
        /// The Folder.
        /// </summary>
        Folder,

        /// <summary>
        /// The Message.
        /// </summary>
        Message,

        /// <summary>
        /// The Attachment.
        /// </summary>
        Attachment,
    }

    /// <summary>
    /// Specifies the handle type.
    /// </summary>
    public enum InputHandleType
    {
        /// <summary>
        /// Represents a message handle.
        /// </summary>
        MessageHandle,

        /// <summary>
        /// Represents a folder handle.
        /// </summary>
        FolderHandle,

        /// <summary>
        /// Represents an attachment handle.
        /// </summary>
        AttachmentHandle,
    }

    /// <summary>
    /// Specify the delete flags when call GetContentsTable and GetHierarchyTable. 
    /// </summary>
    [System.Flags]
    public enum DeleteFlags
    {
        /// <summary>
        /// Initialize the parameter values.
        /// </summary>
        Initial = 0,

        /// <summary>
        /// Check the soft delete result.
        /// </summary>
        SoftDeleteCheck = 1,

        /// <summary>
        /// Check the hard delete result.
        /// </summary>
        HardDeleteCheck = 2,
    }
}