namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;

    #region RPC binding
    /// <summary>
    /// An enumeration represents authentication levels passed to various run-time functions.
    /// </summary>
    public enum AuthenticationLevel : uint
    {
        /// <summary>
        /// Uses the default authentication level for the specified authentication service.
        /// </summary>
        RPC_C_AUTHN_LEVEL_DEFAULT = 0,

        /// <summary>
        /// Performs no authentication.
        /// </summary>
        RPC_C_AUTHN_LEVEL_NONE = 1,

        /// <summary>
        /// Authenticates only when the client establishes a relationship with a server.
        /// </summary>
        RPC_C_AUTHN_LEVEL_CONNECT = 2,

        /// <summary>
        /// Authenticates only at the beginning of each remote procedure call when the server receives the request.
        /// </summary>
        RPC_C_AUTHN_LEVEL_CALL = 3,

        /// <summary>
        /// Authenticates only that all data received is from the expected client. Does not validate the data itself.
        /// </summary>
        RPC_C_AUTHN_LEVEL_PKT = 4,

        /// <summary>
        /// Authenticates and verifies that none of the data transferred between client and server has been modified.
        /// </summary>
        RPC_C_AUTHN_LEVEL_PKT_INTEGRITY = 5,

        /// <summary>
        /// Includes all previous levels, and ensures clear text data can only be seen by the sender and the receiver. In the local case, this involves using a secure channel.
        /// </summary>
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6
    }

    /// <summary>
    /// An enumeration represents the authentication services passed to various run-time functions.
    /// </summary>
    public enum AuthenticationService : uint
    {
        /// <summary>
        /// No authentication.
        /// </summary>
        RPC_C_AUTHN_NONE = 0,

        /// <summary>
        /// Use the Microsoft Negotiate SSP.
        /// </summary>
        RPC_C_AUTHN_GSS_NEGOTIATE = 9,

        /// <summary>
        /// Use the Microsoft NT LAN Manager (NTLM) SSP.
        /// </summary>
        RPC_C_AUTHN_WINNT = 10,
        
        /// <summary>
        /// Use the Microsoft Kerberos SSP.
        /// </summary>
        RPC_C_AUTHN_GSS_KERBEROS = 16
    }
    #endregion

    /// <summary>
    /// The value of the pulFlags that tell the server how to build the rgbOut parameter.
    /// </summary>
    [Flags]
    public enum PulFlags : uint
    {
        /// <summary>
        /// The value of the pulFlags is 0x00000001 means the server MUST NOT compress ROP response payload (rgbOut) or
        /// auxiliary payload (rgbAuxOut).
        /// </summary>
        NoCompression = 1,

        /// <summary>
        /// The value of the pulFlags is 0x00000002 means the server MUST NOT obfuscate the ROP response payload (rgbOut) 
        /// or auxiliary payload (rgbAuxOut).
        /// </summary>
        NoXorMagic = 2,

        /// <summary>
        /// The value of the pulFlags is 0x00000004 means the client allows chaining of ROP response payloads.
        /// </summary>
        Chain = 4
    }

    /// <summary>
    /// The type of the callback address in the rgbCallbackAddress field
    /// </summary>
    public enum Add_Families
    {
        /// <summary>
        /// AF_INET: indicates an address type for IP support
        /// </summary>
        AF_INET = 2,

        /// <summary>
        /// AF_INET6: indicates an address type for IPv6 support
        /// </summary>
        AF_INET6 = 23,

        /// <summary>
        /// The size of the rgbCallbackAddress doesn't correspond to the size of the sockaddr
        /// </summary>
        SIZENOTCORRESPONDSOCKADDRSIZE,

        /// <summary>
        /// The type of the rgbCallbackAddress is not supported by server
        /// </summary>
        NOT_SUPPORTED
    }

    /// <summary>
    /// The type of the ROP commands, refer to MS-OXCROPS for the detailed definition of each ROP command
    /// </summary>
    public enum ROPCommandType
    {
        /// <summary>
        /// This ROP logs on to a mailbox or public folder as administrator.
        /// </summary>
        RopLogon,

        /// <summary>
        /// This ROP logs on to a mailbox or public folder as normal user.
        /// </summary>
        RopLogonNormalUser,

        /// <summary>
        /// This ROP creates a Message object in a mailbox.
        /// </summary>
        RopCreateMessage,

        /// <summary>
        /// This ROP opens a property for streaming access.
        /// </summary>
        RopOpenStream,

        /// <summary>
        /// This ROP writes bytes to a stream.
        /// </summary>
        RopWriteStream,

        /// <summary>
        /// This ROP commits stream operations.
        /// </summary>
        RopCommitStream,

        /// <summary>
        /// This ROP reads bytes from a stream.
        /// </summary>
        RopReadStream,

        /// <summary>
        /// This ROP opens an existing folder in a mailbox.
        /// </summary>
        RopOpenFolder,

        /// <summary>
        /// This ROP creates a new subfolder.
        /// </summary>
        RopCreateFolder,

        /// <summary>
        /// This ROP gets the subfolder hierarchy table for a folder.
        /// </summary>
        RopGetHierarchyTable,

        /// <summary>
        /// This ROP commits the changes made to a message.
        /// </summary>
        RopSaveChangesMessage,

        /// <summary>
        /// This ROP retrieves rows from a table.
        /// </summary>
        RopQueryRows,

        /// <summary>
        /// This ROP retrieves rows from a table.
        /// </summary>
        RopDeleteFolder,

        /// <summary>
        /// This ROP downloads from a folder the content and descendant sub-objects for messages identified by a given set of IDs.
        /// </summary>
        RopFastTransferSourceCopyMessages,

        /// <summary>
        /// This ROP retrieves a stream of data from a fast transfer source object.
        /// </summary>
        RopFastTransferSourceGetBuffer,

        /// <summary>
        /// This ROP sets the properties visible on a table.
        /// </summary>
        RopSetColumns,

        /// <summary>
        /// This ROP registers for notification events.
        /// </summary>
        RopRegisterNotification,

        /// <summary>
        /// A request without any ROPs 
        /// </summary>
        WithoutRops,

        /// <summary>
        /// A request with more than two ROPs
        /// </summary>
        MultipleRops,

        /// <summary>
        /// This ROP gets the content table of a container.
        /// </summary>
        RopGetContentsTable,

        /// <summary>
        /// This ROP synchronizes deleted messages or folders.
        /// </summary>
        RopSynchronizationImportDeletes,

        /// <summary>
        /// This ROP imports new messages or full changes to existing messages into the server replica.
        /// </summary>
        RopSynchronizationImportMessageChange,

        /// <summary>
        /// This ROP converts a short-term ID into a long-term ID. 
        /// </summary>
        RopLongTermIdFromId,

        /// <summary>
        /// This ROP creates a new incremental change synchronization collector. 
        /// </summary>
        RopSynchronizationOpenCollector,

        /// <summary>
        /// This ROP hard delete subfolders and messages in target folder.
        /// </summary>
        RopHardDeleteMessagesAndSubfolders,

        /// <summary>
        /// This ROP releases all resources associated with a Server object. 
        /// </summary>
        RopRelease
    }

    /// <summary>
    /// Version information of the payload data that follows the AUX_HEADER.
    /// </summary>
    public enum AuxVersions
    {
        /// <summary>
        /// Aux version 1.
        /// </summary>
        AUX_VERSION_1 = 0x01,

        /// <summary>
        /// Aux version 2.
        /// </summary>
        AUX_VERSION_2 = 0x02
    }

    /// <summary>
    /// Type of payload data that follows the AUX_HEADER.
    /// </summary>
    public enum AuxTypes
    {
        /// <summary>
        /// Structure AUX_TYPE_PERF_CLIENTINFO.
        /// </summary>
        AUX_TYPE_PERF_CLIENTINFO = 0x02,

        /// <summary>
        /// Structure AUX_TYPE_PERF_SESSIONINFO.
        /// </summary>
        AUX_TYPE_PERF_SESSIONINFO = 0x04,

        /// <summary>
        /// Structure AUX_TYPE_PERF_MDB_SUCCESS.
        /// </summary>
        AUX_TYPE_PERF_MDB_SUCCESS = 0x07,

        /// <summary>
        /// Structure AUX_TYPE_PERF_PROCESSINFO.
        /// </summary>
        AUX_TYPE_PERF_PROCESSINFO = 0x0B,

        /// <summary>
        /// Structure AUX_TYPE_PERF_BG_DEFMDB_SUCCESS.
        /// </summary>
        AUX_TYPE_PERF_BG_DEFMDB_SUCCESS = 0x0C,

        /// <summary>
        /// Structure AUX_TYPE_PERF_BG_DEFGC_SUCCESS.
        /// </summary>
        AUX_TYPE_PERF_BG_DEFGC_SUCCESS = 0x0D,

        /// <summary>
        /// Structure AUX_TYPE_PERF_BG_GC_SUCCESS.
        /// </summary>
        AUX_TYPE_PERF_BG_GC_SUCCESS = 0x0F,

        /// <summary>
        /// Structure AUX_TYPE_PERF_BG_FAILURE.
        /// </summary>
        AUX_TYPE_PERF_BG_FAILURE = 0x10,

        /// <summary>
        /// Structure AUX_TYPE_PERF_ACCOUNTINFO.
        /// </summary>
        AUX_TYPE_PERF_ACCOUNTINFO = 0x18,

        /// <summary>
        /// Structure AUX_CLIENT_CONNECTION_INFO.
        /// </summary>
        AUX_CLIENT_CONNECTION_INFO = 0x4A,

        /// <summary>
        /// Structure AUX_TYPE_SERVER_CAPABILITIES.
        /// </summary>
        AUX_TYPE_SERVER_CAPABILITIES = 0x46
    }

    /// <summary>
    /// Indicates the server version values that are returned to the client on the EcDoConnectEx method call.
    /// </summary>
    public enum ServerVersionValues : long
    {
        /// <summary>
        /// The server supports passing the sentinel value 0xBABE in the BufferSize field of a RopFastTransferSourceGetBuffer request. (6.0.6755.0)
        /// </summary>
        SupportBufferSizeField = 0x600001A63,

        /// <summary>
        /// The server supports passing the sentinel value 0xBABE in the ByteCount field of a RopReadStream request. (8.0.295.0)
        /// </summary>
        SupportByteCountField = 0x800000127,

        /// <summary>
        /// The server supports the flag CLI_WITH_PER_MDB_FIX in the OpenFlags field of a RopLogon request. (8.0.324.0)
        /// </summary>
        SupportOpenFlagsField = 0x800000144,

        /// <summary>
        /// The server supports the EcDoAsyncConnectEx and EcDoAsyncWaitEx RPC function calls. (8.0.358.0)
        /// </summary>
        SupportAsync = 0x800000166,

        /// <summary>
        /// The server supports passing the flag ConversationMembers(0x80) in the TableFlags Field of a RopGetContentsTable request. (14.0.324.0)
        /// </summary>
        SupportTableFlagsField = 0xE00000144,

        /// <summary>
        /// The server supports passing the flag HardDelete(0x02) in the ImportDeleteFlags field of a RopSynchronizationImportDeletes request. (14.0.616.0)
        /// </summary>
        SupportImportDeleteFlagsField = 0xE00000268,

        /// <summary>
        /// The server supports passing the flag FailOnConflict(0x40) in the ImportFlag field of a RopSynchronizationImportMessageChange request. (14.1.67.0)
        /// </summary>
        SupportImportFlagField = 0xE00010043
    }

    /// <summary>
    /// An enumeration identifies the folder ID (FID).
    /// </summary>
    public enum FolderIds
    {
        /// <summary>
        /// Interpersonal Messages Sub-tree.
        /// </summary>
        InterpersonalMessage = 0x03,

        /// <summary>
        /// Inbox folder.
        /// </summary>
        Inbox = 0x04,

        /// <summary>
        /// Sent Items folder.
        /// </summary>
        SentItems = 0x06
    }

    /// <summary>
    /// The ID that identifies the property.
    /// </summary>
    public enum PropertyID
    {
        /// <summary>
        /// Contains a template data.
        /// </summary>
        PidTagTemplateData = 0x0001,

        /// <summary>
        /// Contains a subject.
        /// </summary>
        PidTagSubject = 0x0037,

        /// <summary>
        /// Contains the posting date of the item or entry.
        /// </summary>
        PidTagMessageDeliveryTime = 0x0E06,

        /// <summary>
        /// User defined ID.
        /// </summary>
        UserDefinedId = 0x1234,

        /// <summary>
        /// Contains the time of the last modification to the object in UTC.
        /// </summary>
        PidTagLastModificationTime = 0x3008,

        /// <summary>
        /// Contains the FID of the folder.
        /// </summary>
        PidTagFolderId = 0x6748,

        /// <summary>
        /// Contains a value that contains the MID of the message currently being synchronized.
        /// </summary>
        PidTagMid = 0x674A,

        /// <summary>
        /// Contains an identifier for all instances of a row in the table.
        /// </summary>
        PidTagInstID = 0x674D,

        /// <summary>
        /// Contains an identifier for a single instance of a row in the table.
        /// </summary>
        PidTagInstanceNum = 0x674E,
    }

    /// <summary>
    /// The enum of structures follows the AUX_HEADER in RgbAuxIn
    /// </summary>
    public enum RgbAuxInEnum
    {
        /// <summary>
        /// 2.2.2.4   AUX_PERF_SESSIONINFO
        /// </summary>
        AUX_PERF_SESSIONINFO,

        /// <summary>
        /// 2.2.2.5   AUX_PERF_SESSIONINFO_V2
        /// </summary>
        AUX_PERF_SESSIONINFO_V2,

        /// <summary>
        /// 2.2.2.6   AUX_PERF_CLIENTINFO
        /// </summary>
        AUX_PERF_CLIENTINFO,

        /// <summary>
        /// 2.2.2.8   AUX_PERF_PROCESSINFO
        /// </summary>
        AUX_PERF_PROCESSINFO,

        /// <summary>
        /// 2.2.2.9   AUX_PERF_DEFMDB_SUCCESS
        /// </summary>
        AUX_PERF_DEFMDB_SUCCESS,

        /// <summary>
        /// 2.2.2.10   AUX_PERF_DEFGC_SUCCESS
        /// </summary>
        AUX_PERF_DEFGC_SUCCESS,

        /// <summary>
        /// 2.2.2.11   AUX_PERF_MDB_SUCCESS
        /// </summary>
        AUX_PERF_MDB_SUCCESS_V2,

        /// <summary>
        /// 2.2.2.13   AUX_PERF_GC_SUCCESS
        /// </summary>
        AUX_PERF_GC_SUCCESS,

        /// <summary>
        /// 2.2.2.14   AUX_PERF_GC_SUCCESS_V2
        /// </summary>
        AUX_PERF_GC_SUCCESS_V2,

        /// <summary>
        /// 2.2.2.15   AUX_PERF_FAILURE
        /// </summary>
        AUX_PERF_FAILURE,

        /// <summary>
        /// 2.2.2.16   AUX_PERF_FAILURE_V2
        /// </summary>
        AUX_PERF_FAILURE_V2,

        /// <summary>
        /// 2.2.2.20   AUX_PERF_ACCOUNTINFO
        /// </summary>
        AUX_PERF_ACCOUNTINFO,

        /// <summary>
        /// The AUX_CLIENT_CONNECTION_INFO
        /// </summary>
        AUX_CLIENT_CONNECTION_INFO
    }

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

        /// <summary>
        /// The call was canceled.
        /// </summary>
        RPC_S_CALL_CANCELLED = 1818,
    }
}