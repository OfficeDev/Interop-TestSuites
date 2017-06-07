namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-OXCFXICSAdapter class.
    /// </summary>
    public interface IMS_OXCFXICSAdapter : IAdapter
    {
        /// <summary>
        /// Connect to the server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="connectionType">The type of connection</param>
        void Connect(int serverId, ConnectionType connectionType);

        /// <summary>
        /// Disconnect the connection to server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        void Disconnect(int serverId);

        /// <summary>
        /// Validate if the given two buffers are equal
        /// </summary>
        /// <param name="operation">Fast transfer operation</param>
        /// <param name="firstBufferIndex">The first buffer's index</param>
        /// <param name="secondBufferIndex">The second buffer's index</param>
        /// <returns>Returns true only if the two buffers are equal</returns>
        bool AreEqual(EnumFastTransferOperation operation, int firstBufferIndex, int secondBufferIndex);

        /// <summary>
        /// Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="rsid">Requirement id.</param>
        /// <param name="enabled">Requirement is enable or not.</param>
        void CheckRequirementEnabled(int rsid, out bool enabled);

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        void CheckMAPIHTTPTransportSupported(out bool isSupported);

        /// <summary>
        /// Logon to the Server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="flag">The type of logon.</param>
        /// <param name="logonHandleIndex">The server object handle index.</param>
        /// <param name="inboxFolderIdIndex">The inbox folder Id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult Logon(int serverId, LogonFlags flag, out int logonHandleIndex, out int inboxFolderIdIndex);

        /// <summary>
        /// Open a specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="folderHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult OpenFolder(int serverId, int objHandleIndex, int folderIdIndex, out int folderHandleIndex);

        /// <summary>
        /// Create a folder and return the folder handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderName">The new folder's name.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="folderHandleIndex">The new folder's handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult CreateFolder(int serverId, int objHandleIndex, string folderName, out int folderIdIndex, out int folderHandleIndex);

        /// <summary>
        /// Create a message and return the message handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index for creating message.</param>
        /// <param name="folderIdIndex">The folder Id index.</param>
        /// <param name="associatedFlag">The message is FAI or not.</param>
        /// <param name="messageHandleIndex">The created message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult CreateMessage(int serverId, int folderHandleIndex, int folderIdIndex, bool associatedFlag, out int messageHandleIndex);

        /// <summary>
        /// Create an attachment on specific message object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="messageHandleIndex">The message handle</param>
        /// <param name="attachmentHandleIndex">The attachment handle of created</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult CreateAttachment(int serverId, int messageHandleIndex, out int attachmentHandleIndex);

        /// <summary>
        /// Save the changes property of message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="messageHandleIndex">The message handle index.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SaveChangesMessage(int serverId, int messageHandleIndex, out int messageIdIndex);

        /// <summary>
        /// Commits the changes made to the Attachment object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="attachmentHandleIndex">The attachment handle</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SaveChangesAttachment(int serverId, int attachmentHandleIndex);

        /// <summary>
        /// Delete the specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult DeleteFolder(int serverId, int folderHandleIndex, int folderIdIndex);

        /// <summary>
        /// Open a specific message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The handle index folder object which the message in. </param>
        /// <param name="folderIdIndex">The folder id index of which the specific message in.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <param name="openMessageHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult OpenMessage(int serverId, int folderHandleIndex, int folderIdIndex, int messageIdIndex, out int openMessageHandleIndex);

        /// <summary>
        /// Release the object by handle.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The object handle index.</param>
        /// <returns>The ROP result</returns>
        RopResult Release(int serverId, int objHandleIndex);

        /// <summary>
        /// This method is used to check the second system under test whether is online or not.
        /// </summary>
        /// <param name="isSecondSUTOnline">Check second SUT is online or not</param>
        void CheckSecondSUTOnline(out bool isSecondSUTOnline);    

        /// <summary>
        /// Define the scope and parameters of the synchronization download operation. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The server object handle index.</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="synchronizationFlag">Flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="synchronizationExtraFlag">Extra flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="property">A list of properties and sub-objects to exclude or include.</param>
        /// <param name="downloadcontextHandleIndex">Synchronization download context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationConfigure(int serverId, int folderHandleIndex, SynchronizationTypes synchronizationType, SendOptionAlls option, SynchronizationFlag synchronizationFlag, SynchronizationExtraFlag synchronizationExtraFlag, Sequence<string> property, out int downloadcontextHandleIndex);

        /// <summary>
        /// Upload of an ICS state property into the synchronization context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">The synchronization context handle</param>
        /// <param name="icsPropertyType">Property tags of the ICS state properties.</param>
        /// <param name="isPidTagIdsetGivenInputAsInter32"> identifies Property tags as PtypInteger32.</param>
        /// <param name="icsStateIndex">The index of the ICS State.</param>
        /// <returns>The ICS state property is upload to the server successfully or not.</returns>
        RopResult SynchronizationUploadState(int serverId, int uploadContextHandleIndex, ICSStateProperties icsPropertyType, bool isPidTagIdsetGivenInputAsInter32, int icsStateIndex);

        /// <summary>
        /// Configures the synchronization upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index .</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="synchronizationHandleIndex">Synchronization upload context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationOpenCollector(int serverId, int objHandleIndex, SynchronizationTypes synchronizationType, out int synchronizationHandleIndex);

        /// <summary>
        /// Import new folders, or changes to existing folders, into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Upload context handle.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="properties">Properties to be set.</param>
        /// <param name="localFolderIdIndex">Local folder id index</param>
        /// <param name="folderIdIndex">The folder object id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportHierarchyChange(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex);

        /// <summary>
        /// Import new folders, or changes that have conflict change list to existing folders, into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Upload context handle.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="properties">Properties to be set.</param>
        /// <param name="localFolderIdIndex">Local folder id index</param>
        /// <param name="folderIdIndex">The folder object id index.</param>
        /// <param name="conflictType">The conflict type to generate, include version A include B, version B include A, and not include each other.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportHierarchyChangeWithConflict(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex, ConflictTypes conflictType);

        /// <summary>
        /// Creates a FastTransfer download context for a snapshot of the checkpoint ICS state of the operation identified by the given synchronization download context, or synchronization upload context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Synchronization context index.</param>
        /// <param name="stateHandleIndex">The index of FastTransfer download context for the ICS state.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationGetTransferState(int serverId, int objHandleIndex, out int stateHandleIndex);

        /// <summary>
        /// Import new messages or changes to existing messages into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">A synchronization upload context handle index.</param>
        /// <param name="localMessageidIndex">Message Id index.</param>
        /// <param name="importFlag">An 8-bit flag .</param>
        /// <param name="importMessageHandleIndex">The index of handle that indicate the Message object into which the client will upload the rest of the message changes.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportMessageChange(int serverId, int uploadContextHandleIndex, int localMessageidIndex, ImportFlag importFlag, out int importMessageHandleIndex);

        /// <summary>
        /// Imports message read state changes into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Sync handle.</param>
        /// <param name="objectHandleIndex">Message handle or folder handle or attachments handle.</param>
        /// <param name="readStatus">A boolean value indicating the message read status, true means read.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportReadStateChanges(int serverId, int objHandleIndex, int objectHandleIndex, bool readStatus);

        /// <summary>
        /// Imports deletions of messages or folders into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="uploadcontextHandleIndex">synchronization upload context handle</param>
        /// <param name="objIdIndexes">more object id</param>
        /// <param name="importDeleteFlag">Deletions type</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportDeletes(int serverId, int uploadcontextHandleIndex, Sequence<int> objIdIndexes, byte importDeleteFlag);

        /// <summary>
        /// Imports information about moving a message between two existing folders within the same mailbox.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="synchronizationUploadContextHandleIndex">The index of the synchronization upload context configured for collecting changes to the contents of the message move destination folder.</param>
        /// <param name="sourceFolderIdIndex">The index of the source folder id in object id container.</param>
        /// <param name="destinationFolderIdIndex">The index of the destination folder id in object id container.</param>
        /// <param name="sourceMessageIdIndex">The index of source message id in object id container.</param>
        /// <param name="sourceFolderHandleIndex">The index of source folder handle in handleContainer.</param>
        /// <param name="destinationFolderHandleIndex">The index of destination folder handle in handle container.</param>
        /// <param name="inewerClientChange">If the client has a newer message.</param>
        /// <param name="iolderversion">If the server have an older version of a message .</param>
        /// <param name="icnpc">Verify if the change number has been used.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SynchronizationImportMessageMove(int serverId, int synchronizationUploadContextHandleIndex, int sourceFolderIdIndex, int destinationFolderIdIndex, int sourceMessageIdIndex, int sourceFolderHandleIndex, int destinationFolderHandleIndex, bool inewerClientChange, out bool iolderversion, out bool icnpc);

        /// <summary>
        /// Identifies that a set of IDs either belongs to deleted messages in the specified folder or will never be used for any messages in the specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderhandleIndex">A Folder object handle index.</param>
        /// <param name="longTermIdRangeIndex">An array of LongTermIdRange structures defines a range of IDs, which are reported as unused or deleted.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SetLocalReplicaMidsetDeleted(int serverId, int folderhandleIndex, Sequence<int> longTermIdRangeIndex);

        /// <summary>
        /// Get specific property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleindex">Identify from which the property will be gotten.</param>
        /// <param name="propertyTag">A list of propertyTags.</param>
        /// <returns>Indicate the result of this ROP operation. </returns>
        RopResult GetPropertiesSpecific(int serverId, int handleindex, Sequence<string> propertyTag);

        /// <summary>
        /// Set the specific object's property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleindex">Server object handle index.</param>
        /// <param name="taggedPropertyValueArray">The list of property values.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult SetProperties(int serverId, int handleindex, Sequence<string> taggedPropertyValueArray);

        /// <summary>
        /// Allocates a range of internal identifiers for the purpose of assigning them to client-originated objects in a local replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="idcount">An unsigned 32-bit integer specifies the number of IDs to allocate.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult GetLocalReplicaIds(int serverId, int objHandleIndex, uint idcount);

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant sub-objects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">Type of object handle </param>
        /// <param name="level">Variable indicate whether copy the descendant sub-objects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="propertyTags">Array of properties and sub-objects to exclude.</param>
        /// <param name="copyToHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferSourceCopyTo(int serverId, int sourceHandleIndex, InputHandleType handleType, bool level, CopyToCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int copyToHandleIndex);

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant sub-objects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">Type of object handle. </param>
        /// <param name="level">Variable indicate whether copy the descendant sub-objects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="propertyTags">The list of properties and sub-objects to exclude.</param>
        /// <param name="copyPropertiesHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferSourceCopyProperties(int serverId, int sourceHandleIndex, InputHandleType handleType, bool level, CopyPropertiesCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int copyPropertiesHandleIndex);

        /// <summary>
        /// Initializes a FastTransfer operation on a folder for downloading content and descendant sub-objects for messages identified by a given set of IDs.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="messageIds">The list of MIDs the messages should copy.</param>
        /// <param name="copyMessageHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferSourceCopyMessages(int serverId, int sourceHandleIndex, RopFastTransferSourceCopyMessagesCopyFlags copyFlag, SendOptionAlls option, Sequence<int> messageIds, out int copyMessageHandleIndex);

        /// <summary>
        /// Initializes a FastTransfer operation to download properties and descendant sub-objects for a specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="copyFolderHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferSourceCopyFolder(int serverId, int sourceHandleIndex, CopyFolderCopyFlags copyFlag, SendOptionAlls option, out int copyFolderHandleIndex);

        /// <summary>
        /// Downloads the next portion of a FastTransfer stream.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="downloadHandleIndex">A fast transfer stream object handle index. </param>
        /// <param name="bufferSize">Specifies the maximum amount of data to be output in the TransferBuffer.</param>
        /// <param name="transferBufferIndex">The index of data get from the fast transfer stream.</param>
        /// <param name="abstractFastTransferStream">Fast transfer stream.</param>
        /// <param name="transferDataSmallOrEqualToBufferSize">Variable to not if the transferData is small or equal to bufferSize</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferSourceGetBuffer(int serverId, int downloadHandleIndex, BufferSize bufferSize, out int transferBufferIndex, out AbstractFastTransferStream abstractFastTransferStream, out bool transferDataSmallOrEqualToBufferSize);

        /// <summary>
        /// Tell the server of another server's version.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Server object handle index in handle container.</param>
        /// <param name="otherServerId">Another server's id.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult TellVersion(int serverId, int sourceHandleIndex, int otherServerId);

        /// <summary>
        ///  Uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="transferDataIndex">Transfer data index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferDestinationPutBuffer(int serverId, int sourceHandleIndex, int transferDataIndex);

        /// <summary>
        ///  Uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="transferDataIndex">Transfer data index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferDestinationPutBufferExtended(int serverId, int sourceHandleIndex, int transferDataIndex);

        /// <summary>
        ///  Initializes a FastTransfer operation for uploading content encoded in a client-provided FastTransfer stream into a mailbox
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="configHandleIndex">Configure handle's index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult FastTransferDestinationConfigure(int serverId, int sourceHandleIndex, SourceOperation option, FastTransferDestinationConfigureCopyFlags copyFlag, out int configHandleIndex);

        /// <summary>
        /// Modifies the permissions associated with a folder.
        /// </summary>
        /// <param name="serverId">the server id</param>
        /// <param name="folderHandleIndex">index of folder handle in container</param>
        /// <param name="permissionLevel">The permission level</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult ModifyPermissions(int serverId, int folderHandleIndex, PermissionLevels permissionLevel);

        /// <summary>
        /// Retrieve the content table for a folder. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult GetContentsTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount);

         /// <summary>
        /// Retrieve the hierarchy table for a folder. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        RopResult GetHierarchyTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount);
    }
}