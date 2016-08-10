[assembly: Microsoft.Xrt.Runtime.NativeType("System.Diagnostics.Tracing.*")]

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Model program.
    /// </summary>
    public static class Model
    {
        #region Variables

        /// <summary>
        /// Record the SHOULD/MAY requirements container.
        /// </summary>
        private static MapContainer<int, bool> requirementContainer = new MapContainer<int, bool>();

        /// <summary>
        /// Record the connections data container. 
        /// </summary>
        private static MapContainer<int, ConnectionData> connections = new MapContainer<int, ConnectionData>();

        /// <summary>
        /// Record the prior operation.
        /// </summary>
        private static PriorOperation priorOperation;

        /// <summary>
        /// Record whether Message change is partial or not.
        /// </summary>
        private static bool messagechangePartail;

        /// <summary>
        /// Record the SourceOperation of RopFastTransferDestinationPutBuffer.
        /// </summary>
        private static SourceOperation sourOperation;

        /// <summary>
        /// Record the prior download operation.
        /// </summary>
        private static PriorDownloadOperation priorDownloadOperation;

        /// <summary>
        /// Record the prior upload operation.
        /// </summary>
        private static PriorOperation priorUploadOperation;

        /// <summary>
        /// Record the soft delete message count.
        /// </summary>
        private static int softDeleteMessageCount;

        /// <summary>
        /// Record the soft delete folder count.
        /// </summary>
        private static int softDeleteFolderCount;
        #endregion

        /// <summary>
        /// Gets or sets the priorOperation.
        /// </summary>
        public static PriorOperation PriorOperation
        {
            get
            {
                return priorOperation;
            }

            set
            {
                priorOperation = value;
            }
        }

        #region Assistant Rop Interfaces

        /// <summary>
        ///  Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="rsid"> Indicate the requirement ID.</param>
        /// <param name="enabled"> Indicate the check result whether the requirement is enabled.</param>
        [Rule(Action = "CheckRequirementEnabled(rsid, out enabled)")]
        public static void CheckRequirementEnabled(int rsid, out bool enabled)
        {
            enabled = Choice.Some<bool>();
            requirementContainer.Add(rsid, enabled);
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        [Rule(Action = "CheckMAPIHTTPTransportSupported(out isSupported)")]
        public static void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            isSupported = Choice.Some<bool>();
        }

        /// <summary>
        /// This method is used to check whether the second system under test is online or not.
        /// </summary>
        /// <param name="isSecondSUTOnline"> Indicate the second SUT is online or not.</param>
        [Rule(Action = "CheckSecondSUTOnline(out isSecondSUTOnline)")]
        public static void CheckSecondSUTOnline(out bool isSecondSUTOnline)
        {
            isSecondSUTOnline = Choice.Some<bool>();
        }

        /// <summary>
        /// Connect to the server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="connectionType">The type of connection.</param>
        [Rule(Action = "Connect(serverId, connectionType)")]
        public static void Connect(int serverId, ConnectionType connectionType)
        {
            // Initialize ConnectionData.
            ConnectionData newConnection = new ConnectionData
            {
                FolderContainer = new Sequence<AbstractFolder>(),
                AttachmentContainer = new Sequence<AbstractAttachment>(),
                MessageContainer = new Sequence<AbstractMessage>()
            };

            // Create a new ConnectionData.
            connections.Add(serverId, newConnection);
        }

        /// <summary>
        /// Disconnect the connection to server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        [Rule(Action = "Disconnect(serverId)")]
        public static void Disconnect(int serverId)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Disconnect from server.
            connections.Remove(serverId);
        }

        /// <summary>
        /// Logon the Server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="flag">The type of logon.</param>
        /// <param name="logonHandleIndex">The server object handle index.</param>
        /// <param name="inboxFolderIdIndex">The inbox folder Id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "Logon(serverId,flag, out logonHandleIndex,out inboxFolderIdIndex)/result")]
        public static RopResult Logon(int serverId, LogonFlags flag, out int logonHandleIndex, out int inboxFolderIdIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            logonHandleIndex = AdapterHelper.GetHandleIndex();
            inboxFolderIdIndex = AdapterHelper.GetObjectIdIndex();
            ConnectionData changeConnection = connections[serverId];
            changeConnection.LogonHandleIndex = logonHandleIndex;
            changeConnection.LogonFolderType = flag;

            // Initialize the Container of ConnectionData.
            changeConnection.FolderContainer = new Sequence<AbstractFolder>();
            changeConnection.MessageContainer = new Sequence<AbstractMessage>();
            changeConnection.AttachmentContainer = new Sequence<AbstractAttachment>();
            changeConnection.DownloadContextContainer = new Sequence<AbstractDownloadInfo>();
            changeConnection.UploadContextContainer = new Sequence<AbstractUploadInfo>();

            // Create Inbox folder and set value for abstractInboxfolder.
            AbstractFolder inboxfolder = new AbstractFolder
            {
                FolderIdIndex = inboxFolderIdIndex,
                FolderPermission = PermissionLevels.ReadAny
            };

            // Add inbox folder to FolderContainer.
            changeConnection.FolderContainer = changeConnection.FolderContainer.Add(inboxfolder);
            connections[serverId] = changeConnection;
            RopResult result = RopResult.Success;
            return result;
        }

        /// <summary>
        /// Open a specific message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The handle index folder object which the message in. </param>
        /// <param name="folderIdIndex">The folder id index of which the specific message in.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <param name="openMessageHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "OpenMessage(serverId, objHandleIndex, folderIdIndex, messageIdIndex,out openMessageHandleIndex)/result")]
        public static RopResult OpenMessage(int serverId, int objHandleIndex, int folderIdIndex, int messageIdIndex, out int openMessageHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);
            Condition.IsTrue(connections[serverId].MessageContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            openMessageHandleIndex = 0;

            // Get information of ConnectionData.
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the current message is existent or not.
            AbstractMessage currentMessage = new AbstractMessage();
            bool ismessagExist = false;

            // Record current message.
            int messageIndex = 0;
            foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
            {
                if (tempMessage.MessageIdIndex == messageIdIndex)
                {
                    ismessagExist = true;
                    currentMessage = tempMessage;
                    messageIndex = changeConnection.MessageContainer.IndexOf(tempMessage);
                }
            }

            if (ismessagExist)
            {
                // Set value to current folder.
                currentMessage.MessageHandleIndex = AdapterHelper.GetHandleIndex();
                openMessageHandleIndex = currentMessage.MessageHandleIndex;

                // Update current message.
                changeConnection.MessageContainer = changeConnection.MessageContainer.Update(messageIndex, currentMessage);
                connections[serverId] = changeConnection;
                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Open specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="inboxFolderHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "OpenFolder(serverId, objHandleIndex, folderIdIndex, out inboxFolderHandleIndex)/result")]
        public static RopResult OpenFolder(int serverId, int objHandleIndex, int folderIdIndex, out int inboxFolderHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            inboxFolderHandleIndex = 0;

            // Get information of ConnectionData.
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the CurrentFolder is existent or not.
            AbstractFolder currentfolder = new AbstractFolder();
            bool isFolderExist = false;

            // Record current folder.
            int folderIndex = 0;
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderIdIndex == folderIdIndex)
                {
                    isFolderExist = true;
                    currentfolder = tempfolder;
                    folderIndex = changeConnection.FolderContainer.IndexOf(tempfolder);
                }
            }

            if (isFolderExist)
            {
                // Set value to current folder.
                currentfolder.FolderHandleIndex = AdapterHelper.GetHandleIndex();
                inboxFolderHandleIndex = currentfolder.FolderHandleIndex;

                // Initialize data of part of current folder.
                currentfolder.SubFolderIds = new Set<int>();
                currentfolder.MessageIds = new Set<int>();
                currentfolder.FolderProperties = new Set<string>();

                // Update current folder.
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(folderIndex, currentfolder);
                connections[serverId] = changeConnection;
                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Create a folder and return the folder handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderName">The new folder's name.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="folderHandleIndex">The new folder's handle index.</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "CreateFolder(serverId, objHandleIndex, folderName, out folderIdIndex, out folderHandleIndex)/result")]
        public static RopResult CreateFolder(int serverId, int objHandleIndex, string folderName, out int folderIdIndex, out int folderHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            folderIdIndex = 0;
            folderHandleIndex = 0;

            // Identify whether the Current Folder is existent or not.
            ConnectionData changeConnection = connections[serverId];
            AbstractFolder parentfolder = new AbstractFolder();
            bool isParentFolderExist = false;
            int parentfolderIndex = 0;

            // Find Current Folder.
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == objHandleIndex)
                {
                    isParentFolderExist = true;
                    parentfolder = tempfolder;
                    parentfolderIndex = changeConnection.FolderContainer.IndexOf(tempfolder);
                }
            }

            if (isParentFolderExist)
            {
                // Create a new folder.
                AbstractFolder currentfolder = new AbstractFolder
                {
                    FolderHandleIndex = AdapterHelper.GetHandleIndex()
                };

                // Set value for new folder
                folderHandleIndex = currentfolder.FolderHandleIndex;
                currentfolder.FolderIdIndex = AdapterHelper.GetObjectIdIndex();
                folderIdIndex = currentfolder.FolderIdIndex;
                currentfolder.ParentFolderHandleIndex = parentfolder.FolderHandleIndex;
                currentfolder.ParentFolderIdIndex = parentfolder.FolderIdIndex;
                currentfolder.FolderPermission = PermissionLevels.FolderOwner;

                // Initialize for new folder.
                currentfolder.SubFolderIds = new Set<int>();
                currentfolder.MessageIds = new Set<int>();
                currentfolder.ICSStateContainer = new MapContainer<int, AbstractUpdatedState>();

                // Update SubFolderIds of parent folder.
                parentfolder.SubFolderIds = parentfolder.SubFolderIds.Add(currentfolder.FolderIdIndex);

                // Update parent folder
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(parentfolderIndex, parentfolder);

                // Add new folder to FolderContainer.
                changeConnection.FolderContainer = changeConnection.FolderContainer.Add(currentfolder);
                connections[serverId] = changeConnection;

                if (folderIdIndex > 0)
                {
                    // Because only if the folder is right can return a valid folderIdIndex, then the requirement is verified.
                    ModelHelper.CaptureRequirement(
                        1890,
                        @"[In Identifying Objects and Maintaining Change Numbers] On creation, objects in the mailbox are assigned internal identifiers, commonly known as Folder ID structures ([MS-OXCDATA] section 2.2.1.1) for folders.");
                }

                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Create a message and return the message handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index for creating message.</param>
        /// <param name="folderIdIndex">The folder Id index.</param>
        /// <param name="associatedFlag">The message is FAI or not.</param>
        /// <param name="messageHandleIndex">The created message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "CreateMessage(serverId, folderHandleIndex, folderIdIndex, associatedFlag, out messageHandleIndex)/result")]
        public static RopResult CreateMessage(int serverId, int folderHandleIndex, int folderIdIndex, bool associatedFlag, out int messageHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            messageHandleIndex = -1;

            // Identify whether the Current Folder is existent or not.
            ConnectionData changeConnection = connections[serverId];
            AbstractFolder parentfolder = new AbstractFolder();
            bool isParentFolderExist = false;
            int parentfolderIndex = 0;

            // Find current folder.
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == folderHandleIndex)
                {
                    isParentFolderExist = true;
                    parentfolder = tempfolder;
                    parentfolderIndex = changeConnection.FolderContainer.IndexOf(tempfolder);
                }
            }

            if (isParentFolderExist)
            {
                // Create a new message object.
                AbstractMessage currentMessage = new AbstractMessage
                {
                    IsFAImessage = associatedFlag,
                    IsRead = true,
                    FolderHandleIndex = folderHandleIndex,
                    FolderIdIndex = folderIdIndex,
                    MessageHandleIndex = AdapterHelper.GetHandleIndex(),
                    MessageProperties = new Sequence<string>()
                };

                // Set value for new message.

                // Initialize message properties.
                messageHandleIndex = currentMessage.MessageHandleIndex;

                // Update folder
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(parentfolderIndex, parentfolder);

                // Add new message to MessageContainer.
                changeConnection.MessageContainer = changeConnection.MessageContainer.Add(currentMessage);
                connections[serverId] = changeConnection;

                if (currentMessage.MessageHandleIndex > 0)
                {
                    // Because only if the folder is right can return a valid message handle index, then the requirement is verified.
                    ModelHelper.CaptureRequirement(
                        1890001,
                        @"[In Identifying Objects and Maintaining Change Numbers] On creation, objects in the mailbox are assigned internal identifiers, commonly known as Message ID structures ([MS-OXCDATA] section 2.2.1.2) for messages.");
                }

                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Create an attachment on specific message object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="messageHandleIndex">The message handle</param>
        /// <param name="attachmentHandleIndex">The attachment handle of created</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "CreateAttachment(serverId,messageHandleIndex, out attachmentHandleIndex)/result")]
        public static RopResult CreateAttachment(int serverId, int messageHandleIndex, out int attachmentHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            attachmentHandleIndex = -1;

            // Identify whether the Current message is existent or not.
            ConnectionData changeConnection = connections[serverId];
            AbstractMessage currentMessage = new AbstractMessage();
            bool iscurrentMessageExist = false;
            int currentMessageIndex = 0;

            // Find current message
            foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
            {
                if (tempMessage.MessageHandleIndex == messageHandleIndex)
                {
                    iscurrentMessageExist = true;
                    currentMessage = tempMessage;
                    currentMessageIndex = changeConnection.MessageContainer.IndexOf(tempMessage);
                }
            }

            if (iscurrentMessageExist)
            {
                // Create a new attachment.
                AbstractAttachment currentAttachment = new AbstractAttachment();

                // Set value for new attachment.
                currentMessage.AttachmentCount++;
                changeConnection.MessageContainer.Update(currentMessageIndex, currentMessage);
                currentAttachment.AttachmentHandleIndex = AdapterHelper.GetHandleIndex();
                attachmentHandleIndex = currentAttachment.AttachmentHandleIndex;

                // Add new attachment to attachment container.
                changeConnection.AttachmentContainer = changeConnection.AttachmentContainer.Add(currentAttachment);
                connections[serverId] = changeConnection;
                result = RopResult.Success;
            }

            // There is no negative behavior specified in this protocol, so this operation always return true.
            return result;
        }

        /// <summary>
        /// Save the changes property of message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="messageHandleIndex">The message handle index.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SaveChangesMessage(serverId, messageHandleIndex, out messageIdIndex)/result")]
        public static RopResult SaveChangesMessage(int serverId, int messageHandleIndex, out int messageIdIndex)
        {
            // The contraction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].MessageContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            messageIdIndex = -1;

            // Identify whether the Current message is existent or not.
            ConnectionData changeConnection = connections[serverId];
            AbstractMessage currentMessage = new AbstractMessage();
            bool isMessageExist = false;
            int messageIndex = 0;

            // Find current message
            foreach (AbstractMessage tempMesage in changeConnection.MessageContainer)
            {
                if (tempMesage.MessageHandleIndex == messageHandleIndex)
                {
                    isMessageExist = true;
                    currentMessage = tempMesage;
                    messageIndex = changeConnection.MessageContainer.IndexOf(tempMesage);
                }
            }

            if (isMessageExist)
            {
                // Find the parent folder of relate message
                AbstractFolder parentfolder = new AbstractFolder();
                int parentfolderIndex = 0;
                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if (tempFolder.FolderHandleIndex == currentMessage.FolderHandleIndex)
                    {
                        parentfolder = tempFolder;
                        parentfolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                    }
                }

                // If new message then return a new message id.
                if (currentMessage.MessageIdIndex == 0)
                {
                    currentMessage.MessageIdIndex = AdapterHelper.GetObjectIdIndex();

                    // Because if Create a new messageID then the action which convert GID to a short-term internal identifier and assign it to an imported object execute in MS_OXCFXICSAdapter. So cover this requirement here.
                    ModelHelper.CaptureRequirement(1910, "[In Identifying Objects and Maintaining Change Numbers] 	Convert the GID structure ([MS-OXCDATA] section 2.2.1.3) to a short-term internal identifier and assign it to an imported object, if the external identifier is a GID value.");
                }

                // Set value for the current Message
                messageIdIndex = currentMessage.MessageIdIndex;
                parentfolder.MessageIds = parentfolder.MessageIds.Add(messageIdIndex);
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(parentfolderIndex, parentfolder);

                // Assign a new Change number.
                currentMessage.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();

                // Because of executed import operation before execute RopSaveChangesMessage operation. And assign a new changeNumber. So can cover this requirement here.
                ModelHelper.CaptureRequirement(
                    1906,
                    @"[In Identifying Objects and Maintaining Change Numbers]Upon successful import of a new or changed object using ICS upload, the server MUST do the following when receiving the RopSaveChangesMessage ROP:Assign the object a new internal change number (PidTagChangeNumber property (section 2.2.1.2.3)).");

                // Because of it must execute RopSaveChangesMessage operation after the messaging object each time and assign a new changeNumber. So can cover this requirement Spec here.
                ModelHelper.CaptureRequirement(1898, "[In Identifying Objects and Maintaining Change Numbers]A new change number is assigned to a messaging object each time it is modified.");
                currentMessage.ReadStateChangeNumberIndex = 0;

                // Update current Message into MessageContainer
                changeConnection.MessageContainer = changeConnection.MessageContainer.Update(messageIndex, currentMessage);
                connections[serverId] = changeConnection;

                if (priorOperation == MS_OXCFXICS.PriorOperation.RopCreateMessage && messageIndex > 0)
                {
                    // When the prior operate is create message and in this ROP return a valid messageIDIndex means this requirement verified.
                    ModelHelper.CaptureRequirement(
                        1890001,
                        @"[In Identifying Objects and Maintaining Change Numbers] On creation, objects in the mailbox are assigned internal identifiers, commonly known as Message ID structures ([MS-OXCDATA] section 2.2.1.2) for messages.");
                }

                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Commits the changes made to the Attachment object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="attachmentHandleIndex">The attachment handle</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "SaveChangesAttachment(serverId,attachmentHandleIndex)/result")]
        public static RopResult SaveChangesAttachment(int serverId, int attachmentHandleIndex)
        {
            // Contraction condition is the Attachment is created successful
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].AttachmentContainer.Count > 0);

            // Return Success
            RopResult result = RopResult.Success;
            return result;
        }

        /// <summary>
        /// Release the object by handle.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The object handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "Release(serverId, objHandleIndex)/result")]
        public static RopResult Release(int serverId, int objHandleIndex)
        {
            // The contraction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // The operation success.
            return RopResult.Success;
        }

        /// <summary>
        /// Delete the specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "DeleteFolder(serverId, folderHandleIndex, folderIdIndex)/result")]
        public static RopResult DeleteFolder(int serverId, int folderHandleIndex, int folderIdIndex)
        {
            // The contraction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the Current folder and Parent folder are existent or not.
            AbstractFolder currentfolder = new AbstractFolder();
            AbstractFolder parentfolder = new AbstractFolder();
            bool isCurrentFolderExist = false;
            bool isParentFolderExist = false;

            // Find parent folder.
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == folderHandleIndex)
                {
                    isParentFolderExist = true;
                    parentfolder = tempfolder;
                }
            }

            // Find current folder.
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderIdIndex == folderIdIndex)
                {
                    isCurrentFolderExist = true;
                    currentfolder = tempfolder;
                }
            }

            if (isParentFolderExist && isCurrentFolderExist)
            {
                // Remove current folder from SubFolderIds property of parent folder
                parentfolder.SubFolderIds = parentfolder.SubFolderIds.Remove(currentfolder.FolderIdIndex);

                // Remove current folder
                changeConnection.FolderContainer = changeConnection.FolderContainer.Remove(currentfolder);
                connections[serverId] = changeConnection;
                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Get specific property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleIndex">Identify from which the property will be gotten.</param>
        /// <param name="propertyTag">A list of propertyTags.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "GetPropertiesSpecific(serverId, handleIndex, propertyTag)/result")]
        public static RopResult GetPropertiesSpecific(int serverId, int handleIndex, Sequence<string> propertyTag)
        {
            // The contraction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // Identify whether the Current message is existent or not.
            ConnectionData changeConnection = connections[serverId];

            if (connections[serverId].FolderContainer.Count > 0)
            {
                result = RopResult.Success;
            }
            else if (connections[serverId].MessageContainer.Count > 0)
            {
                AbstractMessage currentMessage = new AbstractMessage();
                bool ismessageExist = false;
                int messageIndex = 0;

                foreach (AbstractMessage tempMesage in changeConnection.MessageContainer)
                {
                    if (tempMesage.MessageHandleIndex == handleIndex)
                    {
                        ismessageExist = true;
                        currentMessage = tempMesage;
                        messageIndex = changeConnection.MessageContainer.IndexOf(tempMesage);
                    }
                }

                if (ismessageExist)
                {
                    // Set value for MessageProperties
                    currentMessage.MessageProperties = propertyTag;
                    changeConnection.MessageContainer = changeConnection.MessageContainer.Update(messageIndex, currentMessage);
                    connections[serverId] = changeConnection;
                    result = RopResult.Success;
                }
            }

            return result;
        }

        /// <summary>
        /// Set the specific object's property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleIndex">Server object handle index.</param>
        /// <param name="propertyTag">The list of property values.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SetProperties(serverId, handleIndex, propertyTag)/result")]
        public static RopResult SetProperties(int serverId, int handleIndex, Sequence<string> propertyTag)
        {
            // The construction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].MessageContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // Get value of current ConnectionData
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the Current message is existent or not.
            AbstractMessage currentMessage = new AbstractMessage();
            bool ismessageExist = false;
            int messageIndex = 0;

            // Find current message.
            foreach (AbstractMessage tempMesage in changeConnection.MessageContainer)
            {
                if (tempMesage.MessageHandleIndex == handleIndex)
                {
                    ismessageExist = true;
                    currentMessage = tempMesage;
                    messageIndex = changeConnection.MessageContainer.IndexOf(tempMesage);
                }
            }

            if (ismessageExist)
            {
                foreach (string propertyName in propertyTag)
                {
                    // Identify whether the property is existent or not in MessageProperties.
                    if (!currentMessage.MessageProperties.Contains(propertyName))
                    {
                        // Add property to MessageProperties.
                        currentMessage.MessageProperties = currentMessage.MessageProperties.Add(propertyName);
                    }
                }

                changeConnection.MessageContainer = changeConnection.MessageContainer.Update(messageIndex, currentMessage);
                connections[serverId] = changeConnection;

                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Modifies the permissions associated with a folder.
        /// </summary>
        /// <param name="serverId">The server id</param>
        /// <param name="folderHandleIndex">index of folder handle in container</param>
        /// <param name="permissionLevel">The permission level</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "ModifyPermissions(serverId, folderHandleIndex, permissionLevel)/result")]
        public static RopResult ModifyPermissions(int serverId, int folderHandleIndex, PermissionLevels permissionLevel)
        {
            // The contraction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the Current folder is existent or not.
            AbstractFolder currentfolder = new AbstractFolder();
            bool isCurrentFolderExist = false;
            int currentfolderIndex = 0;

            // Find current folder
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == folderHandleIndex)
                {
                    isCurrentFolderExist = true;
                    currentfolder = tempfolder;
                    currentfolderIndex = changeConnection.FolderContainer.IndexOf(tempfolder);
                }
            }

            if (isCurrentFolderExist)
            {
                // Set folder Permission for CurrentFolder.
                currentfolder.FolderPermission = permissionLevel;
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(currentfolderIndex, currentfolder);
                connections[serverId] = changeConnection;
                result = RopResult.Success;
            }

            return result;
        }

        #endregion

        #region MS-OXCFXICS operation actions
        /// <summary>
        /// Define the scope and parameters of the synchronization download operation. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder object handle.</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="synchronizationFlag">Flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="synchronizationExtraFlag">Extra Flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="property">A list of properties and sub objects to exclude or include.</param>
        /// <param name="downloadcontextHandleIndex">Synchronization download context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SynchronizationConfigure(serverId, folderHandleIndex, synchronizationType, option, synchronizationFlag, synchronizationExtraFlag,property, out downloadcontextHandleIndex)/result")]
        public static RopResult SynchronizationConfigure(int serverId, int folderHandleIndex, SynchronizationTypes synchronizationType, SendOptionAlls option, SynchronizationFlag synchronizationFlag, SynchronizationExtraFlag synchronizationExtraFlag, Sequence<string> property, out int downloadcontextHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize return value.
            RopResult result = RopResult.Success;
            downloadcontextHandleIndex = -1;

            if ((option & SendOptionAlls.Invalid) == SendOptionAlls.Invalid && (requirementContainer.ContainsKey(3463) && requirementContainer[3463]))
            {
                result = RopResult.InvalidParameter;
                return result;
            }

            // SynchronizationFlag MUST match the value of the Unicode flag from SendOptions field.
            if ((synchronizationFlag & SynchronizationFlag.Unicode) == SynchronizationFlag.Unicode)
            {
                Condition.IsTrue((option & SendOptionAlls.Unicode) == SendOptionAlls.Unicode);
            }

            // When SynchronizationType is 0X04 then Servers return 0x80070057.
            if (synchronizationType == SynchronizationTypes.InvalidParameter)
            {
                if (requirementContainer.ContainsKey(2695) && requirementContainer[2695])
                {
                    result = RopResult.NotSupported;
                }
                else
                {
                    result = RopResult.InvalidParameter;
                    ModelHelper.CaptureRequirement(2695, "[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] Servers MUST return 0x80070057 if SynchronizationType is 0x04.");
                }
            }
            else if ((synchronizationFlag & SynchronizationFlag.Reserved) == SynchronizationFlag.Reserved)
            {
                // When SynchronizationType is Reserved then Servers MUST fail the ROP request.
                result = RopResult.RpcFormat;
                ModelHelper.CaptureRequirement(2180, "[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] The server MUST fail the ROP request if the Reserved flag of the SynchronizationFlags field is set.");
            }
            else
            {
                // Get ConnectionData value.
                ConnectionData changeConnection = connections[serverId];

                // Identify whether the CurrentFolder is existent or not.
                bool isCurrentFolderExist = false;

                // Identify whether the Current Folder is existent or not.
                foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
                {
                    if (tempfolder.FolderHandleIndex == folderHandleIndex)
                    {
                        // Set the value to the variable when the current folder is existent.
                        isCurrentFolderExist = true;
                    }
                }

                // The condition of CurrentFolder is existent.
                if (isCurrentFolderExist)
                {
                    // Initialize the Download information.
                    AbstractDownloadInfo abstractDownloadInfo = new AbstractDownloadInfo
                    {
                        UpdatedState =
                            new AbstractUpdatedState
                            {
                                CnsetRead = new Set<int>(),
                                CnsetSeen = new Set<int>(),
                                CnsetSeenFAI = new Set<int>(),
                                IdsetGiven = new Set<int>()
                            },
                        DownloadHandleIndex = AdapterHelper.GetHandleIndex()
                    };

                    // Get the download Handle for download context.
                    downloadcontextHandleIndex = abstractDownloadInfo.DownloadHandleIndex;
                    ModelHelper.CaptureRequirement(669, "[In RopSynchronizationConfigure ROP Response Buffer]OutputServerObject: This value MUST be the synchronization download context.");

                    // Record the flags.
                    abstractDownloadInfo.Sendoptions = option;
                    abstractDownloadInfo.SynchronizationType = synchronizationType;
                    abstractDownloadInfo.Synchronizationflag = synchronizationFlag;
                    abstractDownloadInfo.SynchronizationExtraflag = synchronizationExtraFlag;

                    // Record the Property.
                    abstractDownloadInfo.Property = property;

                    // Record folder handle of related to the download context. 
                    abstractDownloadInfo.RelatedObjectHandleIndex = folderHandleIndex;
                    switch (synchronizationType)
                    {
                        // Record synchronizationType value for condition of Synchronization type is Contents.
                        case SynchronizationTypes.Contents:
                            abstractDownloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.contentsSync;
                            abstractDownloadInfo.ObjectType = ObjectType.Folder;
                            break;

                        // Record synchronizationType value for condition of Synchronization type is Hierarchy.
                        case SynchronizationTypes.Hierarchy:
                            abstractDownloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.hierarchySync;
                            abstractDownloadInfo.ObjectType = ObjectType.Folder;
                            break;
                        default:

                            // Condition ofsynchronizationType is invalid parameter.
                            result = RopResult.InvalidParameter;
                            break;
                    }

                    // Condition of the operation return success.
                    if (result == RopResult.Success)
                    {
                        // Add the  new value to DownloadContextContainer.
                        changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(abstractDownloadInfo);
                        connections[serverId] = changeConnection;
                        priorDownloadOperation = PriorDownloadOperation.RopSynchronizationConfigure;
                        priorOperation = MS_OXCFXICS.PriorOperation.RopSynchronizationConfigure;

                        ModelHelper.CaptureRequirement(
                            641,
                            @"[In RopSynchronizationConfigure] The RopSynchronizationConfigure ROP ([MS-OXCROPS] section 2.2.13.1) is used to define the synchronization scope and parameters of the synchronization download operation.");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Configures the synchronization upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder object handle index .</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="uploadContextHandleIndex">Synchronization upload context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SynchronizationOpenCollector(serverId, folderHandleIndex, synchronizationType, out uploadContextHandleIndex)/result")]
        public static RopResult SynchronizationOpenCollector(int serverId, int folderHandleIndex, SynchronizationTypes synchronizationType, out int uploadContextHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize return value
            RopResult result = RopResult.InvalidParameter;
            uploadContextHandleIndex = -1;

            // Get ConnectionData value.
            ConnectionData changeConnection = connections[serverId];
            AbstractFolder currentfolder = new AbstractFolder();

            // Identify whether the CurrentFolder is existent or not.
            bool isCurrentFolderExist = false;

            // Identify whether the Current Folder is existent or not.
            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == folderHandleIndex)
                {
                    // Set the value to the variable when the current folder is existent.
                    isCurrentFolderExist = true;
                    currentfolder = tempfolder;
                }
            }

            if (isCurrentFolderExist)
            {
                // Initialize the upload information.
                AbstractUploadInfo abstractUploadInfo = new AbstractUploadInfo
                {
                    UploadHandleIndex = AdapterHelper.GetHandleIndex()
                };
                uploadContextHandleIndex = abstractUploadInfo.UploadHandleIndex;
                ModelHelper.CaptureRequirement(778, "[In RopSynchronizationOpenCollector ROP Response Buffer]OutputServerObject: The value of this field MUST be the synchronization upload context.");
                abstractUploadInfo.SynchronizationType = synchronizationType;
                abstractUploadInfo.RelatedObjectHandleIndex = folderHandleIndex;
                abstractUploadInfo.RelatedObjectIdIndex = currentfolder.FolderIdIndex;

                // Initialize the updatedState information.
                abstractUploadInfo.UpdatedState.IdsetGiven = new Set<int>();
                abstractUploadInfo.UpdatedState.CnsetRead = new Set<int>();
                abstractUploadInfo.UpdatedState.CnsetSeen = new Set<int>();
                abstractUploadInfo.UpdatedState.CnsetSeenFAI = new Set<int>();

                // Add the  new value to UploadContextContainer.
                changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Add(abstractUploadInfo);

                connections[serverId] = changeConnection;

                // Record RopSynchronizationImportHierarchyChange operation.
                priorOperation = PriorOperation.RopSynchronizationOpenCollector;
                result = RopResult.Success;

                if (uploadContextHandleIndex != -1)
                {
                    // Because if uploadContextHandleIndex doesn't equal -1  and the ROP return success, so only if this ROP success and return a valid handler this requirement will be verified.
                    ModelHelper.CaptureRequirement(
                        769,
                        @"[In RopSynchronizationOpenCollector ROP] The RopSynchronizationOpenCollector ROP ([MS-OXCROPS] section 2.2.13.7) configures the synchronization upload operation and returns a handle to a synchronization upload context.");
                }
            }

            return result;
        }

        /// <summary>
        /// Imports deletions of messages or folders into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="uploadContextHandleIndex">synchronization upload context handle</param>
        /// <param name="objIdIndexes">all object id</param>
        /// <param name="importDeleteFlag">Deletions type</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "SynchronizationImportDeletes(serverId,uploadContextHandleIndex,objIdIndexes,importDeleteFlag)/result")]
        public static RopResult SynchronizationImportDeletes(int serverId, int uploadContextHandleIndex, Sequence<int> objIdIndexes, byte importDeleteFlag)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);
            if (requirementContainer.ContainsKey(90205002) && requirementContainer[90205002])
            {
                Condition.IsTrue(((ImportDeleteFlags)importDeleteFlag & ImportDeleteFlags.delete) == ImportDeleteFlags.delete);
            }

            // Initialize return value.
            RopResult result = RopResult.InvalidParameter;

            // When the ImportDeleteFlags flag is set HardDelete by Exchange 2007 then server return 0x80070057.
            if (importDeleteFlag == (byte)ImportDeleteFlags.HardDelete)
            {
                if (requirementContainer.ContainsKey(2593) && requirementContainer[2593])
                {
                    result = RopResult.NotSupported;
                    ModelHelper.CaptureRequirement(2593, "[In Appendix A: Product Behavior] <16> Section 2.2.3.2.4.5.1: The HardDelete flag is not supported by Exchange 2003 or Exchange 2007.");
                    return result;
                }
            }

            // When the ImportDeleteFlags flag is an invalid value (0x10) then server returns 0x80070057.
            if (importDeleteFlag == 0x10)
            {
                if (requirementContainer.ContainsKey(2254001) && requirementContainer[2254001])
                {
                    result = RopResult.NotSupported;
                }
                else
                {
                    result = RopResult.InvalidParameter;
                }

                return result;
            }

            // Get ConnectionData value.
            ConnectionData changeConnection = connections[serverId];

            // Create uploadInfo variable.
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();

            // Identify whether the Current Upload information is existent or not.
            bool isCurrentUploadinfoExist = false;

            // Record the current uploadInfo index.
            int currentUploadIndex = 0;
            foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
            {
                if (tempUploadInfo.UploadHandleIndex == uploadContextHandleIndex)
                {
                    // Set the value to the variable when the current upload context is existent.
                    isCurrentUploadinfoExist = true;
                    uploadInfo = tempUploadInfo;
                    currentUploadIndex = changeConnection.UploadContextContainer.IndexOf(tempUploadInfo);
                }
            }

            if (isCurrentUploadinfoExist)
            {
                // Set the upload information.
                uploadInfo.ImportDeleteflags = importDeleteFlag;
                AbstractFolder currentFolder = new AbstractFolder();

                // Record the current Folder Index
                int currentFolderIndex = 0;
                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if (tempFolder.FolderHandleIndex == uploadInfo.RelatedObjectHandleIndex)
                    {
                        // Set the value to the variable when the current Folder is existent.
                        currentFolder = tempFolder;
                        currentFolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                    }
                }

                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if ((tempFolder.ParentFolderIdIndex == currentFolder.FolderIdIndex) && objIdIndexes.Contains(tempFolder.FolderIdIndex))
                    {
                        // Remove current folder from FolderContainer and parent folder  when the parent Folder is existent.
                        changeConnection.FolderContainer = changeConnection.FolderContainer.Remove(tempFolder);
                        currentFolder.SubFolderIds = currentFolder.SubFolderIds.Remove(tempFolder.FolderIdIndex);
                    }

                    if (importDeleteFlag == (byte)ImportDeleteFlags.Hierarchy)
                    {
                        softDeleteFolderCount += 1;
                    }
                }

                foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
                {
                    if ((tempMessage.FolderIdIndex == currentFolder.FolderIdIndex) && objIdIndexes.Contains(tempMessage.MessageIdIndex))
                    {
                        // Remove current Message from MessageContainer and current folder when current Message is existent.
                        changeConnection.MessageContainer = changeConnection.MessageContainer.Remove(tempMessage);
                        currentFolder.MessageIds = currentFolder.MessageIds.Remove(tempMessage.MessageIdIndex);

                        if (importDeleteFlag == (byte)ImportDeleteFlags.delete)
                        {
                            softDeleteMessageCount += 1;
                        }
                    }
                }

                // Update the FolderContainer. 
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(currentFolderIndex, currentFolder);

                // Update the UploadContextContainer. 
                changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Update(currentUploadIndex, uploadInfo);
                connections[serverId] = changeConnection;

                result = RopResult.Success;

                ModelHelper.CaptureRequirement(
                    2449,
                    @"[In Uploading Changes Using ICS] Value is Success indicates No error occurred, or a conflict has been resolved.");

                // Because if the result is success means deletions of messages or folders into the server replica imported
                ModelHelper.CaptureRequirement(
                    884,
                    @"[In RopSynchronizationImportDeletes ROP] The RopSynchronizationImportDeletes ROP ([MS-OXCROPS] section 2.2.13.5) imports deletions of messages or folders into the server replica.");
                return result;
            }

            return result;
        }

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
        [Rule(Action = "SynchronizationImportHierarchyChange(serverId, uploadContextHandleIndex,parentFolderHandleIndex, properties, localFolderIdIndex, out folderIdIndex)/result")]
        public static RopResult SynchronizationImportHierarchyChange(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].UploadContextContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            folderIdIndex = -1;

            if (parentFolderHandleIndex == -1)
            {
                result = RopResult.NoParentFolder;
                ModelHelper.CaptureRequirement(
                    2450,
                    @"[In Uploading Changes Using ICS] Value is NoParentFolder indicates An attempt is being made to upload a hierarchy change for a folder whose parent folder does not yet exist.");
            }
            else
            {
                // Get ConnectionData value.
                ConnectionData changeConnection = connections[serverId];

                // Record current Upload Information.
                AbstractUploadInfo currentUploadInfo = new AbstractUploadInfo();

                // Identify current Upload handle.
                bool isCurrentUploadHandleExist = false;

                foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
                {
                    if (tempUploadInfo.UploadHandleIndex == uploadContextHandleIndex)
                    {
                        // Set the value to the variable when the current upload context is existent.
                        isCurrentUploadHandleExist = true;
                        currentUploadInfo = tempUploadInfo;
                    }
                }

                if (isCurrentUploadHandleExist)
                {
                    // Initialize the variable
                    AbstractFolder parentfolder = new AbstractFolder();
                    bool isParentFolderExist = false;
                    int parentfolderIndex = 0;
                    AbstractFolder currentFolder = new AbstractFolder();
                    bool isFolderExist = false;
                    int currentFolderIndex = 0;

                    // Research the local folder Id.
                    foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                    {
                        if (tempFolder.FolderIdIndex == localFolderIdIndex)
                        {
                            // Set the value to the  current Folder variable when the current folder  is existent.
                            isFolderExist = true;
                            currentFolder = tempFolder;
                            currentFolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                        }

                        if (tempFolder.FolderIdIndex == currentUploadInfo.RelatedObjectIdIndex)
                        {
                            // Set the value to the parent folder variable when the current parent folder is existent.
                            isParentFolderExist = true;
                            parentfolder = tempFolder;
                            parentfolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                        }
                    }

                    if (isFolderExist & isParentFolderExist)
                    {
                        foreach (string tempProperty in properties)
                        {
                            if (!currentFolder.FolderProperties.Contains(tempProperty))
                            {
                                // Add Property for folder
                                currentFolder.FolderProperties = currentFolder.FolderProperties.Add(tempProperty);
                            }
                        }

                        // Get the new change Number
                        currentFolder.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();

                        // Update the folder Container
                        changeConnection.FolderContainer = changeConnection.FolderContainer.Update(currentFolderIndex, currentFolder);
                    }
                    else
                    {
                        // Create a new folder
                        AbstractFolder newFolder = new AbstractFolder
                        {
                            FolderIdIndex = AdapterHelper.GetObjectIdIndex()
                        };

                        // Set new folder Id
                        folderIdIndex = newFolder.FolderIdIndex;

                        // Set value for new folder
                        newFolder.FolderProperties = properties;
                        newFolder.ParentFolderHandleIndex = parentfolder.FolderHandleIndex;
                        newFolder.ParentFolderIdIndex = parentfolder.FolderIdIndex;
                        newFolder.SubFolderIds = new Set<int>();
                        newFolder.MessageIds = new Set<int>();

                        // Add the new folder to parent folder
                        parentfolder.SubFolderIds = parentfolder.SubFolderIds.Add(newFolder.FolderIdIndex);
                        newFolder.FolderPermission = PermissionLevels.FolderOwner;
                        newFolder.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();
                        ModelHelper.CaptureRequirement(1897, "[In Identifying Objects and Maintaining Change Numbers]When a new object is created, it is assigned a change number.");

                        // Update FolderContainer information
                        changeConnection.FolderContainer = changeConnection.FolderContainer.Add(newFolder);
                        changeConnection.FolderContainer = changeConnection.FolderContainer.Update(parentfolderIndex, parentfolder);
                    }

                    // Return Success
                    connections[serverId] = changeConnection;

                    // Record RopSynchronizationImportHierarchyChange operation.
                    priorUploadOperation = PriorOperation.RopSynchronizationImportHierarchyChange;
                    result = RopResult.Success;

                    ModelHelper.CaptureRequirement(
                        2449,
                        @"[In Uploading Changes Using ICS] Value is Success indicates No error occurred, or a conflict has been resolved.");

                    // Because if the result is success means the folders or changes are imported.
                    ModelHelper.CaptureRequirement(
                        816,
                        @"[In RopSynchronizationImportHierarchyChange ROP] The RopSynchronizationImportHierarchyChange ROP ([MS-OXCROPS] section 2.2.13.4) is used to import new folders, or changes to existing folders, into the server replica.");
                }
            }

            return result;
        }

        /// <summary>
        /// Import new folders, or changes with conflict PCL to existing folders, into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Upload context handle.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="properties">Properties to be set.</param>
        /// <param name="localFolderIdIndex">Local folder id index</param>
        /// <param name="folderIdIndex">The folder object id index.</param>
        /// <param name="conflictType">The conflict type to generate.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SynchronizationImportHierarchyChangeWithConflict(serverId, uploadContextHandleIndex,parentFolderHandleIndex, properties, localFolderIdIndex, out folderIdIndex, conflictType)/result")]
        public static RopResult SynchronizationImportHierarchyChangeWithConflict(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex, ConflictTypes conflictType)
        {
            return SynchronizationImportHierarchyChange(serverId, uploadContextHandleIndex, parentFolderHandleIndex, properties, localFolderIdIndex, out folderIdIndex);
        }

        /// <summary>
        /// Import new messages or changes to existing messages into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">A synchronization upload context handle index.</param>
        /// <param name="messageIdindex">new client message id</param>
        /// <param name="importFlag">An 8-bit flag .</param>
        /// <param name="importMessageHandleIndex">The index of handle that indicate the Message object into which the client will upload the rest of the message changes.</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = "SynchronizationImportMessageChange(serverId, uploadContextHandleIndex,messageIdindex,importFlag, out importMessageHandleIndex)/result")]
        public static RopResult SynchronizationImportMessageChange(int serverId, int uploadContextHandleIndex, int messageIdindex, ImportFlag importFlag, out int importMessageHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].UploadContextContainer.Count > 0 && connections[serverId].LogonHandleIndex > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            importMessageHandleIndex = -1;

            if ((importFlag & ImportFlag.InvalidParameter) == ImportFlag.InvalidParameter &&
                requirementContainer.ContainsKey(3509001) && requirementContainer[3509001])
            {
                result = RopResult.InvalidParameter;
                ModelHelper.CaptureRequirement(
                    3509001,
                    @"[In Appendix A: Product Behavior]  If unknown flags are set, implementation does fail the operation. <40> Section 3.2.5.9.4.2: Exchange 2010, Exchange 2013 and Exchange 2016 fail the ROP [RopSynchronizationImportMessageChange] if unknown bit flags are set.");

                return result;
            }
            else if ((importFlag & ImportFlag.InvalidParameter) == ImportFlag.InvalidParameter &&
                requirementContainer.ContainsKey(350900201) && requirementContainer[350900201])
            {
                result = RopResult.Success;
                ModelHelper.CaptureRequirement(
                    350900201,
                    @"[In Appendix A: Product Behavior]  If unknown flags are set, implementation does not fail the operation. <41> Section 3.2.5.9.4.2: Exchange 2007 do not fail the ROP [RopSynchronizationImportMessageChange] if unknown bit flags are set.");
            }

            // Get ConnectionData value.
            ConnectionData changeConnection = connections[serverId];
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();

            // Identify whether the current Upload information is existent or not.
            bool isCurrentUploadinfoExist = false;

            // Record current Upload information.
            int currentUploadIndex = 0;

            foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
            {
                if (tempUploadInfo.UploadHandleIndex == uploadContextHandleIndex)
                {
                    // Set the value to the  current upload context variable when the current upload context  is existent.
                    isCurrentUploadinfoExist = true;
                    uploadInfo = tempUploadInfo;
                    currentUploadIndex = changeConnection.UploadContextContainer.IndexOf(tempUploadInfo);
                }
            }

            if (isCurrentUploadinfoExist)
            {
                // Create a new Message
                AbstractMessage currentMessage = new AbstractMessage();

                // Identify whether the current message is existent or not.
                bool isMessageExist = false;

                // Record the current Message.
                int currentMessageIndex = 0;
                foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
                {
                    if (tempMessage.MessageIdIndex == messageIdindex)
                    {
                        // Set the value to the variable when the message is existent.
                        isMessageExist = true;
                        currentMessage = tempMessage;
                        currentMessageIndex = changeConnection.MessageContainer.IndexOf(tempMessage);
                    }
                }

                if (isMessageExist)
                {
                    // Set new change number
                    currentMessage.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();
                    ModelHelper.CaptureRequirement(1898, "[In Identifying Objects and Maintaining Change Numbers]A new change number is assigned to a messaging object each time it is modified.");

                    // Update the MessageContainer
                    changeConnection.MessageContainer = changeConnection.MessageContainer.Update(currentMessageIndex, currentMessage);
                }
                else
                {
                    // Set the new message handle 
                    currentMessage.MessageHandleIndex = AdapterHelper.GetHandleIndex();

                    // Set property value of abstract message object 
                    currentMessage.FolderHandleIndex = uploadInfo.RelatedObjectHandleIndex;
                    currentMessage.FolderIdIndex = uploadInfo.RelatedObjectIdIndex;
                    currentMessage.MessageProperties = new Sequence<string>();
                    currentMessage.IsRead = true;
                    if ((importFlag & ImportFlag.Normal) == ImportFlag.Normal)
                    {
                        currentMessage.IsFAImessage = false;
                    }

                    if ((importFlag & ImportFlag.Associated) == ImportFlag.Associated)
                    {
                        currentMessage.IsFAImessage = true;

                        // When the Associated is set and the message being imported is an FAI message this requirement is captured.
                        ModelHelper.CaptureRequirement(
                            813,
                            @"[In RopSynchronizationImportMessageChange ROP Request Buffer] [ImportFlag,when the name is Associated, the value is 0x10] If this flag is set, the message being imported is an FAI message.");
                    }
                    else
                    {
                        currentMessage.IsFAImessage = false;

                        // When the Associated is not set and the message being imported is a normal message this requirement is captured.
                        ModelHelper.CaptureRequirement(
                            814,
                            @"[In RopSynchronizationImportMessageChange ROP Request Buffer] [ImportFlag,when the name is Associated, the value is 0x10] If this flag is not set, the message being imported is a normal message.");
                    }

                    // Out the  new message handle
                    importMessageHandleIndex = currentMessage.MessageHandleIndex;

                    // Because this is out messageHandle so the OutputServerObject is a Message object.
                    ModelHelper.CaptureRequirement(805, "[In RopSynchronizationImportMessageChange ROP Response Buffer]OutputServerObject: The value of this field MUST be the Message object into which the client will upload the rest of the message changes.");
                    currentMessage.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();
                    ModelHelper.CaptureRequirement(1897, "[In Identifying Objects and Maintaining Change Numbers]When a new object is created, it is assigned a change number.");

                    // Add new Message to MessageContainer
                    changeConnection.MessageContainer = changeConnection.MessageContainer.Add(currentMessage);

                    // Record the related FastTransferOperation for Upload Information.
                    uploadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.SynchronizationImportMessageChange;

                    // Update the UploadContextContainer.
                    changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Update(currentUploadIndex, uploadInfo);
                    connections[serverId] = changeConnection;

                    // Record RopSynchronizationImportMessageChange operation.
                    priorOperation = PriorOperation.RopSynchronizationImportMessageChange;
                    result = RopResult.Success;

                    ModelHelper.CaptureRequirement(
                    2449,
                    @"[In Uploading Changes Using ICS] Value is Success indicates No error occurred, or a conflict has been resolved.");

                    // Because if the result is success means the messages or changes are imported.
                    ModelHelper.CaptureRequirement(
                        782,
                        @"[In RopSynchronizationImportMessageChange ROP] The RopSynchronizationImportMessageChange ROP ([MS-OXCROPS] section 2.2.13.2) is used to import new messages or changes to existing messages into the server replica.");
                }
            }

            return result;
        }

        /// <summary>
        /// Imports message read state changes into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Sync handle.</param>
        /// <param name="messageHandleIndex">Message handle</param>
        /// <param name="ireadstatus">An array of MessageReadState structures one per each message that's changing its read state.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SynchronizationImportReadStateChanges(serverId, uploadContextHandleIndex,messageHandleIndex,ireadstatus)/result")]
        public static RopResult SynchronizationImportReadStateChanges(int serverId, int uploadContextHandleIndex, int messageHandleIndex, bool ireadstatus)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].MessageContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            ConnectionData changeConnection = connections[serverId];
            AbstractMessage currentMessage = new AbstractMessage();
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();

            // Identify whether the Message is existent or not and record the index.
            bool isCurrentMessageExist = false;
            int currentMessageindex = 0;

            // Identify whether the Upload information is existent or not and record the index.
            bool isCurrentUploadinfoExist = false;
            int currentUploadIndex = 0;

            foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
            {
                if (tempUploadInfo.UploadHandleIndex == uploadContextHandleIndex)
                {
                    // Set the value to the variable when the upload context is existent.
                    isCurrentUploadinfoExist = true;
                    uploadInfo = tempUploadInfo;
                    currentUploadIndex = changeConnection.UploadContextContainer.IndexOf(tempUploadInfo);
                }
            }

            foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
            {
                if (tempMessage.MessageHandleIndex == messageHandleIndex)
                {
                    // Set the value to the variable when the Message is existent.
                    isCurrentMessageExist = true;
                    currentMessage = tempMessage;
                    currentMessageindex = changeConnection.MessageContainer.IndexOf(tempMessage);
                }
            }

            if (isCurrentMessageExist)
            {
                // Find the parent folder of current message
                AbstractFolder parentfolder = new AbstractFolder();
                int parentfolderIndex = 0;
                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if (tempFolder.FolderHandleIndex == currentMessage.FolderHandleIndex)
                    {
                        parentfolder = tempFolder;
                        parentfolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                        if (parentfolder.FolderPermission == PermissionLevels.None)
                        {
                            return result = RopResult.AccessDenied;
                        }
                    }
                }
            }

            if (isCurrentUploadinfoExist && isCurrentMessageExist)
            {
                // Set the message read status value.
                if (currentMessage.IsRead != ireadstatus)
                {
                    currentMessage.IsRead = ireadstatus;

                    // Get read State changeNumber.
                    currentMessage.ReadStateChangeNumberIndex = ModelHelper.GetChangeNumberIndex();

                    ModelHelper.CaptureRequirement(
                         2260,
                         @"[In Receiving a RopSynchronizationImportReadStateChanges Request]Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated by adding the new change number to the MetaTagCnsetRead property (section 2.2.1.1.4).");
                }

                // Record the related Synchronization Operation.
                uploadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.SynchronizationReadStateChanges;

                // Update the upload context container and message container.
                changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Update(currentUploadIndex, uploadInfo);
                changeConnection.MessageContainer = changeConnection.MessageContainer.Update(currentMessageindex, currentMessage);
                connections[serverId] = changeConnection;
                result = RopResult.Success;

                ModelHelper.CaptureRequirement(
                        2449,
                        @"[In Uploading Changes Using ICS] Value is Success indicates No error occurred, or a conflict has been resolved.");

                // Because if the result is success means message read state changes is imported into the server replica.
                ModelHelper.CaptureRequirement(
                    905,
                    @"[In RopSynchronizationImportReadStateChanges ROP] The RopSynchronizationImportReadStateChanges ROP ([MS-OXCROPS] section 2.2.13.3) imports message read state changes into the server replica.");
            }

            return result;
        }

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
        [Rule(Action = "SynchronizationImportMessageMove(serverId, synchronizationUploadContextHandleIndex,sourceFolderIdIndex,destinationFolderIdIndex,sourceMessageIdIndex,sourceFolderHandleIndex,destinationFolderHandleIndex,inewerClientChange,out iolderversion,out icnpc)/result")]
        public static RopResult SynchronizationImportMessageMove(int serverId, int synchronizationUploadContextHandleIndex, int sourceFolderIdIndex, int destinationFolderIdIndex, int sourceMessageIdIndex, int sourceFolderHandleIndex, int destinationFolderHandleIndex, bool inewerClientChange, out bool iolderversion, out bool icnpc)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 1);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            iolderversion = false;
            icnpc = false;

            // Get the current ConnectionData value.
            ConnectionData changeConnection = connections[serverId];

            // Identify whether Current Upload information is existent or not.
            bool isCurrentUploadinfoExist = false;

            // Record the current Upload information.
            int currentUploadIndex = 0;
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();
            foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
            {
                if (tempUploadInfo.UploadHandleIndex == synchronizationUploadContextHandleIndex)
                {
                    // Set the value for current upload information when  current upload Information is existent.
                    isCurrentUploadinfoExist = true;
                    uploadInfo = tempUploadInfo;
                    currentUploadIndex = changeConnection.UploadContextContainer.IndexOf(tempUploadInfo);
                }
            }

            // Create variable of relate to source Folder.
            AbstractFolder sourceFolder = new AbstractFolder();
            bool isSourceFolderExist = false;
            int sourceFolderIndex = 0;

            // Create variable of relate to destination Folder.
            AbstractFolder destinationFolder = new AbstractFolder();
            bool isdestinationFolderExist = false;
            int destinationFolderIndex = 0;

            // Create a new message.
            AbstractMessage movedMessage = new AbstractMessage();

            // Identify whether the Moved Message is existent or not.
            bool isMovedMessageExist = false;
            int movedMessageIndex = 0;

            foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
            {
                if (tempFolder.FolderIdIndex == sourceFolderIdIndex && tempFolder.FolderHandleIndex == sourceFolderHandleIndex)
                {
                    // Set the value to the variable when the source folder is existent.
                    isSourceFolderExist = true;
                    sourceFolder = tempFolder;
                    sourceFolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                }
            }

            foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
            {
                if (tempFolder.FolderIdIndex == destinationFolderIdIndex && tempFolder.FolderHandleIndex == destinationFolderHandleIndex)
                {
                    // Set the value to the related variable when the destination folder is existent.
                    isdestinationFolderExist = true;
                    destinationFolder = tempFolder;
                    destinationFolderIndex = changeConnection.FolderContainer.IndexOf(tempFolder);
                }
            }

            foreach (AbstractMessage tempMessage in changeConnection.MessageContainer)
            {
                if (tempMessage.MessageIdIndex == sourceMessageIdIndex)
                {
                    // Set the value to the related variable when the source Message is existent.
                    isMovedMessageExist = true;
                    movedMessage = tempMessage;
                    movedMessageIndex = changeConnection.MessageContainer.IndexOf(tempMessage);
                }
            }

            if (isSourceFolderExist && isdestinationFolderExist && isMovedMessageExist && isCurrentUploadinfoExist)
            {
                // Set value for the new abstract message property.
                movedMessage.FolderIdIndex = destinationFolder.FolderIdIndex;
                movedMessage.FolderHandleIndex = destinationFolder.FolderHandleIndex;

                // Assigned a new change. 
                movedMessage.ChangeNumberIndex = ModelHelper.GetChangeNumberIndex();

                // Assigned a new message id.
                movedMessage.MessageIdIndex = AdapterHelper.GetObjectIdIndex();

                // Update message Container.
                changeConnection.MessageContainer = changeConnection.MessageContainer.Update(movedMessageIndex, movedMessage);

                // Remove the current message id from MessageIds of source folder.
                sourceFolder.MessageIds = sourceFolder.MessageIds.Remove(movedMessage.MessageIdIndex);
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(sourceFolderIndex, sourceFolder);

                // Remove the current message id from MessageIds of destination Folder.
                destinationFolder.MessageIds = destinationFolder.MessageIds.Add(movedMessage.MessageIdIndex);
                changeConnection.FolderContainer = changeConnection.FolderContainer.Update(destinationFolderIndex, destinationFolder);

                // Add information of Upload context
                uploadInfo.IsnewerClientChange = inewerClientChange;
                uploadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.SynchronizationImportMessageMove;

                // Update the upload context container
                changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Update(currentUploadIndex, uploadInfo);
                connections[serverId] = changeConnection;

                // Identify whether the IsnewerClientChange is true or false.
                if (uploadInfo.IsnewerClientChange == false)
                {
                    result = RopResult.Success;

                    ModelHelper.CaptureRequirement(
                        2449,
                        @"[In Uploading Changes Using ICS] Value is Success indicates No error occurred, or a conflict has been resolved.");

                    // Because if the result is success means the information about moving a message between two existing folders within the same mailbox imported
                    ModelHelper.CaptureRequirement(
                        839,
                        @"[In RopSynchronizationImportMessageMove ROP] The RopSynchronizationImportMessageMove ROP ([MS-OXCROPS] section 2.2.13.6) imports information about moving a message between two existing folders within the same mailbox.");
                }
                else
                {
                    // Set out put parameter value
                    iolderversion = true;
                    ModelHelper.CaptureRequirement(
                        875,
                        @"[In RopSynchronizationImportMessageMove ROP Response Buffer] [ 
Return value (4 bytes):] The following table[In section 2.2.3.2.4.4] contains additional return values[NewerClientChange] , if the ROP succeeded, but the server replica had an older version of a message than the local replica, the return value is 0x00040821.");
                    icnpc = true;
                    ModelHelper.CaptureRequirement(
                        876,
                        @"[In RopSynchronizationImportMessageMove ROP Response Buffer] [ 
Return value (4 bytes):] The following table[In section 2.2.3.2.4.4] contains additional return values[NewerClientChange] , if the values of the ChangeNumber and PredecessorChangeList fields, specified in section 2.2.3.2.4.4.1, were not applied to the destination message, the return value is 0x00040821.");
                    ModelHelper.CaptureRequirement(
                        1892,
                        "[In Identifying Objects and Maintaining Change Numbers]Copying of messaging objects within a mailbox or moving messages between folders of the same mailbox translates into creation of new messaging objects and therefore, new internal identifiers MUST be assigned to new copies.");
                    result = RopResult.NewerClientChange;
                }

                // Record RopSynchronizationImportMessageMove operation.
                priorUploadOperation = PriorOperation.RopSynchronizationImportMessageMove;
            }

            return result;
        }

        /// <summary>
        /// Creates a FastTransfer download context for a snapshot of the checkpoint ICS state of the operation identified by the given synchronization download context, or synchronization upload context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="syncHandleIndex">Synchronization context index.</param>
        /// <param name="downloadcontextHandleIndex">The index of FastTransfer download context for the ICS state.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SynchronizationGetTransferState(serverId, syncHandleIndex,out downloadcontextHandleIndex)/result")]
        public static RopResult SynchronizationGetTransferState(int serverId, int syncHandleIndex, out int downloadcontextHandleIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].UploadContextContainer.Count > 0 || connections[serverId].DownloadContextContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            downloadcontextHandleIndex = -1;

            // Get the current ConnectionData value.
            ConnectionData changeConnection = connections[serverId];
            AbstractDownloadInfo downloadInfo = new AbstractDownloadInfo();
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();

            // Identify whether the Download context or Upload context is existent or not and record the index.
            bool isCurrentDownloadInfoExist = false;
            bool isCurrentUploadInfoExist = false;

            foreach (AbstractDownloadInfo temp in changeConnection.DownloadContextContainer)
            {
                if (temp.DownloadHandleIndex == syncHandleIndex)
                {
                    // Set the value to the related variable when the current Download context is existent.
                    isCurrentDownloadInfoExist = true;
                    downloadInfo = temp;
                }
            }

            if (!isCurrentDownloadInfoExist)
            {
                foreach (AbstractUploadInfo tempInfo in changeConnection.UploadContextContainer)
                {
                    if (tempInfo.UploadHandleIndex == syncHandleIndex)
                    {
                        // Set the value to the related variable when the current upload context is existent.
                        isCurrentUploadInfoExist = true;
                        uploadInfo = tempInfo;
                    }
                }
            }

            if (isCurrentDownloadInfoExist || isCurrentUploadInfoExist)
            {
                // Create a new download context.
                AbstractDownloadInfo newDownloadInfo = new AbstractDownloadInfo
                {
                    DownloadHandleIndex = AdapterHelper.GetHandleIndex()
                };

                // Out the new downloadHandle.
                downloadcontextHandleIndex = newDownloadInfo.DownloadHandleIndex;
                ModelHelper.CaptureRequirement(765, "[In RopSynchronizationGetTransferState ROP Response Buffer]OutputServerObject: The value of this field MUST be the FastTransfer download context for the ICS state.");
                if (isCurrentDownloadInfoExist)
                {
                    // Set the new Download context value
                    newDownloadInfo.RelatedObjectHandleIndex = downloadInfo.RelatedObjectHandleIndex;
                    newDownloadInfo.SynchronizationType = downloadInfo.SynchronizationType;
                    newDownloadInfo.UpdatedState = downloadInfo.UpdatedState;
                    priorDownloadOperation = PriorDownloadOperation.RopSynchronizationGetTransferState;
                }
                else
                {
                    // Set the new Upload context value
                    newDownloadInfo.RelatedObjectHandleIndex = uploadInfo.RelatedObjectHandleIndex;
                    newDownloadInfo.SynchronizationType = uploadInfo.SynchronizationType;
                    newDownloadInfo.UpdatedState = uploadInfo.UpdatedState;
                }

                // Set the abstractFastTransferStreamType for new  Down loadContext value
                newDownloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.state;
                newDownloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.SynchronizationGetTransferState;

                // Add new download context to DownloadContextContainer.
                changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(newDownloadInfo);
                connections[serverId] = changeConnection;
                result = RopResult.Success;

                // Because context created if the RopSynchronizationGetTransferState execute successful.
                ModelHelper.CaptureRequirement(
                    758,
                    @"[In RopSynchronizationGetTransferState ROP] The RopSynchronizationGetTransferState ROP ([MS-OXCROPS] section 2.2.13.8) creates a FastTransfer download context for the checkpoint ICS state of the operation identified by the given synchronization download context or synchronization upload context at the current moment in time.");
            }

            return result;
        }

        /// <summary>
        /// Upload of an ICS state property into the synchronization context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">The synchronization context handle</param>
        /// <param name="icsPropertyType">Property tags of the ICS state property.</param>
        /// <param name="isPidTagIdsetGivenInputAsInter32"> identifies Property tags as PtypInteger32.</param>
        /// <param name="icsStateIndex">The index of the ICS State.</param>
        /// <returns>The ICS state property is upload to the server successfully or not.</returns>
        [Rule(Action = "SynchronizationUploadState(serverId, uploadContextHandleIndex, icsPropertyType, isPidTagIdsetGivenInputAsInter32, icsStateIndex)/result")]
        public static RopResult SynchronizationUploadState(int serverId, int uploadContextHandleIndex, ICSStateProperties icsPropertyType, bool isPidTagIdsetGivenInputAsInter32, int icsStateIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].DownloadContextContainer.Count > 0 || connections[serverId].UploadContextContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();
            AbstractDownloadInfo downLoadInfo = new AbstractDownloadInfo();

            // Get the current ConnectionData value.
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the DownloadContext or UploadContext is existent or not and record the index.
            bool isCurrentUploadinfoExist = false;
            int currentUploadIndex = 0;
            bool isCurrentDownLoadinfoExist = false;
            int currentDownLoadIndex = 0;

            foreach (AbstractUploadInfo tempUploadInfo in changeConnection.UploadContextContainer)
            {
                if (tempUploadInfo.UploadHandleIndex == uploadContextHandleIndex)
                {
                    // Set the value to the related variable when the current upload context is existent.
                    isCurrentUploadinfoExist = true;
                    uploadInfo = tempUploadInfo;
                    currentUploadIndex = changeConnection.UploadContextContainer.IndexOf(tempUploadInfo);
                    break;
                }
            }

            if (!isCurrentUploadinfoExist)
            {
                foreach (AbstractDownloadInfo tempDownLoadInfo in changeConnection.DownloadContextContainer)
                {
                    if (tempDownLoadInfo.DownloadHandleIndex == uploadContextHandleIndex)
                    {
                        // Set the value to the related variable when the current Download context is existent.
                        isCurrentDownLoadinfoExist = true;
                        downLoadInfo = tempDownLoadInfo;
                        currentDownLoadIndex = changeConnection.DownloadContextContainer.IndexOf(tempDownLoadInfo);
                        break;
                    }
                }
            }

            if (isCurrentDownLoadinfoExist || isCurrentUploadinfoExist)
            {
                if (isCurrentUploadinfoExist)
                {
                    if (icsStateIndex != 0)
                    {
                        AbstractFolder currentFolder = new AbstractFolder();
                        foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                        {
                            if (tempFolder.FolderHandleIndex == uploadInfo.RelatedObjectHandleIndex)
                            {
                                // Set the value to the related variable when the current Folder is existent.
                                currentFolder = tempFolder;
                                Condition.IsTrue(currentFolder.ICSStateContainer.ContainsKey(icsStateIndex));
                            }
                        }

                        // Add ICS State to ICSStateContainer of current folder.
                        AbstractUpdatedState updatedState = currentFolder.ICSStateContainer[icsStateIndex];
                        switch (icsPropertyType)
                        {
                            case ICSStateProperties.PidTagIdsetGiven:

                                // Set IdsetGiven value
                                uploadInfo.UpdatedState.IdsetGiven = updatedState.IdsetGiven;
                                break;
                            case ICSStateProperties.PidTagCnsetRead:

                                // Set CnsetRead value
                                uploadInfo.UpdatedState.CnsetRead = updatedState.CnsetRead;
                                break;
                            case ICSStateProperties.PidTagCnsetSeen:

                                // Set CnsetSeen value
                                uploadInfo.UpdatedState.CnsetSeen = updatedState.CnsetSeen;
                                break;
                            case ICSStateProperties.PidTagCnsetSeenFAI:

                                // Set CnsetSeenFAI value
                                uploadInfo.UpdatedState.CnsetSeenFAI = updatedState.CnsetSeenFAI;
                                break;
                            default:
                                break;
                        }

                        // Update the UploadContextContainer context.
                        changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Update(currentUploadIndex, uploadInfo);
                    }
                }
                else
                {
                    if (icsStateIndex != 0)
                    {
                        AbstractFolder currentFolder = new AbstractFolder();
                        foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                        {
                            if (tempFolder.FolderHandleIndex == downLoadInfo.RelatedObjectHandleIndex)
                            {
                                // Set the value to the related variable when the current Folder is existent.
                                currentFolder = tempFolder;

                                // Identify ICS State whether exist index or not  in ICSStateContainer
                                Condition.IsTrue(currentFolder.ICSStateContainer.ContainsKey(icsStateIndex));
                            }
                        }

                        // Add update state to ICSStateContainer of current folder.
                        AbstractUpdatedState updatedState = currentFolder.ICSStateContainer[icsStateIndex];
                        switch (icsPropertyType)
                        {
                            case ICSStateProperties.PidTagIdsetGiven:

                                // Set IdsetGiven value.
                                downLoadInfo.UpdatedState.IdsetGiven = updatedState.IdsetGiven;
                                break;
                            case ICSStateProperties.PidTagCnsetRead:

                                // Set CnsetRead value.
                                downLoadInfo.UpdatedState.CnsetRead = updatedState.CnsetRead;
                                break;
                            case ICSStateProperties.PidTagCnsetSeen:

                                // Set CnsetSeen value.
                                downLoadInfo.UpdatedState.CnsetSeen = updatedState.CnsetSeen;
                                break;
                            case ICSStateProperties.PidTagCnsetSeenFAI:

                                // Set CnsetSeenFAI value.
                                downLoadInfo.UpdatedState.CnsetSeenFAI = updatedState.CnsetSeenFAI;
                                break;
                            default:
                                break;
                        }

                        // Update the DownloadContextContainer.
                        changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Update(currentDownLoadIndex, downLoadInfo);
                    }
                }

                connections[serverId] = changeConnection;
                if (isPidTagIdsetGivenInputAsInter32)
                {
                    // Identify the property tag whether PtypInteger32 is or not.
                    if (requirementContainer.Keys.Contains(2657) && requirementContainer[2657])
                    {
                        result = RopResult.Success;
                        ModelHelper.CaptureRequirement(2657, "[In Receiving the MetaTagIdsetGiven ICS State Property] Implementation does accept this MetaTagIdsetGiven property when the property tag identifies it as PtypInteger32. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                    else
                    {
                        return result;
                    }
                }

                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Allocates a range of internal identifiers for the purpose of assigning them to client-originated objects in a local replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="logonHandleIndex">The server object handle index.</param>
        /// <param name="idcount">An unsigned 32-bit integer specifies the number of IDs to allocate.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "GetLocalReplicaIds(serverId, logonHandleIndex,idcount)/result")]
        public static RopResult GetLocalReplicaIds(int serverId, int logonHandleIndex, uint idcount)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].LogonHandleIndex > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // Get the current ConnectionData value.
            ConnectionData changeConnection = connections[serverId];
            if (logonHandleIndex == changeConnection.LogonHandleIndex)
            {
                // Set localId Count value.
                changeConnection.LocalIdCount = idcount;
                result = RopResult.Success;
                connections[serverId] = changeConnection;

                if (idcount > 0)
                {
                    // Because only if result is success and the idcount larger than 0 indicate a range of internal identifiers (2) for the purpose of assigning them to client-originated objects in a local replica are allocated.
                    ModelHelper.CaptureRequirement(
                        925,
                        @"[In RopGetLocalReplicaIds ROP] The RopGetLocalReplicaIds ROP ([MS-OXCROPS] section 2.2.13.13) allocates a range of internal identifiers for the purpose of assigning them to client-originated objects in a local replica.");
                }
            }

            return result;
        }

        /// <summary>
        /// Identifies that a set of IDs either belongs to deleted messages in the specified folder or will never be used for any messages in the specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">A Folder object handle</param>
        /// <param name="longTermIdRangeIndex">The range of LongTermId.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = "SetLocalReplicaMidsetDeleted(serverId, folderHandleIndex, longTermIdRangeIndex)/result")]
        public static RopResult SetLocalReplicaMidsetDeleted(int serverId, int folderHandleIndex, Sequence<int> longTermIdRangeIndex)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            Condition.IsTrue(connections[serverId].FolderContainer.Count > 0);

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // Get the current ConnectionData value
            ConnectionData changeConnection = connections[serverId];

            // Identify whether the current folder is existent or not and record the index.
            bool isCurrentFolderExist = false;

            foreach (AbstractFolder tempfolder in changeConnection.FolderContainer)
            {
                if (tempfolder.FolderHandleIndex == folderHandleIndex)
                {
                    // Set the value to the related variable when the current Folder is existent.
                    isCurrentFolderExist = true;
                }
            }

            if (isCurrentFolderExist == false)
            {
                // The server return invalid parameter when current folder is not exist.
                result = RopResult.InvalidParameter;
            }
            else
            {
                // The server return Success.
                result = RopResult.Success;

                // When the ROP success means server add ranges of IDs supplied through this ROP to the deleted item list.
                ModelHelper.CaptureRequirement(
                    2269,
                    @"[In Receiving a RopSetLocalReplicaMidsetDeleted Request] A server MUST add ranges of IDs supplied through this ROP to the deleted item list.");

                ModelHelper.CaptureRequirement(
                    940,
                    @"[In RopSetLocalReplicaMidsetDeleted ROP] The RopSetLocalReplicaMidsetDeleted ROP ([MS-OXCROPS] section 2.2.13.12) identifies that a set of IDs either belongs to deleted messages in the specified folder or will never be used for any messages in the specified folder.");
            }

            return result;
        }

        #endregion

        #region FastTransfer related actions

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant subObjects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">The input handle type</param>
        /// <param name="level">Variable indicate whether copy the descendant subObjects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation .</param>
        /// <param name="propertyTags">Array of properties and subObjects to exclude.</param>
        /// <param name="downloadContextHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("FastTransferSourceCopyTo(serverId,objHandleIndex,handleType,level,copyFlag,option,propertyTags,out downloadContextHandleIndex)/result"))]
        public static RopResult FastTransferSourceCopyTo(int serverId, int objHandleIndex, InputHandleType handleType, bool level, CopyToCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int downloadContextHandleIndex)
        {
            // The contraction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // The copyFlag conditions.
            if (((copyFlag == CopyToCopyFlags.Invalid && (requirementContainer.ContainsKey(3445) && requirementContainer[3445])) || (option == SendOptionAlls.Invalid && (requirementContainer.ContainsKey(3463) && requirementContainer[3463]))) ||
                ((copyFlag & CopyToCopyFlags.Move) == CopyToCopyFlags.Move && (requirementContainer.ContainsKey(3442001) && requirementContainer[3442001])) ||
                ((copyFlag & CopyToCopyFlags.Move) == CopyToCopyFlags.Move && (requirementContainer.ContainsKey(3442003) && requirementContainer[3442003])))
            {
                downloadContextHandleIndex = -1;
                if ((copyFlag & CopyToCopyFlags.Move) == CopyToCopyFlags.Move)
                {
                    // CopyToCopyFlags value is Move.
                    if (requirementContainer.ContainsKey(3442001) && requirementContainer[3442001])
                    {
                        // When the ROP return invalid parameter this requirement verified.
                        ModelHelper.CaptureRequirement(
                        3442001,
                        @"[In Appendix A: Product Behavior] Implementation does not support. <31> Section 3.2.5.8.1.1: Exchange 2010 and Exchange 2013 do not support the Move flag for the RopFastTransferSourceCopyTo ROP (section 2.2.3.1.1.1).");
                    }

                    if (requirementContainer.ContainsKey(3442003) && requirementContainer[3442003])
                    {
                        result = RopResult.InvalidParameter;

                        // When the ROP return invalid parameter this requirement verified.
                        ModelHelper.CaptureRequirement(
                        3442003,
                        @"[In Appendix A: Product Behavior] If the server receives the Move flag, implementation does fail the operation with an error code InvalidParameter (0x80070057).  <31> Section 3.2.5.8.1.1: The server sets the value of the ReturnValue field to InvalidParameter (0x80070057) if it receives this flag [Move flag].(Microsoft Exchange Server 2010, Exchange Server 2013 and  Exchange Server 2016 follow this behavior.)");
                    }
                }

                return result;
            }
            else
            {
                // Create a new download context
                AbstractDownloadInfo downloadInfo = new AbstractDownloadInfo();
                bool isObjExist = false;

                // Get value of ConnectionData
                ConnectionData changeConnection = connections[serverId];
                connections.Remove(serverId);

                // Find current message
                if (handleType == InputHandleType.MessageHandle)
                {
                    foreach (AbstractMessage temp in changeConnection.MessageContainer)
                    {
                        if (temp.MessageHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadInfo.DownloadHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadContextHandleIndex = downloadInfo.DownloadHandleIndex;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.CopyToCopyFlag = copyFlag;
                    downloadInfo.IsLevelTrue = level;

                    // Record FastTransfer Operation.
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyTo;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.MessageContent;
                    downloadInfo.ObjectType = ObjectType.Message;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;

                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyTo;

                    // Add new download context to downloadContext Container.
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }
                else if (handleType == InputHandleType.FolderHandle)
                {
                    // Find current folder.
                    foreach (AbstractFolder temp in changeConnection.FolderContainer)
                    {
                        if (temp.FolderHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadInfo.DownloadHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadContextHandleIndex = downloadInfo.DownloadHandleIndex;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.CopyToCopyFlag = copyFlag;

                    // Record FastTransfer Operation
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyTo;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.folderContent;
                    downloadInfo.ObjectType = ObjectType.Folder;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;

                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyTo;

                    // Add new download context to DownloadContextContainer
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }
                else
                {
                    // Find current attachment
                    foreach (AbstractAttachment temp in changeConnection.AttachmentContainer)
                    {
                        if (temp.AttachmentHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadInfo.DownloadHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadContextHandleIndex = downloadInfo.DownloadHandleIndex;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.CopyToCopyFlag = copyFlag;

                    // Record FastTransfer Operation
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyTo;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.attachmentContent;
                    downloadInfo.ObjectType = ObjectType.Attachment;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;

                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyTo;

                    // Add new download context to DownloadContextContainer
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }

                connections.Add(serverId, changeConnection);

                result = RopResult.Success;
                ModelHelper.CaptureRequirement(
                    361,
                    @"[In RopFastTransferSourceCopyTo ROP] The RopFastTransferSourceCopyTo ROP ([MS-OXCROPS] section 2.2.12.6) initializes a FastTransfer operation to download content from a given messaging object and its descendant subobjects.");

                if ((copyFlag & CopyToCopyFlags.Move) == CopyToCopyFlags.Move)
                {
                    if (requirementContainer.ContainsKey(3442002) && requirementContainer[3442002])
                    {
                        ModelHelper.CaptureRequirement(
                            3442002,
                            @"[In Appendix A: Product Behavior] Implementation does support Move flag [for the RopFastTransferSourceCopyTo ROP]. (Microsoft Exchange Server 2007 follow this behavior.)");
                    }

                    if (requirementContainer.ContainsKey(3442004) && requirementContainer[3442004])
                    {
                        ModelHelper.CaptureRequirement(
                            3442004,
                            @"[In Appendix A: Product Behavior] If the server receives the Move flag, implementation does not fail the operation.(<31> Section 3.2.5.8.1.1: Microsoft Exchange Server 2007 follows this behavior.)");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant subObjects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">Input Handle Type</param>
        /// <param name="level">Variable indicate whether copy the descendant subObjects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation .</param>
        /// <param name="propertyTags">The list of properties and subObjects to exclude.</param>
        /// <param name="downloadContextHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = ("FastTransferSourceCopyProperties(serverId,objHandleIndex,handleType,level,copyFlag,option,propertyTags,out downloadContextHandleIndex)/result"))]
        public static RopResult FastTransferSourceCopyProperties(int serverId, int objHandleIndex, InputHandleType handleType, bool level, CopyPropertiesCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int downloadContextHandleIndex)
        {
            // The contraction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // SendOptionAll value is Invalid parameter 
            if ((option == SendOptionAlls.Invalid && (requirementContainer.ContainsKey(3470) && requirementContainer[3470])) ||
                (copyFlag == CopyPropertiesCopyFlags.Invalid && (requirementContainer.ContainsKey(3466) && requirementContainer[3466])))
            {
                downloadContextHandleIndex = -1;
            }
            else if (((copyFlag & CopyPropertiesCopyFlags.Move) == CopyPropertiesCopyFlags.Move) && (requirementContainer.ContainsKey(3466) && requirementContainer[3466]))
            {
                // CopyPropertiesCopyFlags value is Move.
                result = RopResult.NotImplemented;
                downloadContextHandleIndex = -1;
            }
            else
            {
                // Create a new download context.
                AbstractDownloadInfo downloadInfo = new AbstractDownloadInfo();
                ConnectionData changeConnection = connections[serverId];
                bool isObjExist = false;

                connections.Remove(serverId);
                if (handleType == InputHandleType.MessageHandle)
                {
                    foreach (AbstractMessage temp in changeConnection.MessageContainer)
                    {
                        if (temp.MessageHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadContextHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadInfo.DownloadHandleIndex = downloadContextHandleIndex;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.MessageContent;

                    // Record the FastTransferOperation
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyProperties;
                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyProperties;

                    // Set value for new download context.
                    downloadInfo.CopyPropertiesCopyFlag = copyFlag;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;
                    downloadInfo.ObjectType = ObjectType.Message;
                    downloadInfo.IsLevelTrue = level;

                    // Add new download context to DownloadContextContainer.
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }
                else if (handleType == InputHandleType.FolderHandle)
                {
                    // Find current folder.
                    foreach (AbstractFolder temp in changeConnection.FolderContainer)
                    {
                        if (temp.FolderHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadContextHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadInfo.DownloadHandleIndex = downloadContextHandleIndex;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.folderContent;

                    // Record the FastTransferOperation
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyProperties;
                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyProperties;

                    // Set value for new download context.
                    downloadInfo.CopyPropertiesCopyFlag = copyFlag;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;
                    downloadInfo.ObjectType = ObjectType.Folder;

                    // Add new download context to DownloadContextContainer.
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }
                else
                {
                    // Find the current Attachment
                    foreach (AbstractAttachment temp in changeConnection.AttachmentContainer)
                    {
                        if (temp.AttachmentHandleIndex == objHandleIndex)
                        {
                            isObjExist = true;
                        }
                    }

                    Condition.IsTrue(isObjExist);

                    // Set value for new download context.
                    downloadContextHandleIndex = AdapterHelper.GetHandleIndex();
                    downloadInfo.DownloadHandleIndex = downloadContextHandleIndex;
                    downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.attachmentContent;

                    // Record the FastTransferOperation
                    downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyProperties;
                    priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyProperties;

                    // Set value for new download context.
                    downloadInfo.CopyPropertiesCopyFlag = copyFlag;
                    downloadInfo.Property = propertyTags;
                    downloadInfo.Sendoptions = option;
                    downloadInfo.ObjectType = ObjectType.Attachment;
                    downloadInfo.RelatedObjectHandleIndex = objHandleIndex;

                    // Add new download context to DownloadContextContainer.
                    changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                }

                connections.Add(serverId, changeConnection);
                result = RopResult.Success;

                // If the server returns success result, it means the RopFastTransferSourceCopyProperties ROP initializes the FastTransfer operation successfully. And then this requirement can be captured.
                ModelHelper.CaptureRequirement(
                    431,
                    @"[In RopFastTransferSourceCopyProperties ROP] The RopFastTransferSourceCopyProperties ROP ([MS-OXCROPS] section 2.2.12.7) initializes a FastTransfer operation to download content from a specified messaging object and its descendant sub objects.");
            }

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation on a folder for downloading content and descendant subObjects for messages identified by a given set of IDs.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="messageIds">The list of MIDs the messages should copy.</param>
        /// <param name="downloadContextHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("FastTransferSourceCopyMessages(serverId,objHandleIndex,copyFlag,option,messageIds,out downloadContextHandleIndex)/result"))]
        public static RopResult FastTransferSourceCopyMessages(int serverId, int objHandleIndex, RopFastTransferSourceCopyMessagesCopyFlags copyFlag, SendOptionAlls option, Sequence<int> messageIds, out int downloadContextHandleIndex)
        {
            // The contraction conditions
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            if (option == SendOptionAlls.Invalid)
            {
                if (requirementContainer.ContainsKey(3479) && requirementContainer[3479])
                {
                    // SendOption flags value is invalid
                    downloadContextHandleIndex = -1;
                    return result;
                }
            }

            // Modify the logical
            if ((copyFlag & RopFastTransferSourceCopyMessagesCopyFlags.Unused3) == RopFastTransferSourceCopyMessagesCopyFlags.Unused3)
            {
                // CopyFlags is set to Unused3
                downloadContextHandleIndex = -1;
            }
            else
            {
                // Identify whether the current folder is existent or not.
                ConnectionData changeConnection = connections[serverId];
                bool isFolderExist = false;

                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if (tempFolder.FolderHandleIndex == objHandleIndex)
                    {
                        isFolderExist = true;
                    }
                }

                Condition.IsTrue(isFolderExist);

                // Set value for new download context.
                AbstractDownloadInfo downloadInfo = new AbstractDownloadInfo();
                downloadContextHandleIndex = AdapterHelper.GetHandleIndex();
                downloadInfo.DownloadHandleIndex = downloadContextHandleIndex;

                connections.Remove(serverId);
                downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.MessageList;
                downloadInfo.CopyMessageCopyFlag = copyFlag;

                // Record the FastTransferOperation
                downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyMessage;
                priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyMessage;

                // Set value for new download context.
                downloadInfo.Sendoptions = option;
                downloadInfo.RelatedObjectHandleIndex = objHandleIndex;
                downloadInfo.ObjectType = ObjectType.Folder;

                // Add new download context to DownloadContextContainer.
                changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                connections.Add(serverId, changeConnection);
                priorOperation = MS_OXCFXICS.PriorOperation.RopFastTransferSourceCopyMessage;
                result = RopResult.Success;

                // If the server returns success result, it means the RopFastTransferSourceCopyMessages ROP initializes the FastTransfer operation successfully. And then this requirement can be captured.
                ModelHelper.CaptureRequirement(
                    3125,
                    @"[In RopFastTransferSourceCopyMessages ROP] The RopFastTransferSourceCopyMessages ROP ([MS-OXCROPS] section 2.2.12.5) initializes a FastTransfer operation on a folder for downloading content and descendant subobjects of messages identified by a set of MID structures ([MS-OXCDATA] section 2.2.1.2).");
            }

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation to download properties and descendant subObjects for a specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="downloadContextHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("FastTransferSourceCopyFolder(serverId,folderHandleIndex,copyFlag,option, out downloadContextHandleIndex)/result"))]
        public static RopResult FastTransferSourceCopyFolder(int serverId, int folderHandleIndex, CopyFolderCopyFlags copyFlag, SendOptionAlls option, out int downloadContextHandleIndex)
        {
            // The contraction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;

            // Modify the logical
            if ((option == SendOptionAlls.Invalid && (requirementContainer.ContainsKey(3487) && requirementContainer[3487])) ||
                (copyFlag == CopyFolderCopyFlags.Invalid && (requirementContainer.ContainsKey(3483) && requirementContainer[3483])))
            {
                // SendOption is Invalid parameter and CopyFolderCopyFlags is Invalid parameter.
                downloadContextHandleIndex = -1;
                return result;
            }
            else if (copyFlag == CopyFolderCopyFlags.Move && (requirementContainer.ContainsKey(526001) && !requirementContainer[526001]))
            {
                downloadContextHandleIndex = -1;
                ModelHelper.CaptureRequirement(
                            526001,
                            @"[In Appendix A: Product Behavior] [CopyFlags] [When the flag name is Move, value is 0x01] Implementation does set the Move flag on a download operation to indicate the following: The server does not output any objects in a FastTransfer stream that the client does not have permissions to delete. <7> Section 2.2.3.1.1.4.1: In Exchange 2007, the Move bit flag is read by the server.");

                return result;
            }
            else
            {
                ConnectionData changeConnection = connections[serverId];

                // Identify whether the current folder is existent or not.
                bool isFolderExist = false;
                foreach (AbstractFolder tempFolder in changeConnection.FolderContainer)
                {
                    if (tempFolder.FolderHandleIndex == folderHandleIndex)
                    {
                        isFolderExist = true;
                    }
                }

                Condition.IsTrue(isFolderExist);

                // Create a new download context.
                AbstractDownloadInfo downloadInfo = new AbstractDownloadInfo();
                downloadContextHandleIndex = AdapterHelper.GetHandleIndex();
                downloadInfo.DownloadHandleIndex = downloadContextHandleIndex;
                connections.Remove(serverId);

                // Record the FastTransferOperation and Stream Type.
                downloadInfo.AbstractFastTransferStreamType = FastTransferStreamType.TopFolder;
                downloadInfo.RelatedFastTransferOperation = EnumFastTransferOperation.FastTransferSourceCopyFolder;
                priorDownloadOperation = PriorDownloadOperation.RopFastTransferSourceCopyFolder;

                // Set value for new download context.
                downloadInfo.CopyFolderCopyFlag = copyFlag;
                downloadInfo.Sendoptions = option;
                downloadInfo.ObjectType = ObjectType.Folder;
                downloadInfo.RelatedObjectHandleIndex = folderHandleIndex;

                // Add new download context to downloadContainer.
                changeConnection.DownloadContextContainer = changeConnection.DownloadContextContainer.Add(downloadInfo);
                connections.Add(serverId, changeConnection);

                result = RopResult.Success;

                // If the server returns success result, it means the RopFastTransferSourceCopyFolder ROP initializes the FastTransfer operation successfully. And then this requirement can be captured.
                ModelHelper.CaptureRequirement(
                            502,
                            @"[In RopFastTransferSourceCopyFolder ROP] The RopFastTransferSourceCopyFolder ROP ([MS-OXCROPS] section 2.2.12.4) initializes a FastTransfer operation to download properties and descendant subobjects for a specified folder.");
            }

            return result;
        }

        /// <summary>
        /// Downloads the next portion of a FastTransfer stream.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="downloadHandleIndex">A fastTransfer stream object handle index. </param>
        /// <param name="bufferSize">Specifies the maximum amount of data to be output in the TransferBuffer.</param>
        /// <param name="transferBufferIndex">The index of data get from the fastTransfer stream.</param>
        /// <param name="abstractFastTransferStream">The abstractFastTransferStream.</param>
        /// <param name="transferDataSmallOrEqualToBufferSize">Variable to not if the transferData is small or equal to bufferSize</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("FastTransferSourceGetBuffer(serverId,downloadHandleIndex,bufferSize,out transferBufferIndex,out abstractFastTransferStream ,out transferDataSmallOrEqualToBufferSize)/result"))]
        public static RopResult FastTransferSourceGetBuffer(int serverId, int downloadHandleIndex, BufferSize bufferSize, out int transferBufferIndex, out AbstractFastTransferStream abstractFastTransferStream, out bool transferDataSmallOrEqualToBufferSize)
        {
            // The contractions conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize the return value.
            RopResult result = RopResult.InvalidParameter;
            transferBufferIndex = -1;
            abstractFastTransferStream = new AbstractFastTransferStream();
            transferDataSmallOrEqualToBufferSize = false;

            if (bufferSize == BufferSize.Greater)
            {
                result = RopResult.BufferTooSmall;
                return result;
            }

            // Get current ConnectionData value.
            ConnectionData currentConnection = connections[serverId];

            // Create a new currentDownloadContext.
            AbstractDownloadInfo currentDownloadContext = new AbstractDownloadInfo();

            // Identify whether the Download context is existent or not.
            bool isDownloadHandleExist = false;

            // Find the current Download Context
            foreach (AbstractDownloadInfo tempDownloadContext in currentConnection.DownloadContextContainer)
            {
                if (tempDownloadContext.DownloadHandleIndex == downloadHandleIndex)
                {
                    // Set the value to the related variable when the download context is existent.
                    isDownloadHandleExist = true;
                    result = RopResult.Success;
                    currentDownloadContext = tempDownloadContext;
                    int infoIndex = currentConnection.DownloadContextContainer.IndexOf(tempDownloadContext);

                    // Get the Data buffer index
                    transferBufferIndex = AdapterHelper.GetStreamBufferIndex();
                    currentDownloadContext.DownloadStreamIndex = transferBufferIndex;
                    abstractFastTransferStream.StreamType = currentDownloadContext.AbstractFastTransferStreamType;
                    currentConnection.DownloadContextContainer = currentConnection.DownloadContextContainer.Update(infoIndex, currentDownloadContext);
                }
            }

            // Create new variable relate to current folder.
            AbstractFolder currentFolder = new AbstractFolder();
            int currentFolderIndex = 0;

            // Identify current whether DownloadHandle is existent or not.
            if (isDownloadHandleExist)
            {
                // If bufferSize is set to a value other than 0xBABE
                if (bufferSize != BufferSize.Normal)
                {
                    transferDataSmallOrEqualToBufferSize = true;

                    ModelHelper.CaptureRequirement(
                        2142,
                        @"[In Receiving a RopFastTransferSourceGetBuffer Request]If the value of BufferSize in the ROP request is set to a value other than 0xBABE, the following semantics apply:The server MUST output, at most, the number of bytes specified by the BufferSize field in the 
                        TransferBuffer field even if more data is available.");
                    ModelHelper.CaptureRequirement(
                        2143,
                        @"[In Receiving a RopFastTransferSourceGetBuffer Request]If the value of BufferSize in the ROP request is set to a value other than 
                        0xBABE, the following semantics apply:The server returns less bytes than the value specified by the BufferSize field, or the server 
                        returns the number of bytes specified by the BufferSize field in the TransferBuffer field.");
                }

                #region Requirements about RopOperation Response

                // FolderHandleIndex is the Index of the FastTransfer download context
                if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                {
                    ModelHelper.CaptureRequirement(384, "[In RopFastTransferSourceCopyTo ROP Response Buffer] OutputServerObject: The value of this field MUST be the FastTransfer download context.");
                }
                else if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                {
                    ModelHelper.CaptureRequirement(455, "[In RopFastTransferSourceCopyProperties ROP Response Buffer] OutputServerObject: The value of this field MUST be the FastTransfer download context. ");
                }
                else if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyMessage)
                {
                    ModelHelper.CaptureRequirement(487, @"[In RopFastTransferSourceCopyMessages ROP Response Buffer]OutputServerObject: The value of this field MUST be the FastTransfer download context.");
                }
                else if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyFolder)
                {
                    ModelHelper.CaptureRequirement(511, @"[In RopFastTransferSourceCopyFolder ROP Response Buffer]OutputServerObject: The value of this field MUST be the FastTransfer download context.");
                }
                #endregion

                // Get the related folder for the download handle in the folder container.
                foreach (AbstractFolder tempfolder in currentConnection.FolderContainer)
                {
                    if (tempfolder.FolderHandleIndex == currentDownloadContext.RelatedObjectHandleIndex)
                    {
                        currentFolder = tempfolder;
                        currentFolderIndex = currentConnection.FolderContainer.IndexOf(tempfolder);
                        break;
                    }
                }

                // Identify the abstractFastTransferStream type
                switch (currentDownloadContext.AbstractFastTransferStreamType)
                {
                    // The hierarchySync element contains the result of the hierarchy synchronization download operation
                    case FastTransferStreamType.hierarchySync:
                        {
                            if (currentDownloadContext.SynchronizationType == SynchronizationTypes.Hierarchy && priorDownloadOperation == PriorDownloadOperation.RopSynchronizationConfigure)
                            {
                                // Because if the synchronizationType in synchronization configure and the stream type return are Hierarchy indicate this requirement verified.
                                ModelHelper.CaptureRequirement(
                                    3322,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopSynchronizationConfigure, ROP request buffer field conditions is The SynchronizationType field is set to Hierarchy (0x02), Root element in the produced FastTransfer stream is hierarchySync.");
                            }

                            // Create a new HierarchySync.
                            abstractFastTransferStream.AbstractHierarchySync = new AbstractHierarchySync
                            {
                                FolderchangeInfo = new AbstractFolderChange(),
                                AbstractDeletion =
                                {
                                    IdSetDeleted = new Set<int>()
                                },
                                FinalICSState =
                                    new AbstractState
                                    {
                                        AbstractICSStateIndex = AdapterHelper.GetICSStateIndex(),
                                        IsNewCnsetSeenFAIPropertyChangeNumber = false
                                    }
                            };

                            // Assigned a new final ICS State for HierarchySync.

                            // This isn't New ChangeNumber for CnsetSeenFAIProperty in Initialize.

                            // Because of the SynchronizationType must optioned "Hierarchy" value in RopSynchronizationConfigure operation  if FastTransferStreamType is hierarchySync.
                            ModelHelper.CaptureRequirement(1209, "[In state Element] [MetaTagCnsetSeenFAI, Conditional] MetaTagCnsetSeenFAI MUST NOT be present if the SynchronizationType field is set to Hierarchy (0x02), as specified in section 2.2.3.2.1.1.1.");

                            // This isn't New ChangeNumber for CnsetReadProperty in Initialize.
                            abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetReadPropertyChangeNumber = false;
                            ModelHelper.CaptureRequirement(1211, "[In state Element] [MetaTagCnsetRead,Conditional] MetaTagCnsetRead MUST NOT be present if the SynchronizationType field is set to Hierarchy (0x02).");

                            // In case of SynchronizationExtraFlag is EID in current DownloadContext.
                            if (((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid) && (currentFolder.SubFolderIds.Count > 0))
                            {
                                // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                ModelHelper.CaptureRequirement(2730, "[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] Reply is the same whether unknown flags [0x00000010] is set or not.");

                                // In case of SynchronizationExtraFlag is EID in current DownloadContext.
                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid)
                                {
                                    // The PidTagFolderId must be exist  in folder change of HierarchySync.
                                    abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagFolderIdExist = true;

                                    // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                    ModelHelper.CaptureRequirement(1095, "[In folderChange Element] [PidTagFolderId, Conditional] PidTagFolderId MUST be present if and only if the Eid flag of the SynchronizationExtraFlags field is set.");
                                    ModelHelper.CaptureRequirement(
                                                        2191,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagFolderId property (section 2.2.1.2.2) in a folder change header if and only if the Eid flag of the SynchronizationExtraFlags field flag is set.");
                                    ModelHelper.CaptureRequirement(
                                                      2761,
                                                      @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMid property (section 2.2.1.2.1) in a message change header if and only if the Eid flag of the SynchronizationExtraFlags field is set.");

                                    if (currentDownloadContext.SynchronizationType == SynchronizationTypes.Hierarchy)
                                    {
                                        ModelHelper.CaptureRequirement(
                                                       3500,
                                                       @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagFolderId property in the folder change header if the SynchronizationType field is set to Hierarchy (0x02), as specified in section 2.2.3.2.1.1.1.");
                                    }
                                }
                                else
                                {
                                    // The PidTagFolderId must be not exist  in folder change of HierarchySync.
                                    abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagFolderIdExist = false;
                                    ModelHelper.CaptureRequirement(
                                                   716001,
                                                   @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is Eid, the value is 0x00000001] The server does not include the PidTagFolderId (section 2.2.1.2.2) property in the folder change header when the Eid flag of the SynchronizationExtraFlag field is not set.");
                                }
                            }
                            else
                            {
                                // The PidTagFolderId must be not exist  in folder change of HierarchySync.
                                abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagFolderIdExist = false;
                                ModelHelper.CaptureRequirement(
                                                   716001,
                                                   @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is Eid, the value is 0x00000001] The server does not include the PidTagFolderId (section 2.2.1.2.2) property in the folder change header when the Eid flag of the SynchronizationExtraFlag field is not set.");
                            }

                            // In case of SynchronizationExtraFlag is NoForeignIdentifiers in current DownloadContext.
                            if (((currentDownloadContext.Synchronizationflag & SynchronizationFlag.NoForeignIdentifiers) == SynchronizationFlag.NoForeignIdentifiers) && (currentFolder.SubFolderIds.Count > 0))
                            {
                                // The PidTagParentFolderId must be exist  in folder change of HierarchySync.
                                abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagParentFolderIdExist = true;
                                ModelHelper.CaptureRequirement(1097, "[In folderChange Element] [PidTagParentFolderId, Conditional] PidTagParentFolderId MUST be present if the NoForeignIdentifiers flag of the SynchronizationFlags field is set.");

                                // The PidTagParentSourceKey and PidTagSourceKey  must be exist  in folder change of HierarchySync.
                                abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagParentSourceKeyValueZero = false;
                                abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagSourceKeyValueZero = false;

                                ModelHelper.CaptureRequirement(
                                    2178001,
                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoForeignIdentifiers flag of the SynchronizationFlags field is set, server will return null values for the PidTagSourceKey property (section 2.2.1.2.5) and PidTagParentSourceKey (section 2.2.1.2.6) properties when producing the FastTransfer stream for folder and message changes.");
                                ModelHelper.CaptureRequirement(2077, "[In Generating the PidTagSourceKey Value] When requested by the client, the server MUST output the PidTagSourceKey property (section 2.2.1.2.5) value if it is persisted.");
                                ModelHelper.CaptureRequirement(
                                    2179001,
                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoForeignIdentifiers flag of the SynchronizationFlags field is not set, server will return not null values for the PidTagSourceKey and PidTagParentSourceKey properties when producing the FastTransfer stream for folder and message changes.");
                            }
                            else
                            {
                                // The PidTagFolderId must be not exist  in folder change of HierarchySync.
                                abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagParentFolderIdExist = false;
                            }

                            // Sub folder count.
                            int subFolderCount = 0;

                            // Record all descendant folders.
                            Set<int> allDescendantFolders = new Set<int>();
                            if (currentFolder.SubFolderIds.Count > 0)
                            {
                                // Search the current folder in FolderContainer.
                                foreach (AbstractFolder tempFolder in currentConnection.FolderContainer)
                                {
                                    if (currentFolder.FolderIdIndex == tempFolder.ParentFolderIdIndex)
                                    {
                                        // Set the value to the related variable when the current folder is existent.
                                        if (!currentDownloadContext.UpdatedState.IdsetGiven.Contains(tempFolder.FolderIdIndex))
                                        {
                                            // Because of identify that It is a folder by (CurrentFolder.FolderIdIndex == tempFolder.ParentFolderIdIndex) and identify that change number is not in PidTagCnsetSeen by (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(tempFolder.changeNumberIndex)). so can cover requirement here.
                                            ModelHelper.CaptureRequirement(
                                                2042,
                                                @"[In Determining What Differences To Download] For every object in the synchronization scope, servers MUST do the following: 	Include the following syntactical elements in the FastTransfer stream of the OutputServerObject field of the FastTransfer download ROPs, as specified in section 2.2.3.1.1, if one of the following applies:
	Include the folderChange element, as specified in section 2.2.4.3.5, if the object specified by the InputServerObject field of the FastTransfer download ROP request is a Folder object
	And the change number is not included in the value of the MetaTagCnsetSeen property (section 2.2.1.1.2).");

                                            // Add current folder id to IdsetGiven of DownloadContext when the IdsetGiven of updatedState contains the current folder id.
                                            currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Add(tempFolder.FolderIdIndex);

                                            // Assign a new change number for CnsetSeenProperty.
                                            abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;

                                            if (priorUploadOperation == PriorOperation.RopSynchronizationImportHierarchyChange)
                                            {
                                                ModelHelper.CaptureRequirement(
                                                    2235,
                                                    @"[In Receiving a RopSynchronizationImportHierarchyChange Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include a new change number in the MetaTagCnsetSeen property (section 2.2.1.1.2).");
                                            }

                                            subFolderCount++;
                                        }
                                        else
                                        {
                                            if (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(tempFolder.ChangeNumberIndex))
                                            {
                                                // Add current folder id to updatedState.IdsetGiven when the IdsetGiven of updatedState contains the current folder id.
                                                currentDownloadContext.UpdatedState.CnsetSeen = currentDownloadContext.UpdatedState.CnsetSeen.Add(tempFolder.ChangeNumberIndex);
                                                abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;

                                                // Because of identify that It is a folder by (CurrentFolder.FolderIdIndex == tempFolder.ParentFolderIdIndex) and identify that change number is not in PidTagCnsetSeen by (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(tempFolder.changeNumberIndex)). so can cover requirement here.
                                                ModelHelper.CaptureRequirement(
                                                    2042,
                                                    @"[In Determining What Differences To Download] For every object in the synchronization scope, servers MUST do the following: 	Include the following syntactical elements in the FastTransfer stream of the OutputServerObject field of the FastTransfer download ROPs, as specified in section 2.2.3.1.1, if one of the following applies:
	Include the folderChange element, as specified in section 2.2.4.3.5, if the object specified by the InputServerObject field of the FastTransfer download ROP request is a Folder object
	And the change number is not included in the value of the MetaTagCnsetSeen property (section 2.2.1.1.2).");

                                                if (priorUploadOperation == PriorOperation.RopSynchronizationImportHierarchyChange)
                                                {
                                                    ModelHelper.CaptureRequirement(
                                                        2235,
                                                        @"[In Receiving a RopSynchronizationImportHierarchyChange Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include a new change number in the MetaTagCnsetSeen property (section 2.2.1.1.2).");
                                                }

                                                subFolderCount++;
                                            }
                                        }

                                        // In case of current folder's subFolder count greater the 0.
                                        if (tempFolder.SubFolderIds.Count > 0)
                                        {
                                            // Find the second Folder in FolderContainer which was created under current folder.
                                            foreach (AbstractFolder secondFolder in currentConnection.FolderContainer)
                                            {
                                                if (secondFolder.ParentFolderIdIndex == tempFolder.FolderIdIndex)
                                                {
                                                    if (!currentDownloadContext.UpdatedState.IdsetGiven.Contains(secondFolder.FolderIdIndex))
                                                    {
                                                        // Add current folder id to updatedState.IdsetGiven when the IdsetGiven of updatedState contains the current folder id.
                                                        currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Add(secondFolder.FolderIdIndex);

                                                        // Assign a new change number for CnsetSeenPropery.
                                                        abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;

                                                        if (priorUploadOperation == PriorOperation.RopSynchronizationImportHierarchyChange)
                                                        {
                                                            ModelHelper.CaptureRequirement(
                                                                2235,
                                                                @"[In Receiving a RopSynchronizationImportHierarchyChange Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include a new change number in the MetaTagCnsetSeen property (section 2.2.1.1.2).");
                                                        }

                                                        subFolderCount++;
                                                    }
                                                    else
                                                    {
                                                        if (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(secondFolder.ChangeNumberIndex))
                                                        {
                                                            // Add current folder id to updatedState.CnsetSeen when the CnsetSeen of updatedState contains the current folder id.
                                                            currentDownloadContext.UpdatedState.CnsetSeen = currentDownloadContext.UpdatedState.CnsetSeen.Add(secondFolder.ChangeNumberIndex);

                                                            // Assign a new change number for CnsetSeenProperty.
                                                            abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;

                                                            if (priorUploadOperation == PriorOperation.RopSynchronizationImportHierarchyChange)
                                                            {
                                                                ModelHelper.CaptureRequirement(
                                                                    2235,
                                                                    @"[In Receiving a RopSynchronizationImportHierarchyChange Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include a new change number in the MetaTagCnsetSeen property (section 2.2.1.1.2).");
                                                            }

                                                            subFolderCount++;
                                                        }
                                                    }
                                                }

                                                // Add second Folder to descendant folder.
                                                allDescendantFolders = allDescendantFolders.Add(secondFolder.FolderIdIndex);
                                            }
                                        }

                                        // Add second Folder to descendant folder.
                                        allDescendantFolders = allDescendantFolders.Add(tempFolder.FolderIdIndex);
                                    }
                                }
                            }

                            // Search the Descendant folderId in IdsetGiven.
                            foreach (int folderId in currentDownloadContext.UpdatedState.IdsetGiven)
                            {
                                // Identify whether the last Updated Id is or not in Descendant Folder.
                                if (!allDescendantFolders.Contains(folderId))
                                {
                                    if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.NoDeletions) == SynchronizationFlag.NoDeletions)
                                    {
                                        if (requirementContainer.ContainsKey(1062) && requirementContainer[1062])
                                        {
                                            // Deletions isn't present in deletions.
                                            abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IsDeletionPresent = false;
                                            ModelHelper.CaptureRequirement(
                                                1062,
                                                @"[In deletions Element] Implementation does not present deletions element if the NoDeletions flag of the SynchronizationFlag field, as specified in section 2.2.3.2.1.1.1, was set when the synchronization download operation was configured. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                        }

                                        // Deletions isn't present in deletions.
                                        abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IsDeletionPresent = false;
                                        ModelHelper.CaptureRequirement(
                                            2165,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoDeletions flag of the SynchronizationFlags field is set, the server MUST NOT download information about item deletions, as specified in section 2.2.4.3.3");
                                    }
                                    else
                                    {
                                        // Deletions is present in deletions.
                                        abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IsDeletionPresent = true;

                                        // PidTagIdsetNoLongerInScope be present in deletions
                                        abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IsPidTagIdsetNoLongerInScopeExist = false;
                                        ModelHelper.CaptureRequirement(
                                                    1333,
                                                    @"[In deletions Element] [The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:] MUST adhere to the following restrictions: MetaTagIdsetNoLongerInScope MUST NOT be present if the Hierarchy value of the SynchronizationType field is set, as specified in section 2.2.3.2.1.1.1.");

                                        // Are folders that have never been reported as deleted
                                        abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IsPidTagIdsetExpiredExist = false;
                                        ModelHelper.CaptureRequirement(
                                                    2046,
                                                    @"[In Determining What Differences To Download] [For every object in the synchronization scope, servers MUST do the following:] If the NoDeletions flag of the SynchronizationFlags field is not set, include the deletions element, as specified in section 2.2.4.3.3, for objects that either: 
Have their internal identifiers present in the value of the MetaTagIdsetGiven property (section 2.2.1.1.1) and are missing from the server replica.Are folders that have never been reported as deleted folders.
Are folders that have never been reported as deleted folders.");

                                        ModelHelper.CaptureRequirement(
                                            1337002,
                                            @"[In deletions Element] [The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:] MUST adhere to the following restrictions: ] MetaTagIdsetExpired (section 2.2.1.3.3) MUST NOT be present if the Hierarchy value of the SynchronizationType field is set. ");

                                        // Because isPidTagIdsetExpiredExist value is false so have their internal identifiers present in PidTagIdsetGiven and isDeletionPresent is true to those are missing from the server replica. So cover requirement here.
                                        ModelHelper.CaptureRequirement(
                                            2045,
                                            @"[In Determining What Differences To Download] [For every object in the synchronization scope, servers MUST do the following:] If the NoDeletions flag of the SynchronizationFlag field is not set, include the deletions element, as specified in section 2.2.4.3.3, for objects that either:
	                                         Have their internal identifiers present in the value of the MetaTagIdsetGiven property (section 2.2.1.1.1) and are missing from the server replica.Are folders that have never been reported as deleted folders.
                                            Are folders that have never been reported as deleted folders.");

                                        // Identify the current ExchangeServer Version.
                                        if (requirementContainer.ContainsKey(2652) && requirementContainer[2652])
                                        {
                                            // Add folderId to IdSetDeleted of abstractDeletion
                                            abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IdSetDeleted = abstractFastTransferStream.AbstractHierarchySync.AbstractDeletion.IdSetDeleted.Add(folderId);

                                            // Because of execute RopSynchronizationImportHierarchyChange or RopSynchronizationImportMessageChange operation before execute RopSynchronizationImportDeletes operation, record deletions and  prevent it restoring them back finished in MS_OXCFXICSAdapter. So cover this requirement here.
                                            ModelHelper.CaptureRequirement(
                                                2652,
                                                @"[In Receiving a RopSynchronizationImportDeletes Request] Implementation does record deletions of objects that never existed in the server replica, in order to prevent the RopSynchronizationImportHierarchyChange (section 2.2.3.2.4.3) or RopSynchronizationImportMessageChange (section 2.2.3.2.4.2) ROPs from restoring them back. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                        }

                                        // Delete folder id from DownloadContext.
                                        currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Remove(folderId);

                                        // If NoDeletions flag is not set and the operation can be executed successfully, the element isDeletionPresent as true will be returned, which means the deletion elements is downloaded, so this requirement can be verified.
                                        ModelHelper.CaptureRequirement(
                                            2167,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoDeletions flag of the SynchronizationFlags field is not set, the server MUST download information about item deletions, as specified in section 2.2.4.3.3.");
                                    }

                                    // This is a new changeNumber
                                    abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;
                                    if (priorUploadOperation == PriorOperation.RopSynchronizationImportHierarchyChange)
                                    {
                                        ModelHelper.CaptureRequirement(
                                            2235,
                                            @"[In Receiving a RopSynchronizationImportHierarchyChange Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include a new change number in the MetaTagCnsetSeen property (section 2.2.1.1.2).");
                                    }
                                }
                            }

                            // The lDescendantFolder is existent.
                            if (currentFolder.SubFolderIds == allDescendantFolders)
                            {
                                if (allDescendantFolders.Count > 0)
                                {
                                    // Parent folder is existent and PidTagParentSourceKey is not in folder change
                                    abstractFastTransferStream.AbstractHierarchySync.FolderchangeInfo.IsPidTagParentSourceKeyValueZero = true;
                                }
                            }
                            else
                            {
                                ModelHelper.CaptureRequirement(1129, "[In hierarchySync Element]The folderChange elements for the parent folders MUST be output before any of their child folders.");
                                abstractFastTransferStream.AbstractHierarchySync.IsParentFolderBeforeChild = true;
                            }

                            // Set value for finalICSState.
                            abstractFastTransferStream.AbstractHierarchySync.FinalICSState.IdSetGiven = currentDownloadContext.UpdatedState.IdsetGiven;
                            abstractFastTransferStream.AbstractHierarchySync.FolderCount = subFolderCount;
                            currentFolder.ICSStateContainer.Add(abstractFastTransferStream.AbstractHierarchySync.FinalICSState.AbstractICSStateIndex, currentDownloadContext.UpdatedState);

                            // Because of the "foreach" Search the Descendant folderId in IdsetGive  and the end of "foreach" search. So cover this requirement here.
                            ModelHelper.CaptureRequirement(1128, "[In hierarchySync Element]There MUST be exactly one folderChange element for each descendant folder of the root of the synchronization operation (that is the folder that was passed to the RopSynchronizationConfigure ROP, as specified in section 2.2.3.2.1.1) that is new or has been changed since the last synchronization.");

                            // Update the FolderContainer.
                            currentConnection.FolderContainer = currentConnection.FolderContainer.Update(currentFolderIndex, currentFolder);
                            connections[serverId] = currentConnection;

                            break;
                        }

                    // The state element contains the final ICS state of the synchronization download operation.
                    case FastTransferStreamType.state:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopSynchronizationGetTransferState)
                            {
                                // Because if the isRopSynchronizationGetTransferState called the steam type return by RopFastTransferSourceGetBuffer should be State
                                ModelHelper.CaptureRequirement(
                                    3323,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopSynchronizationGetTransferState, ROP request buffer field conditions is always, Root element in the produced FastTransfer stream is state.");
                            }

                            // Create a new abstractState.
                            abstractFastTransferStream.AbstractState = new AbstractState
                            {
                                AbstractICSStateIndex = AdapterHelper.GetICSStateIndex()
                            };

                            // Assign a new ICS State Index.

                            // The new IdSetGiven of State value equal to IdsetGiven of current download context.
                            if (requirementContainer.ContainsKey(350400101) && requirementContainer[350400101] && (priorOperation == PriorOperation.RopSynchronizationOpenCollector))
                            {
                                abstractFastTransferStream.AbstractState.IdSetGiven = new Set<int>();
                                ModelHelper.CaptureRequirement(
                                    350400101,
                                    @"[In Appendix A: Product Behavior] Implementation does use this behavior. <40> Section 3.2.5.9.3.1: In Exchange 2007, the RopSynchronizationGetTransferState ROP (section 2.2.3.2.3.1) returns a checkpoint ICS state that is reflective of the current status.");
                            }
                            else
                            {
                                abstractFastTransferStream.AbstractState.IdSetGiven = currentDownloadContext.UpdatedState.IdsetGiven;
                            }

                            // Add the new stat to ICSStateContainer.
                            currentFolder.ICSStateContainer.Add(abstractFastTransferStream.AbstractState.AbstractICSStateIndex, currentDownloadContext.UpdatedState);
                            currentConnection.FolderContainer = currentConnection.FolderContainer.Update(currentFolderIndex, currentFolder);
                            connections[serverId] = currentConnection;
                            break;
                        }

                    // The messageContent element represents the content of a message.
                    case FastTransferStreamType.contentsSync:
                        {
                            if (currentDownloadContext.SynchronizationType == SynchronizationTypes.Contents && priorDownloadOperation == PriorDownloadOperation.RopSynchronizationConfigure)
                            {
                                // Because if the synchronizationType in synchronization configure and the stream type return are contents indicate this requirement verified.
                                ModelHelper.CaptureRequirement(
                                    3321,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopSynchronizationConfigure, ROP request buffer field conditions is The SynchronizationType field is set to Contents (0x01), Root element in the produced FastTransfer stream is contentsSync.");
                            }

                            // Create a new abstractContentsSync.
                            abstractFastTransferStream.AbstractContentsSync = new AbstractContentsSync
                            {
                                MessageInfo = new Set<AbstractMessageChangeInfo>(),
                                AbstractDeletion =
                                {
                                    IdSetDeleted = new Set<int>()
                                },
                                FinalICSState = new AbstractState
                                {
                                    // Assign abstractICSStateIndex.
                                    AbstractICSStateIndex = AdapterHelper.GetICSStateIndex()
                                }
                            };

                            // Create a new finalICSState.
                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Progress) == SynchronizationFlag.Progress)
                            {
                                // The current ContentsSync is include progressTotal.
                                abstractFastTransferStream.AbstractContentsSync.IsprogessTotalPresent = true;
                                ModelHelper.CaptureRequirement(1179, "[In progressTotal Element]This element MUST be present if the Progress flag of the SynchronizationFlags field, as specified in section 2.2.3.2.1.1.1, was set when configuring the synchronization download operation and a server supports progress reporting.");

                                if (requirementContainer.ContainsKey(2675) && requirementContainer[2675])
                                {
                                    // The current ContentsSync is include progressTotal.
                                    abstractFastTransferStream.AbstractContentsSync.IsprogessTotalPresent = true;

                                    // The current server is exchange.
                                    ModelHelper.CaptureRequirement(
                                        2675,
                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] Implementation does inject the progressTotal element, as specified in section 2.2.4.3.19, into the FastTransfer stream, if the Progress flag of the SynchronizationFlag field is set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }
                            }
                            else
                            {
                                // The current ContentsSync is not include progressTotal.
                                abstractFastTransferStream.AbstractContentsSync.IsprogessTotalPresent = false;
                                ModelHelper.CaptureRequirement(
                                    2188,
                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the Progress flag of the SynchronizationFlags field is not set, the server MUST not inject the progressTotal element into the FastTransfer stream.");
                                ModelHelper.CaptureRequirement(1180, "[In progressTotal Element]This element MUST NOT be present if the Progress flag of the SynchronizationFlags field was not set when configuring the synchronization download operation.");
                            }

                            // Search the message ids in IdsetGiven.
                            foreach (int messageId in currentDownloadContext.UpdatedState.IdsetGiven)
                            {
                                if (!currentFolder.MessageIds.Contains(messageId))
                                {
                                    // Assign a new change number for CnsetRead Property.
                                    if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.NoDeletions) == SynchronizationFlag.NoDeletions)
                                    {
                                        // The server MUST NOT download information about item deletions.
                                        abstractFastTransferStream.AbstractContentsSync.AbstractDeletion.IsDeletionPresent = false;
                                        ModelHelper.CaptureRequirement(
                                            2165,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoDeletions flag of the SynchronizationFlags field is set, the server MUST NOT download information about item deletions, as specified in section 2.2.4.3.3");
                                        ModelHelper.CaptureRequirement(
                                            2166,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoDeletions flag of the SynchronizationFlags field is set, the server MUST respond as if the IgnoreNoLongerInScope flag was set.");
                                    }
                                    else if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.IgnoreNoLongerInScope) != SynchronizationFlag.IgnoreNoLongerInScope)
                                    {
                                        // The server MUST  download information about item deletions.
                                        abstractFastTransferStream.AbstractContentsSync.AbstractDeletion.IsDeletionPresent = true;
                                        ModelHelper.CaptureRequirement(
                                            2167,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the NoDeletions flag of the SynchronizationFlags field is not set, the server MUST download information about item deletions, as specified in section 2.2.4.3.3.");

                                        // Identify the current ExchangeServer Version.
                                        if (requirementContainer.ContainsKey(2652) && requirementContainer[2652])
                                        {
                                            // Add messageId to abstractDeletion of abstractDeletion.
                                            abstractFastTransferStream.AbstractContentsSync.AbstractDeletion.IdSetDeleted = abstractFastTransferStream.AbstractContentsSync.AbstractDeletion.IdSetDeleted.Add(messageId);

                                            // Because of execute RopSynchronizationImportHierarchyChange or RopSynchronizationImportMessageChange operation before execute RopSynchronizationImportDeletes operation, record deletions and  prevent it restoring them back finished in MS_OXCFXICSAdapter. So cover this requirement here.
                                            ModelHelper.CaptureRequirement(
                                                2652,
                                                @"[In Receiving a RopSynchronizationImportDeletes Request] Implementation does record deletions of objects that never existed in the server replica, in order to prevent the RopSynchronizationImportHierarchyChange (section 2.2.3.2.4.3) or RopSynchronizationImportMessageChange (section 2.2.3.2.4.2) ROPs from restoring them back. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                        }

                                        currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Remove(messageId);
                                        ModelHelper.CaptureRequirement(
                                            2169,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the IgnoreNoLongerInScope flag of the SynchronizationFlags field is not set, the server MUST download information about messages that went out of scope as deletions, as specified in section 2.2.4.3.3.");
                                    }

                                    if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.IgnoreNoLongerInScope) == SynchronizationFlag.IgnoreNoLongerInScope)
                                    {
                                        // PidTagIdsetNoLongerInScope MUST NOT be present.
                                        abstractFastTransferStream.AbstractContentsSync.AbstractDeletion.IsPidTagIdsetNoLongerInScopeExist = false;
                                        ModelHelper.CaptureRequirement(1334, @"[In deletions Element] [The following restrictions exist on the contained propList element, as specified in section 2.2.4.3.20:] MUST adhere to the following restrictions: MetaTagIdsetNoLongerInScope MUST NOT be present if the IgnoreNoLongerInScope flag of the SynchronizationFlags field is set.");
                                        ModelHelper.CaptureRequirement(
                                            2168,
                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the IgnoreNoLongerInScope flag of the SynchronizationFlags field is set, the server MUST NOT download information about messages that went out of scope as deletions, as specified in section 2.2.4.3.3.");
                                    }
                                }
                            }
                            #region Set message change contents

                            // Search the current message which match search condition in MessageContainer
                            foreach (AbstractMessage tempMessage in currentConnection.MessageContainer)
                            {
                                if (tempMessage.FolderIdIndex == currentFolder.FolderIdIndex)
                                {
                                    // The found message is FAI message.
                                    if (tempMessage.IsFAImessage)
                                    {
                                        if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.FAI) == SynchronizationFlag.FAI)
                                        {
                                            if (!currentDownloadContext.UpdatedState.IdsetGiven.Contains(tempMessage.MessageIdIndex))
                                            {
                                                // Create a new messageChange of contentSync.
                                                AbstractMessageChangeInfo newMessageChange = new AbstractMessageChangeInfo
                                                {
                                                    MessageIdIndex = tempMessage.MessageIdIndex,
                                                    IsMessageChangeFull = true
                                                };

                                                // Set the value for new messageChange.

                                                // This MessageChangeFull is in  new message change.
                                                if ((currentDownloadContext.Sendoptions & SendOptionAlls.PartialItem) != SendOptionAlls.PartialItem)
                                                {
                                                    ModelHelper.CaptureRequirement(
                                                        1135,
                                                        @"[In messageChange Element] A server MUST use the messageChangeFull element, instead of the messageChangePartial element, if any of the following are true: 	The PartialItem flag of the SendOptions field was not set, as specified in section 2.2.3.2.1.1.");
                                                }

                                                // Because SynchronizationFlag is set FAI and isMessageChangeFull is true  so can cover requirement here.
                                                ModelHelper.CaptureRequirement(1137, @"[In messageChange Element] [A server MUST use the messageChangeFull element, instead of the messageChangePartial element, if any of the following are true:] The message is an FAI message.");
                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.CN) == SynchronizationExtraFlag.CN)
                                                {
                                                    // PidTagChangeNumber must be present .
                                                    newMessageChange.IsPidTagChangeNumberExist = true;

                                                    // Because the SynchronizationExtraFlag is CN and the SynchronizationExtraFlag only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(1367, "[In messageChangeHeader Element, PidTagChangeNumber,Conditional]PidTagChangeNumber MUST be present if and only if the CN flag of the SynchronizationExtraFlags field is set.");

                                                    // Because the SynchronizationExtraFlag is CN and the SynchronizationExtraFlag only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2196,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagChangeNumber property (section 2.2.1.2.3) in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // PidTagChangeNumber must not present .
                                                    newMessageChange.IsPidTagChangeNumberExist = false;

                                                    // Because the SynchronizationExtraFlag is not CN and the SynchronizationExtraFlag only don't set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2197,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST NOT include the PidTagChangeNumber property in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid)
                                                {
                                                    // The server MUST include the PidTagMid property in the messageChange.
                                                    newMessageChange.IsPidTagMidExist = true;

                                                    // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set EID. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(1363, "[In messageChangeHeader Element] [PidTagMid,Conditional] PidTagMid MUST be present if and only if the Eid flag of the SynchronizationExtraFlags field is set, as specified in section 2.2.3.2.1.1.1.");
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagFolderId property in the messageChange. 
                                                    newMessageChange.IsPidTagMidExist = false;
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.MessageSize) == SynchronizationExtraFlag.MessageSize)
                                                {
                                                    // The server include the PidTagMessageSize property in the messageChange. 
                                                    newMessageChange.IsPidTagMessageSizeExist = true;

                                                    // Because the SynchronizationExtraFlag is MessageSize and the SynchronizationExtraFlag only set MessageSize. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2195,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMessageSize property (section 2.2.1.6) in the message change header if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");
                                                    ModelHelper.CaptureRequirement(
                                                        1365,
                                                        @"[In messageChangeHeader Element] [PidTagMessageSize,Conditional] PidTagMessageSize MUST be present if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");
                                                }

                                                if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Progress) == SynchronizationFlag.Progress)
                                                {
                                                    // The server must include the ProgressPerMessage in the messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = true;

                                                    // Message object that follows is FAI
                                                    newMessageChange.FollowedFAIMessage = true;

                                                    // Because the SynchronizationExtraFlag is Progress and the SynchronizationExtraFlag only set Progress. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1171,
                                                        @"[In progressPerMessage Element]MUST be present if and only if the progessTotal element, as specified in
                                                        section 2.2.4.3.18, was output within the same ancestor contentsSync element, as specified in section
                                                        2.2.4.3.2.");
                                                    ModelHelper.CaptureRequirement(
                                                        1382,
                                                        @"[In progressPerMessage Element] [[PtypBoolean] 0x0000000B] [The server returns] TRUE (0x01 or any non-zero value) if the Message object that follows is FAI.");
                                                }
                                                else
                                                {
                                                    // The server don't include the ProgressPerMessage in the messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = false;
                                                    ModelHelper.CaptureRequirement(
                                                        1172,
                                                        @"[In progressPerMessage Element] [ProgressPerMessage Element] MUST NOT be present if the Progress flag of the SynchronizationFlags field was not set when configuring the synchronization download operation.");
                                                }

                                                if (tempMessage.MessageProperties.Contains("PidTagBody"))
                                                {
                                                    if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.BestBody) != SynchronizationFlag.BestBody)
                                                    {
                                                        // Implementation does only support the message body (2) which is always in the original format
                                                        if (requirementContainer.ContainsKey(3118002) && requirementContainer[3118002])
                                                        {
                                                            newMessageChange.IsRTFformat = false;
                                                            ModelHelper.CaptureRequirement(
                                                                3118002,
                                                                @"[In Appendix A: Product Behavior] Implementation does only support the message body which is always in the original format. <3> Section 2.2.3.1.1.1.1: In Exchange 2013 and Exchange 2016, the message body is always in the original format.");
                                                        }

                                                        if (requirementContainer.ContainsKey(2117002) && requirementContainer[2117002])
                                                        {
                                                            // Identify whether message bodies in the compressed RTF format is or not.
                                                            newMessageChange.IsRTFformat = true;

                                                            // Because the BestBody flag of the CopyFlags field is not set before and the prior ROP is RopFastTransferSourceCopyMessages, so this requirement can be captured.
                                                            ModelHelper.CaptureRequirement(
                                                                2117002,
                                                                @"[In Appendix A: Product Behavior] <33> Section 3.2.5.8.1.3: Implementation does support the BestBody flag. If the BestBody flag of the CopyFlags field is not set, implementation does output message bodies in the compressed RTF (Microsoft Exchange Server 2007 and Exchange Server 2010 follow this behavior.)");
                                                        }

                                                        if (requirementContainer.ContainsKey(3118003) && requirementContainer[3118003])
                                                        {
                                                            // Identify whether message bodies in the compressed RTF format is or not.
                                                            newMessageChange.IsRTFformat = true;

                                                            // Because the BestBody flag of the CopyFlags field is not set before and the prior ROP is RopFastTransferSourceCopyMessages, so this requirement can be captured.
                                                            ModelHelper.CaptureRequirement(
                                                                3118003,
                                                                @"[In Appendix A: Product Behavior] Implementation does support this flag [BestBody flag] [in RopFastTransferSourceCopyTo ROP]. (<3> Section 2.2.3.1.1.1.1: Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
                                                        }

                                                        if (requirementContainer.ContainsKey(499001) && requirementContainer[499001])
                                                        {
                                                            // Identify whether message bodies in the compressed RTF format is or not.
                                                            newMessageChange.IsRTFformat = false;

                                                            // Because the BestBody flag of the CopyFlags field is set before and the prior ROP is RopFastTransferSourceCopyMessages, so this requirement can be captured.
                                                            ModelHelper.CaptureRequirement(
                                                                   499001,
                                                                   @"[In Appendix A: Product Behavior] Implementation does only support the message body which is always in the original format. <5> Section 2.2.3.1.1.3.1: In Exchange 2013 and Exchange 2016, the message body is always in the original format.");
                                                        }

                                                        if (requirementContainer.ContainsKey(2182002) && requirementContainer[2182002])
                                                        {
                                                            // Identify whether message bodies in the compressed RTF format is or not.
                                                            newMessageChange.IsRTFformat = true;

                                                            // Because the BestBody flag of the CopyFlags field is set before and the prior ROP is RopFastTransferSourceCopyMessages, so this requirement can be captured.
                                                            ModelHelper.CaptureRequirement(
                                                                   2182002,
                                                                   @"[In Appendix A: Product Behavior] Implementation does support BestBody flag [in RopSynchronizationConfigure ROP]. (<38> Section 3.2.5.9.1.1: Microsoft Exchange Server 2007 and Exchange Server 2010 follow this behavior.)");
                                                        }
                                                    }
                                                }

                                                // Add new messagChange to ContentsSync.
                                                abstractFastTransferStream.AbstractContentsSync.MessageInfo = abstractFastTransferStream.AbstractContentsSync.MessageInfo.Add(newMessageChange);

                                                // Add message id to IdsetGiven of current download context.
                                                currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Add(tempMessage.MessageIdIndex);
                                                ModelHelper.CaptureRequirement(
                                                    2172,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the FAI flag of the SynchronizationFlags field is set, the server MUST download information about changes to FAI messages, as specified by the folderContents element in section 2.2.4.3.7.");
                                            }
                                            else if (!currentDownloadContext.UpdatedState.CnsetSeenFAI.Contains(tempMessage.ChangeNumberIndex))
                                            {
                                                // The message change number is not in CnsetSeenFAI property.
                                                // Create a new  message change.
                                                AbstractMessageChangeInfo newMessageChange = new AbstractMessageChangeInfo
                                                {
                                                    MessageIdIndex = tempMessage.MessageIdIndex,
                                                    IsMessageChangeFull = true
                                                };

                                                // Set message id for new messageChange

                                                // The server don't include messagchangeFull in messageChange.
                                                ModelHelper.CaptureRequirement(
                                                    1137,
                                                    @"[In messageChange Element] [A server MUST use the messageChangeFull element, instead of the messageChangePartial element, if any of the following are true:] The message is an FAI message.");
                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.CN) == SynchronizationExtraFlag.CN)
                                                {
                                                    // The server MUST include the PidTagChangeNumber property in the message change
                                                    newMessageChange.IsPidTagChangeNumberExist = true;

                                                    // Because the SynchronizationExtraFlag is CN and the SynchronizationExtraFlag only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2196,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagChangeNumber property (section 2.2.1.2.3) in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagChangeNumber property in the message change
                                                    newMessageChange.IsPidTagChangeNumberExist = false;

                                                    // Because the SynchronizationExtraFlag is not CN and the SynchronizationExtraFlag  not only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2197,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST NOT include the PidTagChangeNumber property in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid)
                                                {
                                                    // The server MUST include the PidTagFolderId property in the messageChange.
                                                    newMessageChange.IsPidTagMidExist = true;

                                                    // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2191,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagFolderId property (section 2.2.1.2.2) in a folder change header if and only if the Eid flag of the SynchronizationExtraFlags field flag is set.");

                                                    ModelHelper.CaptureRequirement(
                                                        2761,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMid property (section 2.2.1.2.1) in a message change header if and only if the Eid flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // The server MUST NOT include the PidTagMid property in the messageChange.
                                                    newMessageChange.IsPidTagMidExist = false;

                                                    // Because the SynchronizationExtraFlag is Eid and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        716002,
                                                        @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is Eid, the value is 0x00000001] The server does not include the PidTagMid (section 2.2.1.2.1) property in the message change header when the Eid flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.MessageSize) == SynchronizationExtraFlag.MessageSize)
                                                {
                                                    ModelHelper.CaptureRequirement(
                                                        2195,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMessageSize property (section 2.2.1.6) in the message change header if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");

                                                    // The server MUST include the PidTagMessageSize property in the message change.
                                                    newMessageChange.IsPidTagMessageSizeExist = true;

                                                    // Because the SynchronizationExtraFlag is MessageSize and the SynchronizationExtraFlag only set MessageSize. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1365,
                                                        @"[In messageChangeHeader Element] [PidTagMessageSize,Conditional] PidTagMessageSize MUST be present if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    newMessageChange.IsPidTagMessageSizeExist = false;

                                                    // When the MessageSize flag of the SynchronizationExtraFlag field is not set and the server does not include the PidTagMessageSize property (section 2.2.1.6) in the message change header this requirement captured.
                                                    ModelHelper.CaptureRequirement(
                                                        718001,
                                                        @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is MessageSize, the value is 0x00000002] The server does not include the PidTagMessageSize property (section 2.2.1.6) in the message change header if the MessageSize flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Progress) == SynchronizationFlag.Progress)
                                                {
                                                    // The server include ProgressPerMessag in messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = true;

                                                    // The message is a FAl Message.
                                                    newMessageChange.FollowedFAIMessage = true;

                                                    // Because the SynchronizationExtraFlag is Progress and the SynchronizationExtraFlag only set Progress. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1171,
                                                        @"[In progressPerMessage Element]MUST be present if and only if the progessTotal element, as specified in
                                                        section 2.2.4.3.18, was output within the same ancestor contentsSync element, as specified in section
                                                        2.2.4.3.2.");
                                                    ModelHelper.CaptureRequirement(1382, "[In progressPerMessage Element] [[PtypBoolean] 0x0000000B] [The server returns] TRUE (0x01 or any non-zero value) if the Message object that follows is FAI.");
                                                }
                                                else
                                                {
                                                    // The server don't include ProgressPerMessag in messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = false;
                                                    ModelHelper.CaptureRequirement(
                                                        1172,
                                                        @"[In progressPerMessage Element] [ProgressPerMessage Element] MUST NOT be present if the Progress flag of the SynchronizationFlags field was not set when configuring the synchronization download operation.");
                                                }

                                                if (!messagechangePartail && (requirementContainer.ContainsKey(2172) && !requirementContainer[2172]))
                                                {
                                                    abstractFastTransferStream.AbstractContentsSync.MessageInfo = new Set<AbstractMessageChangeInfo>();
                                                }
                                                else
                                                {
                                                    // Add new messagChange to ContentsSync.
                                                    abstractFastTransferStream.AbstractContentsSync.MessageInfo = abstractFastTransferStream.AbstractContentsSync.MessageInfo.Add(newMessageChange);
                                                    ModelHelper.CaptureRequirement(
                                                    2172,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the FAI flag of the SynchronizationFlags field is set, the server MUST download information about changes to FAI messages, as specified by the folderContents element in section 2.2.4.3.7.");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // Because SynchronizationFlag FAI flag is not set and no have download context add to abstractFastTransferStream. So can cover this requirement here.
                                            ModelHelper.CaptureRequirement(
                                                2173,
                                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the FAI flag of the SynchronizationFlags field is not set, the server MUST NOT download information about changes to FAI messages, as specified by the folderContents element in section 2.2.4.3.7.");
                                        }
                                    }
                                    else
                                    {
                                        // The found message is Normal message.
                                        // SynchronizationFlag is Normal in download context.
                                        if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Normal) == SynchronizationFlag.Normal)
                                        {
                                            // The message id not in IdsetGiven of download context.
                                            if (!currentDownloadContext.UpdatedState.IdsetGiven.Contains(tempMessage.MessageIdIndex))
                                            {
                                                // Create a newMessageChange.
                                                AbstractMessageChangeInfo newMessageChange = new AbstractMessageChangeInfo
                                                {
                                                    MessageIdIndex = tempMessage.MessageIdIndex,
                                                    IsMessageChangeFull = true
                                                };

                                                // Set message id for newMessageChange.

                                                // This messagechangeFull is in newMessageChange.

                                                // Because the  message object include MId  and it is initial ICS state through set sequence.
                                                ModelHelper.CaptureRequirement(1136, "[In messageChange Element] [A server MUST use the messageChangeFull element, instead of the messageChangePartial element, if any of the following are true:] The MID structure ([MS-OXCDATA] section 2.2.1.2) of the message to be output is not in the MetaTagIdsetGiven property (section 2.2.1.1.1) from the initial ICS state.");
                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.CN) == SynchronizationExtraFlag.CN)
                                                {
                                                    // The server MUST include the PidTagChangeNumber property in the messageChange.
                                                    newMessageChange.IsPidTagChangeNumberExist = true;

                                                    // Because the SynchronizationExtraFlag is CN and the SynchronizationExtraFlag only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2196,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagChangeNumber property (section 2.2.1.2.3) in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagChangeNumber property in the messageChange.
                                                    newMessageChange.IsPidTagChangeNumberExist = false;

                                                    // Because the SynchronizationExtraFlag is not CN and the SynchronizationExtraFlag not only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2197,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST NOT include the PidTagChangeNumber property in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid)
                                                {
                                                    // The server MUST include the PidTagFolderId property in the messageChange.
                                                    newMessageChange.IsPidTagMidExist = true;

                                                    // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2191,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagFolderId property (section 2.2.1.2.2) in a folder change header if and only if the Eid flag of the SynchronizationExtraFlags field flag is set.");

                                                    ModelHelper.CaptureRequirement(
                                                      2761,
                                                      @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMid property (section 2.2.1.2.1) in a message change header if and only if the Eid flag of the SynchronizationExtraFlags field is set.");

                                                    if (currentDownloadContext.SynchronizationType == SynchronizationTypes.Contents)
                                                    {
                                                        ModelHelper.CaptureRequirement(
                                                            3501,
                                                            @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMid property in the message change header if the SynchronizationType field is set Contents (0x01), as specified in section 2.2.3.2.1.1.1.");
                                                    }
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagFolderId property in the messageChange.
                                                    newMessageChange.IsPidTagMidExist = false;

                                                    // Because the SynchronizationExtraFlag is Eid and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        716002,
                                                        @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is Eid, the value is 0x00000001] The server does not include the PidTagMid (section 2.2.1.2.1) property in the message change header when the Eid flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.MessageSize) == SynchronizationExtraFlag.MessageSize)
                                                {
                                                    ModelHelper.CaptureRequirement(
                                                        2195,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMessageSize property (section 2.2.1.6) in the message change header if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");

                                                    // The server MUST include the PidTagMessageSize property in the messageChange.
                                                    newMessageChange.IsPidTagMessageSizeExist = true;

                                                    // Because the SynchronizationExtraFlag is MessageSize and the SynchronizationExtraFlag only set MessageSize. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1365,
                                                        @"[In messageChangeHeader Element] [PidTagMessageSize,Conditional] PidTagMessageSize MUST be present if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    newMessageChange.IsPidTagMessageSizeExist = false;
                                                    ModelHelper.CaptureRequirement(
                                                        718001,
                                                        @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is MessageSize, the value is 0x00000002] The server does not include the PidTagMessageSize property (section 2.2.1.6) in the message change header if the MessageSize flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Progress) == SynchronizationFlag.Progress)
                                                {
                                                    // The server include the progessTotal in the messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = true;

                                                    // The message object is normal message.
                                                    newMessageChange.FollowedFAIMessage = false;

                                                    // Because the SynchronizationExtraFlag is Progress and the SynchronizationExtraFlag only set Progress. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1171,
                                                        @"[In progressPerMessage Element]MUST be present if and only if the progessTotal element, as specified in
                                                        section 2.2.4.3.18, was output within the same ancestor contentsSync element, as specified in section
                                                        2.2.4.3.2.");
                                                    ModelHelper.CaptureRequirement(
                                                        1383,
                                                        @"[In progressPerMessage Element] [[PtypBoolean] 0x0000000B] otherwise[if the Message object that follows is not FAI] ,[the server returns] FALSE (0x00).");
                                                }
                                                else
                                                {
                                                    // The server don't include the progessTotal in the messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = false;
                                                    ModelHelper.CaptureRequirement(
                                                        1172,
                                                        @"[In progressPerMessage Element] [ProgressPerMessage Element] MUST NOT be present if the Progress flag of the SynchronizationFlags field was not set when configuring the synchronization download operation.");
                                                }

                                                // Add new messageChange to ContentsSync.
                                                abstractFastTransferStream.AbstractContentsSync.MessageInfo = abstractFastTransferStream.AbstractContentsSync.MessageInfo.Add(newMessageChange);

                                                // Add messageId to IdsetGiven of download context.
                                                currentDownloadContext.UpdatedState.IdsetGiven = currentDownloadContext.UpdatedState.IdsetGiven.Add(tempMessage.MessageIdIndex);

                                                // Because SynchronizationFlag is Normal
                                                ModelHelper.CaptureRequirement(
                                                    2174,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the Normal flag of the SynchronizationFlags flag is set, the server MUST download information about changes to normal messages, as specified in section 2.2.4.3.11.");
                                            }
                                            else if (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(tempMessage.ChangeNumberIndex))
                                            {
                                                // The message change number is not in CnsetSeenFAI property.
                                                // Create a newMessageChange.
                                                AbstractMessageChangeInfo newMessageChange = new AbstractMessageChangeInfo
                                                {
                                                    // Set messageId for newMessageChange.
                                                    MessageIdIndex = tempMessage.MessageIdIndex
                                                };

                                                if ((currentDownloadContext.Sendoptions & SendOptionAlls.PartialItem) != SendOptionAlls.PartialItem)
                                                {
                                                    // The server include MessageChangeFull in messageChange.
                                                    newMessageChange.IsMessageChangeFull = true;
                                                    ModelHelper.CaptureRequirement(
                                                        1135,
                                                        @"[In messageChange Element]A server MUST use the messageChangeFull element, instead of the
                                                        messageChangePartial element, if any of the following are true:The PartialItem flag of the SendOptions
                                                        field was not set, as specified in section 2.2.3.2.1.1.");
                                                }
                                                else
                                                {
                                                    messagechangePartail = false;

                                                    // The server include MessageChangePartial in messageChange.
                                                    newMessageChange.IsMessageChangeFull = false;
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.CN) == SynchronizationExtraFlag.CN)
                                                {
                                                    // The server include the PidTagChangeNumber property in the messageChange.
                                                    newMessageChange.IsPidTagChangeNumberExist = true;

                                                    // Because the SynchronizationExtraFlag is CN and the SynchronizationExtraFlag only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2196,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagChangeNumber property (section 2.2.1.2.3) in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagChangeNumber property in the messageChange.
                                                    newMessageChange.IsPidTagChangeNumberExist = false;

                                                    // Because the SynchronizationExtraFlag is not  CN and the SynchronizationExtraFlag not only set CN. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2197,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST NOT include the PidTagChangeNumber property in the message change header if and only if the CN flag of the SynchronizationExtraFlags field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.Eid) == SynchronizationExtraFlag.Eid)
                                                {
                                                    // The server include the PidTagMid property in messageChange.
                                                    newMessageChange.IsPidTagMidExist = true;

                                                    // Because the SynchronizationExtraFlag is EID and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2191,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagFolderId property (section 2.2.1.2.2) in a folder change header if and only if the Eid flag of the SynchronizationExtraFlags field flag is set.");

                                                    ModelHelper.CaptureRequirement(
                                                      2761,
                                                      @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMid property (section 2.2.1.2.1) in a message change header if and only if the Eid flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    // The server don't include the PidTagMid property in messageChange.
                                                    newMessageChange.IsPidTagMidExist = false;

                                                    // Because the SynchronizationExtraFlag is Eid and the SynchronizationExtraFlag only set Eid. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        716002,
                                                        @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is Eid, the value is 0x00000001] The server does not include the PidTagMid (section 2.2.1.2.1) property in the message change header when the Eid flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.MessageSize) == SynchronizationExtraFlag.MessageSize)
                                                {
                                                    // Because the SynchronizationExtraFlag is MessageSize and the SynchronizationExtraFlag only set MessageSize. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        2195,
                                                        @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] The server MUST include the PidTagMessageSize property (section 2.2.1.6) in the message change header if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");

                                                    // The server include the PidTagMessageSize property in messageChange.
                                                    newMessageChange.IsPidTagMessageSizeExist = true;
                                                    ModelHelper.CaptureRequirement(
                                                        1365,
                                                        @"[In messageChangeHeader Element] [PidTagMessageSize,Conditional] PidTagMessageSize MUST be present if and only if the MessageSize flag of the SynchronizationExtraFlags field is set.");
                                                }
                                                else
                                                {
                                                    newMessageChange.IsPidTagMessageSizeExist = false;
                                                    ModelHelper.CaptureRequirement(
                                                       718001,
                                                       @"[In RopSynchronizationConfigure ROP Request Buffer] [SynchronizationExtraFlags, When the flag name is MessageSize, the value is 0x00000002] The server does not include the PidTagMessageSize property (section 2.2.1.6) in the message change header if the MessageSize flag of the SynchronizationExtraFlag field is not set.");
                                                }

                                                if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Progress) == SynchronizationFlag.Progress)
                                                {
                                                    // The server include ProgressPerMessage in messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = true;

                                                    // The message object is a normal message.
                                                    newMessageChange.FollowedFAIMessage = false;

                                                    // Because the SynchronizationExtraFlag is MessageSize and the SynchronizationExtraFlag only set MessageSize. So can cover requirement here.
                                                    ModelHelper.CaptureRequirement(
                                                        1171,
                                                        @"[In progressPerMessage Element]MUST be present if and only if the progessTotal element, as specified in
                                                        section 2.2.4.3.18, was output within the same ancestor contentsSync element, as specified in section
                                                        2.2.4.3.2.");
                                                    ModelHelper.CaptureRequirement(
                                                        1383,
                                                        @"[In progressPerMessage Element] [[PtypBoolean] 0x0000000B] otherwise[if the Message object that follows is not FAI] ,[the server returns] FALSE (0x00).");
                                                }
                                                else
                                                {
                                                    // The server don't include ProgressPerMessage in messageChange.
                                                    newMessageChange.IsProgressPerMessagePresent = false;
                                                    ModelHelper.CaptureRequirement(
                                                        1172,
                                                        @"[In progressPerMessage Element] [ProgressPerMessage Element] MUST NOT be present if the Progress flag of the SynchronizationFlags field was not set when configuring the synchronization download operation.");
                                                }

                                                // Add new messageChange to ContentsSync.
                                                abstractFastTransferStream.AbstractContentsSync.MessageInfo = abstractFastTransferStream.AbstractContentsSync.MessageInfo.Add(newMessageChange);

                                                // Because the MessageChange.followedFAIMessage default expect value is false and it is a normal messages.
                                                ModelHelper.CaptureRequirement(
                                                    2174,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the Normal flag of the SynchronizationFlags flag is set, the server MUST download information about changes to normal messages, as specified in section 2.2.4.3.11.");
                                            }
                                        }
                                        else
                                        {
                                            // Because the MessageChange.followedFAIMessage default expect value is false and it is a normal messages.
                                            ModelHelper.CaptureRequirement(
                                                2175,
                                                @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the Normal flag of the SynchronizationFlags field is not set, the server MUST NOT download information about changes to normal messages, as specified in section 2.2.4.3.11.");
                                        }
                                    }
                                }
                            }
                            #endregion
                            if (((currentDownloadContext.SynchronizationExtraflag & SynchronizationExtraFlag.OrderByDeliveryTime) == SynchronizationExtraFlag.OrderByDeliveryTime) && abstractFastTransferStream.AbstractContentsSync.MessageInfo.Count >= 2)
                            {
                                // The server MUST sort messages by the value of their PidTagMessageDeliveryTime property when generating a sequence of messageChange.
                                abstractFastTransferStream.AbstractContentsSync.IsSortByMessageDeliveryTime = true;
                                ModelHelper.CaptureRequirement(
                                    2198,
                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] If the OrderByDeliveryTime flag of the SynchronizationExtraFlags field is set, the server MUST sort messages by the value of their PidTagMessageDeliveryTime property ([MS-OXOMSG] section 2.2.3.9) when generating a sequence of messageChange elements for the FastTransfer stream, as specified in section 2.2.4.2.");

                                // The server MUST sort messages by the value of their PidTagLastModificationTime property when generating a sequence of messageChange.
                                abstractFastTransferStream.AbstractContentsSync.IsSortByLastModificationTime = true;
                                ModelHelper.CaptureRequirement(
                                    2199,
                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationExtraFlags Constraints] If the OrderByDeliveryTime flag of the SynchronizationExtraFlags field is set, the server MUST sort messages by the PidTagLastModificationTime property ([MS-OXPROPS] section 2.753) if the former[PidTagMessageDeliveryTime] is missing, when generating a sequence of messageChange elements for the FastTransfer stream, as specified in section 2.2.4.2.");
                            }

                            // Search the message in MessageContainer.
                            foreach (AbstractMessage tempMessage in currentConnection.MessageContainer)
                            {
                                if (tempMessage.FolderIdIndex == currentFolder.FolderIdIndex)
                                {
                                    // Identify whether the readStateChange is or not included in messageContent.
                                    if (!currentDownloadContext.UpdatedState.CnsetRead.Contains(tempMessage.ReadStateChangeNumberIndex))
                                    {
                                        if (tempMessage.ReadStateChangeNumberIndex != 0)
                                        {
                                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.ReadState) != SynchronizationFlag.ReadState)
                                            {
                                                // Download information about changes to the read state of messages.
                                                abstractFastTransferStream.AbstractContentsSync.IsReadStateChangesExist = false;
                                                ModelHelper.CaptureRequirement(
                                                    2171,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the ReadState flag of the SynchronizationFlags field is not set, the server MUST NOT download information about changes to the read state of messages, as specified in section 2.2.4.3.22.");

                                                // Download information about changes to the read state of messages.
                                                if (requirementContainer.ContainsKey(1193) && requirementContainer[1193])
                                                {
                                                    abstractFastTransferStream.AbstractContentsSync.IsReadStateChangesExist = false;
                                                    ModelHelper.CaptureRequirement(
                                                        1193,
    @"[In readStateChanges Element] Implementation does not present this element if the ReadState flag of the SynchronizationFlag field was not set when configuring the synchronization download operation. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                }
                                            }

                                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.ReadState) == SynchronizationFlag.ReadState)
                                            {
                                                if (requirementContainer.Keys.Contains(2665) && requirementContainer[2665])
                                                {
                                                    // Assign a new changeNumber for CnsetReadProperty.
                                                    abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetReadPropertyChangeNumber = true;

                                                    // Because identify the read state of a message changes by(currentMessage.IsRead != IReadstatus) and get new change number. So can cover this requirement here. 
                                                    ModelHelper.CaptureRequirement(2665, "[In Tracking Read State Changes] Implementation does assign a new value to the separate change number(the read state change number) on the message, whenever the read state of a message changes on the server. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                }

                                                abstractFastTransferStream.AbstractContentsSync.IsReadStateChangesExist = true;

                                                // Assign a new changeNumber for CnsetReadProperty.
                                                abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetReadPropertyChangeNumber = true;

                                                ModelHelper.CaptureRequirement(
                                                    2087,
                                                    "[In Tracking Read State Changes] An IDSET structure of change numbers associated with message read state transitions, either from read to unread, or unread to read (as determined by the PidTagMessageFlags property in [MS-OXCMSG] section 2.2.1.6) are included in the MetaTagCnsetRead property (section 2.2.1.1.4), which is part of the ICS state and is never directly set on any objects.");

                                                ModelHelper.CaptureRequirement(
                                                    3315,
                                                    "[In readStateChanges Element] This element MUST be present if there are changes to the read state of messages.");
                                            }
                                        }

                                        if (tempMessage.ReadStateChangeNumberIndex == 0)
                                        {
                                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.ReadState) == SynchronizationFlag.ReadState)
                                            {
                                                // Download information about changes to the read state of messages.
                                                abstractFastTransferStream.AbstractContentsSync.IsReadStateChangesExist = false;

                                                // Assign a new changeNumber for CnsetReadProperty.
                                                abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetReadPropertyChangeNumber = true;
                                                ModelHelper.CaptureRequirement(
                                                    2170,
                                                    @"[In Receiving a RopSynchronizationConfigure ROP Request] [SynchronizationType Constraints] If the ReadState flag of the SynchronizationFlags field is set, the server MUST also download information about changes to the read state of messages, as specified in section 2.2.4.3.22.");

                                                ModelHelper.CaptureRequirement(
                                                    2048,
                                                    @"[In Determining What Differences To Download] [For every object in the synchronization scope, servers MUST do the following:] If the ReadState flag of the SynchronizationFlags field is set, include the readStateChanges element, as specified in section 2.2.4.3.22, for messages that: 
	Do not have their change numbers for read and unread state in the MetaTagCnsetRead property (section 2.2.1.1.4)
	And are not FAI messages and have not had change information downloaded for them in this session.");
                                            }
                                        }
                                    }

                                    // The message change number is not in CnsetSeenFAI property.
                                    if (!currentDownloadContext.UpdatedState.CnsetSeenFAI.Contains(tempMessage.ChangeNumberIndex))
                                    {
                                        // The message is FAI message.
                                        if (tempMessage.IsFAImessage)
                                        {
                                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.FAI) == SynchronizationFlag.FAI)
                                            {
                                                ModelHelper.CaptureRequirement(
                                                    2044,
                                                        @"[In Determining What Differences To Download] [For every object in the synchronization scope, servers MUST do the following:] [Include the following syntactical elements in the FastTransfer stream of the OutputServerObject field of the FastTransfer download ROPs, as specified in section 2.2.3.1.1, if one of the following applies:] 
	                                                Include the messageChangeFull element, as specified in section 2.2.4.3.13, if the object specified by the InputServerObject field is an FAI message, meaning the PidTagAssociated property (section 2.2.1.5) is set to TRUE
                                                    And the FAI flag of the SynchronizationFlag field was set
                                                    And the change number is not included in the value of the MetaTagCnsetSeenFAI property (section 2.2.1.1.3).");

                                                // Add message changeNumber to CnsetSeenFAI of current download context.
                                                currentDownloadContext.UpdatedState.CnsetSeenFAI = currentDownloadContext.UpdatedState.CnsetSeenFAI.Add(tempMessage.ChangeNumberIndex);

                                                if (priorUploadOperation == MS_OXCFXICS.PriorOperation.RopSynchronizationImportMessageMove)
                                                {
                                                    // Assign a new changeNumber.
                                                    abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetSeenFAIPropertyChangeNumber = true;
                                                    ModelHelper.CaptureRequirement(
                                                        2247,
                                                        @"[In Receiving a RopSynchronizationImportMessageMove Request] Upon successful completion of this ROP, the ICS state on the synchronization context MUST be updated to include change numbers of messages in the destination folder in or MetaTagCnsetSeenFAI (section 2.2.1.1.3) property, when the message is an FAI message.");
                                                }

                                                // Assign a new changeNumber.
                                                abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetSeenFAIPropertyChangeNumber = true;

                                                if (requirementContainer.ContainsKey(218300301) && requirementContainer[218300301])
                                                {
                                                    abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetSeenFAIPropertyChangeNumber = false;
                                                }

                                                // Because this ROP is called after successful import of a new or changed object using ICS upload and the server represent the imported version in the MetaTagCnsetSeen property for FAI message, this requirement is captured.
                                                ModelHelper.CaptureRequirement(
                                                    190701,
                                                    @"[In Identifying Objects and Maintaining Change Numbers] Upon successful import of a new or changed object using ICS upload, the server MUST do the following when receiving RopSaveChangesMessage ROP: This is necessary because the server MUST be able to represent the imported version in the MetaTagCnsetSeenFAI (section 2.2.1.1.3) property for FAI message.");
                                            }
                                        }
                                        else
                                        {
                                            // SynchronizationFlag  is Normal of current Download Context.
                                            if ((currentDownloadContext.Synchronizationflag & SynchronizationFlag.Normal) == SynchronizationFlag.Normal)
                                            {
                                                if (!currentDownloadContext.UpdatedState.CnsetSeen.Contains(tempMessage.ChangeNumberIndex))
                                                {
                                                    // Add  Message changeNumber to CnsetSeen.
                                                    currentDownloadContext.UpdatedState.CnsetSeen = currentDownloadContext.UpdatedState.CnsetSeen.Add(tempMessage.ChangeNumberIndex);
                                                    ModelHelper.CaptureRequirement(
                                                        2043,
                                                            @"[In Determining What Differences To Download] [For every object in the synchronization scope, servers MUST do the following:] [Include the following syntactical elements in the FastTransfer stream of the OutputServerObject field of the FastTransfer download ROPs, as specified in section 2.2.3.1.1, if one of the following applies:] 
	Include the messageChange element, as specified in section 2.2.4.3.11, if the object specified by the InputServerObject field is a normal message
	And the Normal flag of the SynchronizationFlags field was set, as specified in section 2.2.3.2.1.1.1; 	And the change number is not included in the value of the MetaTagCnsetSeen property.");
                                                }
                                                else if (abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetReadPropertyChangeNumber)
                                                {
                                                    if (requirementContainer.ContainsKey(2666) && requirementContainer[2666])
                                                    {
                                                        abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = false;
                                                        ModelHelper.CaptureRequirement(
                                                            2666,
                                                            @"[In Tracking Read State Changes] Implementation does not modify the change number of the message unless other changes to a message were made at the same time. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                                    }
                                                }

                                                // Assign a new changeNumber.
                                                abstractFastTransferStream.AbstractContentsSync.FinalICSState.IsNewCnsetSeenPropertyChangeNumber = true;
                                                ModelHelper.CaptureRequirement(
                                                    2246,
                                                    @"[In Receiving a RopSynchronizationImportMessageMove Request]Upon successful completion of this ROP, the ICS 
                                                    state on the synchronization context MUST be updated to include change numbers of messages in the destination 
                                                    folder in the MetaTagCnsetSeen (section 2.2.1.1.2)  when the message is a normal message.");

                                                // Because this ROP is called after successful import of a new or changed object using ICS upload and the server represent the imported version in the MetaTagCnsetSeen property for normal message, this requirement is captured.
                                                ModelHelper.CaptureRequirement(
                                                    1907,
                                                    @"[In Identifying Objects and Maintaining Change Numbers] Upon successful import of a new or changed object using ICS upload, the server MUST do the following when receiving RopSaveChangesMessage ROP: This is necessary because the server MUST be able to represent the imported version in the MetaTagCnsetSeen (section 2.2.1.1.2) property for normal message.");
                                            }
                                        }
                                    }
                                }
                            }

                            // Set ContentsSync IdSetGiven value with IdsetGiven value of current download context.
                            abstractFastTransferStream.AbstractContentsSync.FinalICSState.IdSetGiven = currentDownloadContext.UpdatedState.IdsetGiven;

                            // Add ICS State to ICSStateContainer.
                            currentFolder.ICSStateContainer.Add(abstractFastTransferStream.AbstractContentsSync.FinalICSState.AbstractICSStateIndex, currentDownloadContext.UpdatedState);

                            // Update the FolderContainer.
                            currentConnection.FolderContainer = currentConnection.FolderContainer.Update(currentFolderIndex, currentFolder);
                            connections[serverId] = currentConnection;

                            break;
                        }

                    // In case of the FastTransferStreamType of download context include folderContent.
                    case FastTransferStreamType.folderContent:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyTo
                                && currentDownloadContext.ObjectType == ObjectType.Folder)
                            {
                                ModelHelper.CaptureRequirement(
                                    3324,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyTo, ROP request buffer field conditions is The InputServerObject field is a Folder object<22>, Root element in the produced FastTransfer stream is folderContent.");
                            }

                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyProperties && currentDownloadContext.ObjectType == ObjectType.Folder)
                            {
                                ModelHelper.CaptureRequirement(
                                    3325,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyProperties, ROP request buffer field conditions is The InputServerObject field is a Folder object<22>, Root element in the produced FastTransfer stream is folderContent.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentDownloadContext.ObjectType == ObjectType.Folder && sourOperation == SourceOperation.CopyProperties)
                            {
                                ModelHelper.CaptureRequirement(
                                    598,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyProperties, if the value of InputServerObject field is a Folder Object, Root element in FastTransfer stream is folderContent element.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentDownloadContext.ObjectType == ObjectType.Folder && sourOperation == SourceOperation.CopyTo)
                            {
                                ModelHelper.CaptureRequirement(
                                    595,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyTo, if the value of the InputServerObject field is a Folder Object, Root element in FastTransfer stream is folderContent element.");
                            }

                            // Initialize the entity variable
                            abstractFastTransferStream.AbstractFolderContent = new AbstractFolderContent
                            {
                                AbsFolderMessage = new AbstractFolderMessage()
                            };

                            bool isSubFolderExist = false;
                            AbstractFolder subFolder = new AbstractFolder();

                            // Search the folder container to find if the download folder contains subFolder
                            foreach (AbstractFolder tempSubFolder in currentConnection.FolderContainer)
                            {
                                if (currentFolder.SubFolderIds.Contains(tempSubFolder.FolderIdIndex))
                                {
                                    isSubFolderExist = true;
                                    subFolder = tempSubFolder;
                                    break;
                                }
                            }

                            // The downLaod folder's subFolder exist
                            if (isSubFolderExist)
                            {
                                if (subFolder.FolderPermission == PermissionLevels.None && (currentDownloadContext.CopyToCopyFlag == CopyToCopyFlags.Move || currentDownloadContext.CopyPropertiesCopyFlag == CopyPropertiesCopyFlags.Move))
                                {
                                    abstractFastTransferStream.AbstractFolderContent.IsNoPermissionObjNotOut = true;
                                    if (requirementContainer.ContainsKey(2667) && requirementContainer[2667])
                                    {
                                        ModelHelper.CaptureRequirement(2667, @"[In Receiving a RopFastTransferSourceCopyTo ROP Request] Implementation does not output any objects in a FastTransfer stream that the client does not have permissions to delete, if the Move flag of the CopyFlags field is set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                    }

                                    if (requirementContainer.ContainsKey(2669) && requirementContainer[2669])
                                    {
                                        ModelHelper.CaptureRequirement(
                                            2669,
                                            @"[In Receiving a RopFastTransferSourceCopyProperties Request] Implementation does not output any objects in a FastTransfer stream that the client does not have permissions to delete, if the Move flag of the CopyFlags field is specified for a download operation. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                    }
                                }
                            }
                            else if (subFolder.FolderPermission == PermissionLevels.FolderVisible && currentDownloadContext.CopyToCopyFlag == CopyToCopyFlags.Move)
                            {
                                if (requirementContainer.ContainsKey(118201) && requirementContainer[118201])
                                {
                                    abstractFastTransferStream.AbstractFolderContent.IsPidTagEcWarningOut = true;
                                } 

abstractFastTransferStream.AbstractFolderContent.AbsFolderMessage.MessageList.AbsMessage.AbsMessageContent.IsNoPermissionMessageNotOut = true;
                            }

                            // Search the currentDownloadContext.property to find if the specific property is required download
                            foreach (string propStr in currentDownloadContext.Property)
                            {
                                if (isSubFolderExist && subFolder.FolderPermission != PermissionLevels.None)
                                {
                                    // PidTagContainerHierarchy property is required to download
                                    if (propStr == "PidTagContainerHierarchy")
                                    {
                                        // CopyFolder operation will copy subFolder IFF CopySubfolders copyFlag is set
                                        if (((currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyFolder
                                            && currentDownloadContext.CopyFolderCopyFlag == CopyFolderCopyFlags.CopySubfolders) ||
                                            currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                            && currentConnection.LogonFolderType != LogonFlags.Ghosted)
                                        {
                                            if (currentConnection.LogonFolderType != LogonFlags.Ghosted || (requirementContainer.ContainsKey(1113) && requirementContainer[1113] && currentConnection.LogonFolderType == LogonFlags.Ghosted))
                                            {
                                                abstractFastTransferStream.AbstractFolderContent.IsSubFolderPrecededByPidTagFXDelProp = true;
                                                ModelHelper.CaptureRequirement(
                                                    1113,
                                                    @"[In folderContent Element] Under conditions specified in section 3.2.5.10, the PidTagContainerHierarchy property ([MS-OXPROPS] section 2.636) included in a subFolder element MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1).");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                        {
                                            if (currentConnection.LogonFolderType != LogonFlags.Ghosted || (requirementContainer.ContainsKey(1113) && requirementContainer[1113] && currentConnection.LogonFolderType == LogonFlags.Ghosted))
                                            {
                                                // The FastTransferOperation is FastTransferSourceCopyTo and the properties that not required in the propertyTag is download
                                                abstractFastTransferStream.AbstractFolderContent.IsSubFolderPrecededByPidTagFXDelProp = true;
                                                ModelHelper.CaptureRequirement(
                                                    1113,
                                                     @"[In folderContent Element] Under conditions specified in section 3.2.5.10, the PidTagContainerHierarchy property ([MS-OXPROPS] section 2.636) included in a subFolder element MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1).");
                                            }
                                        }
                                    }
                                }

                                if (currentConnection.LogonFolderType != LogonFlags.Ghosted)
                                {
                                    // PidTagFolderAssociatedContents property is required to download
                                    if (propStr == "PidTagFolderAssociatedContents")
                                    {
                                        if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                        {
                                            // The FastTransferOperation is FastTransferSourceCopyProperties and the properties that required in the propertyTag is download
                                            abstractFastTransferStream.AbstractFolderContent.AbsFolderMessage.IsFolderMessagesPrecededByPidTagFXDelProp = true;
                                            ModelHelper.CaptureRequirement(
                                                2620,
                                                @"[In folderMessages Element] Under conditions specified in section 3.2.5.10, when included in the folderMessages element, the PidTagFolderAssociatedContents ([MS-OXPROPS] section 2.690) and PidTagContainerContents ([MS-OXPROPS] section 2.634) properties MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1).");
                                        }
                                    }
                                    else
                                    {
                                        if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                        {
                                            // The FastTransferOperation is FastTransferSourceCopyTo and the properties that not required in the propertyTag is download
                                            abstractFastTransferStream.AbstractFolderContent.AbsFolderMessage.IsFolderMessagesPrecededByPidTagFXDelProp = true;
                                            ModelHelper.CaptureRequirement(
                                                2620,
                                                @"[In folderMessages Element] Under conditions specified in section 3.2.5.10, when included in the folderMessages element, the PidTagFolderAssociatedContents ([MS-OXPROPS] section 2.690) and PidTagContainerContents ([MS-OXPROPS] section 2.634) properties MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1).");
                                        }
                                    }

                                    // PidTagContainerContents property is required to download
                                    if (propStr == "PidTagContainerContents")
                                    {
                                        if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                        {
                                            // The FastTransferOperation is FastTransferSourceCopyProperties and the properties that required in the propertyTag is download
                                            abstractFastTransferStream.AbstractFolderContent.AbsFolderMessage.IsFolderMessagesPrecededByPidTagFXDelProp = true;
                                        }
                                    }
                                    else
                                    {
                                        if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                        {
                                            // The FastTransferOperation is FastTransferSourceCopyTo and the properties that not required in the propertyTag is download
                                            abstractFastTransferStream.AbstractFolderContent.AbsFolderMessage.IsFolderMessagesPrecededByPidTagFXDelProp = true;
                                        }
                                    }
                                }
                            }

                            break;
                        }

                    // In case of the FastTransferStreamType of download context include messageContent.
                    case FastTransferStreamType.MessageContent:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyTo && currentDownloadContext.ObjectType == ObjectType.Message)
                            {
                                ModelHelper.CaptureRequirement(
                                    3326,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyTo, ROP request buffer field conditions is The InputServerObject field is a Message object, Root element in the produced FastTransfer stream is messageContent.");
                            }

                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyProperties && currentDownloadContext.ObjectType == ObjectType.Message)
                            {
                                ModelHelper.CaptureRequirement(
                                    3327,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyProperties, ROP request buffer field conditions is The InputServerObject field is a Message object, Root element in the produced FastTransfer stream is messageContent.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentDownloadContext.ObjectType == ObjectType.Message && sourOperation == SourceOperation.CopyProperties)
                            {
                                ModelHelper.CaptureRequirement(
                                    596,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyProperties, if the value of the InputServerObject field is a Message object, Root element in FastTransfer stream is messageContent element.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentDownloadContext.ObjectType == ObjectType.Message && sourOperation == SourceOperation.CopyTo)
                            {
                                ModelHelper.CaptureRequirement(
                                    599,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyTo, if the value of the InputServerObject field is a Message object, Root element in FastTransfer stream is messageContent element.");
                            }

                            // Initialize the entity variable
                            abstractFastTransferStream.AbstractMessageContent = new AbstractMessageContent
                            {
                                AbsMessageChildren = new AbstractMessageChildren()
                            };

                            AbstractFolder messageParentFolder = new AbstractFolder();

                            // Search the MessagecCntianer to find the downLaodMessage
                            foreach (AbstractMessage cumessage in currentConnection.MessageContainer)
                            {
                                if (cumessage.MessageHandleIndex == currentDownloadContext.RelatedObjectHandleIndex)
                                {
                                    foreach (AbstractFolder cufolder in currentConnection.FolderContainer)
                                    {
                                        if (cufolder.FolderHandleIndex == cumessage.FolderHandleIndex)
                                        {
                                            messageParentFolder = cufolder;
                                            break;
                                        }
                                    }

                                    break;
                                }
                            }

                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyTo)
                            {
                                if ((currentDownloadContext.CopyToCopyFlag & CopyToCopyFlags.BestBody) == CopyToCopyFlags.BestBody)
                                {
                                    if (requirementContainer.ContainsKey(211501) && requirementContainer[211501])
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.IsRTFFormat = false;
                                        ModelHelper.CaptureRequirement(211501, @"[In Receiving a RopFastTransferSourceCopyTo ROP Request] Implementation does output the message body, and the body of the Embedded Message object, in their original format, if the BestBody flag of the CopyFlags field is set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                    }

                                    if (requirementContainer.ContainsKey(3118003) && requirementContainer[3118003])
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.IsRTFFormat = false;
                                        ModelHelper.CaptureRequirement(3118003, @"[In Appendix A: Product Behavior] Implementation does support this flag [BestBody flag] [in RopFastTransferSourceCopyTo ROP]. (<3> Section 2.2.3.1.1.1.1: Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
                                    }
                                }
                            }

                            #region Verify Requirements about Sendoptions
                            if ((currentDownloadContext.Sendoptions & SendOptionAlls.UseCpid) == SendOptionAlls.UseCpid && connections.Count > 1)
                            {
                                if (currentDownloadContext.Sendoptions == SendOptionAlls.UseCpid)
                                {
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                                    ModelHelper.CaptureRequirement(
                                             3453,
                                             @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When UseCpid only, if the properties are stored in Unicode on the server, the server MUST return the properties using the Unicode code page (code page property type 0x84B0).");

                                    if (currentDownloadContext.RelatedFastTransferOperation != EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                    {
                                        if (requirementContainer.ContainsKey(3454) && requirementContainer[3454])
                                        {
                                            abstractFastTransferStream.AbstractMessageContent.StringPropertiesInOtherCodePage = true;
                                            ModelHelper.CaptureRequirement(
                                                   3454,
                                                   @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When UseCpid only, otherwise[the properties are not stored in Unicode on the server] the server MUST send the string using the code page property type of the code page in which the property is stored on the server.");
                                        }
                                    }

                                    ModelHelper.CaptureRequirement(
                                            3452,
                                            @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When UseCpid only, String properties MUST be output using code page property types, as specified in section 2.2.4.1.1.1.");
                                }
                                else if (currentDownloadContext.Sendoptions == (SendOptionAlls.Unicode | SendOptionAlls.UseCpid))
                                {
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                                    ModelHelper.CaptureRequirement(
                                             3457,
                                             @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When UseCpid and Unicode, If string properties are stored in Unicode on the server, the server MUST return the properties using the Unicode code page (code page property type 0x84B0).");

                                    if (currentDownloadContext.RelatedFastTransferOperation != EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                    {
                                        if (requirementContainer.ContainsKey(3454) && requirementContainer[3454])
                                        {
                                            abstractFastTransferStream.AbstractMessageContent.StringPropertiesInOtherCodePage = true;
                                            ModelHelper.CaptureRequirement(
                                                      3780,
                                                      @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] UseCpid and Unicode: If string properties are not stored in Unicode on the server, the server MUST send the string using the code page property type of the code page in which the property is stored on the server.");
                                        }
                                    }

                                    ModelHelper.CaptureRequirement(
                                              3456,
                                              @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When UseCpid and Unicode, String properties MUST be output using code page property types, as specified in section 2.2.4.1.1.1. ");
                                }
                                else if (currentDownloadContext.Sendoptions == (SendOptionAlls.UseCpid | SendOptionAlls.ForceUnicode))
                                {
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                                    ModelHelper.CaptureRequirement(
                                                3782,
                                                @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] UseCpid and ForceUnicode: String properties MUST be output using the Unicode code page (code page property type 0x84B0).");
                                }
                                else if (currentDownloadContext.Sendoptions == (SendOptionAlls.UseCpid | SendOptionAlls.ForceUnicode | SendOptionAlls.Unicode))
                                {
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicodeCodePage = true;
                                    ModelHelper.CaptureRequirement(
                                                 3459,
                                                 @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] WhenUseCpid, Unicode, and ForceUnicode, The combination of the UseCpid and Unicode flags is the ForUpload flag.
                                                String properties MUST be output using the Unicode code page (code page property type 0x84B0).");
                                }
                            }
                            else
                            {
                                if (currentDownloadContext.Sendoptions == SendOptionAlls.ForceUnicode)
                                {
                                    // The String properties is saved in the server using unicode format in this test environment ,so the string properties must out in Unicode format
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicode = true;
                                    ModelHelper.CaptureRequirement(
                                                3451,
                                                @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When ForceUnicode only, String properties MUST be output in Unicode with a property type of PtypUnicode.");
                                }
                                else if (((currentDownloadContext.Sendoptions & SendOptionAlls.Unicode) != SendOptionAlls.Unicode) && (currentDownloadContext.Sendoptions != (SendOptionAlls.ForceUnicode | SendOptionAlls.Unicode)))
                                {
                                    // String properties MUST be output in code page
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicode = false;

                                    ModelHelper.CaptureRequirement(
                                        3447,
                                        @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When none of the three flags[Unicode, ForceUnicode, and UseCpid] are set, String properties MUST be output in the code page set on the current connection with a property type of PtypString8 ([MS-OXCDATA] section 2.11.1). ");
                                }
                                else if (currentDownloadContext.Sendoptions == (SendOptionAlls.ForceUnicode | SendOptionAlls.Unicode))
                                {
                                    // String properties MUST be output in Unicode
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicode = true;

                                    ModelHelper.CaptureRequirement(
                                        3455,
                                        @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When Unicode and ForceUnicode, String properties MUST be output in Unicode with a property type of PtypUnicode.");
                                }
                                else if (currentDownloadContext.Sendoptions == SendOptionAlls.Unicode)
                                {
                                    // The string properties is saved in the server using Unicode format in this test environment ,so the string properties must out in Unicode format
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicode = true;

                                    ModelHelper.CaptureRequirement(
                                       3448,
                                       @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When Unicode only, String properties MUST be output either in Unicode with a property type of PtypUnicode ([MS-OXCDATA] section 2.11.1), or in the code page set on the current connection with a property type of PtypString8. ");

                                    ModelHelper.CaptureRequirement(
                                        3449,
                                        @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When Unicode only, if the properties are stored in Unicode on the server, the server MUST return the properties in Unicode. ");

                                    ModelHelper.CaptureRequirement(
                                        3450,
                                        @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When Unicode only, if the properties are not stored in Unicode on the server, the server MUST return the properties in the code page set on the current connection.");
                                }
                                else if (currentDownloadContext.Sendoptions == SendOptionAlls.ForceUnicode)
                                {
                                    // The string properties is saved in the server using unicode format in this test environment ,so the string properties must out in Unicode format
                                    abstractFastTransferStream.AbstractMessageContent.StringPropertiesInUnicode = true;
                                    ModelHelper.CaptureRequirement(
                                                3451,
                                                @"[In Receiving a RopFastTransferSourceCopyTo Request] [valid combinations of the Unicode, ForceUnicode, and UseCpid flags of the SendOptions field] When ForceUnicode only, String properties MUST be output in Unicode with a property type of PtypUnicode.");
                                }
                            }
                            #endregion

                            if (!currentDownloadContext.IsLevelTrue)
                            {
                                bool isPidTagMessageAttachmentsExist = false;
                                bool isPidTagMessageRecipientsExist = false;

                                // Search the currentDownloadContext.property to find if the specific property is required to download
                                foreach (string propStr in currentDownloadContext.Property)
                                {
                                    if (propStr == "PidTagMessageAttachments")
                                    {
                                        isPidTagMessageAttachmentsExist = true;
                                    }

                                    if (propStr == "PidTagMessageRecipients")
                                    {
                                        isPidTagMessageRecipientsExist = true;
                                    }

                                    // CopyTo operation's propertyTags specific the properties not to download
                                    if (currentDownloadContext.RelatedFastTransferOperation != EnumFastTransferOperation.FastTransferSourceCopyTo)
                                    {
                                        // The PidTagMessageRecipients property is required to download
                                        if (propStr == "PidTagMessageAttachments")
                                        {
                                            // The PidTagMessageAttachments property is required to download
                                            abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.AttachmentPrecededByPidTagFXDelProp = true;
                                            ModelHelper.CaptureRequirement(
                                                3304,
                                                @"[In messageChildren Element] Under the conditions specified in section 3.2.5.10 [Effect of Property and Subobject Filters on Download] , the PidTagMessageRecipients ([MS-OXPROPS] section 2.784) property included in a recipient element and the PidTagMessageAttachments ([MS-OXPROPS] section 2.774) property included in an attachment element MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1)and .");

                                            // The AttachmentPrecededByPidTagFXDelProp true means server outputs the MetaTagFXDelProp property before outputting subobjects, such as attachment.
                                            ModelHelper.CaptureRequirement(
                                                2276,
                                                @"[In Effect of Property and Subobject Filters on Download] Whenever subobject filters have an effect, servers MUST output a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1) immediately before outputting subobjects of a particular type, to differentiate between the cases where a set of subobjects (such as attachments or recipients) was filtered in, but was empty, and where it was filtered out.");

                                            ModelHelper.CaptureRequirement(
                                                3464,
                                                @"[In Receiving a RopFastTransferSourceCopyProperties Request] If the Level field is set to 0x00, the server MUST copy descendant subobjects by using the property list specified by the PropertyTags field. ");

                                            ModelHelper.CaptureRequirement(
                                                3783,
                                                @"[In Receiving a RopFastTransferSourceCopyProperties Request] Subobjects are not copied unless listed in the value of the PropertyTags field.");
                                        }

                                        if (propStr == "PidTagMessageRecipients")
                                        {
                                            // The PidTagMessageRecipients property is required to download
                                            abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.RecipientPrecededByPidTagFXDelProp = true;
                                            ModelHelper.CaptureRequirement(
                                                3304,
                                                @"[In messageChildren Element] Under the conditions specified in section 3.2.5.10 [Effect of Property and Subobject Filters on Download] , the PidTagMessageRecipients ([MS-OXPROPS] section 2.784) property included in a recipient element and the PidTagMessageAttachments ([MS-OXPROPS] section 2.774) property included in an attachment element MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1)and .");

                                            // The RecipientPrecededByPidTagFXDelProp true means server outputs the MetaTagFXDelProp property before outputting subobjects, such as recipients.
                                            ModelHelper.CaptureRequirement(
                                                2276,
                                                @"[In Effect of Property and Subobject Filters on Download] Whenever subobject filters have an effect, servers MUST output a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1) immediately before outputting subobjects of a particular type, to differentiate between the cases where a set of subobjects (such as attachments or recipients) was filtered in, but was empty, and where it was filtered out.");
                                        }
                                    }
                                    else if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                    {
                                        if (propStr == "PidTagMessageAttachments")
                                        {
                                            // The PidTagMessageAttachments property is not required to download
                                            abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.AttachmentPrecededByPidTagFXDelProp = false;

                                            ModelHelper.CaptureRequirement(
                                                3439,
                                                @"[In Receiving a RopFastTransferSourceCopyTo ROP Request] If the Level field is set to 0x00, the server MUST copy descendant subobjects by using the property list specified by the PropertyTags field. ");

                                            ModelHelper.CaptureRequirement(
                                                3440,
                                                @"[In Receiving a RopFastTransferSourceCopyTo ROP Request] Subobjects are only copied when they are not listed in the value of the PropertyTags field. ");
                                        }

                                        if (propStr == "PidTagMessageRecipients")
                                        {
                                            // The PidTagMessageRecipients property is not required to download
                                            abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.RecipientPrecededByPidTagFXDelProp = false;
                                        }
                                    }
                                }

                                if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo && currentConnection.AttachmentContainer.Count > 0)
                                {
                                    if (!isPidTagMessageAttachmentsExist)
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.AttachmentPrecededByPidTagFXDelProp = true;

                                        ModelHelper.CaptureRequirement(
                                            3304,
                                            @"[In messageChildren Element] Under the conditions specified in section 3.2.5.10 [Effect of Property and Subobject Filters on Download] , the PidTagMessageRecipients ([MS-OXPROPS] section 2.784) property included in a recipient element and the PidTagMessageAttachments ([MS-OXPROPS] section 2.774) property included in an attachment element MUST be preceded by a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1)and .");
                                    }
                                }

                                if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                {
                                    if (!isPidTagMessageRecipientsExist)
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.RecipientPrecededByPidTagFXDelProp = true;

                                        // The RecipientPrecededByPidTagFXDelProp true means server outputs the MetaTagFXDelProp property before outputting subobjects, such as recipients.
                                        ModelHelper.CaptureRequirement(
                                            2276,
                                            @"[In Effect of Property and Subobject Filters on Download] Whenever subobject filters have an effect, servers MUST output a MetaTagFXDelProp meta-property (section 2.2.4.1.5.1) immediately before outputting subobjects of a particular type, to differentiate between the cases where a set of subobjects (such as attachments or recipients) was filtered in, but was empty, and where it was filtered out.");
                                    }
                                }
                            }
                            else
                            {
                                abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.AttachmentPrecededByPidTagFXDelProp = false;
                                abstractFastTransferStream.AbstractMessageContent.AbsMessageChildren.RecipientPrecededByPidTagFXDelProp = false;

                                if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyTo)
                                {
                                    ModelHelper.CaptureRequirement(
                                        3441,
                                        @"[In Receiving a RopFastTransferSourceCopyTo ROP Request] If the Level field is set to a nonzero value, the server MUST exclude all descendant subobjects from being copied.");
                                }

                                if (currentDownloadContext.RelatedFastTransferOperation == EnumFastTransferOperation.FastTransferSourceCopyProperties)
                                {
                                    ModelHelper.CaptureRequirement(
                                        3465,
                                        @"[In Receiving a RopFastTransferSourceCopyProperties Request] If the Level field is set to a nonzero value, the server MUST exclude all descendant subobjects from being copied.");
                                }
                            }

                            break;
                        }

                    // In case of the FastTransferStreamType of download context include messageList.
                    case FastTransferStreamType.MessageList:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyMessage)
                            {
                                ModelHelper.CaptureRequirement(
                                    3330,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyMessages, ROP request buffer field conditions is always, Root element in the produced FastTransfer stream is messageList.");
                            }

                            if (priorOperation == MS_OXCFXICS.PriorOperation.RopFastTransferDestinationConfigure && sourOperation == SourceOperation.CopyMessages)
                            {
                                ModelHelper.CaptureRequirement(
                                           601,
                                           @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyMessages, Root element in FastTransfer stream is messageList.");
                            }

                            // Initialize the entity variable
                            abstractFastTransferStream.AbstractMessageList = new AbstractMessageList
                            {
                                AbsMessage = new AbsMessage
                                {
                                    AbsMessageContent = new AbstractMessageContent()
                                }
                            };

                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyMessage)
                            {
                                if ((currentDownloadContext.CopyMessageCopyFlag & RopFastTransferSourceCopyMessagesCopyFlags.BestBody) == RopFastTransferSourceCopyMessagesCopyFlags.BestBody)
                                {
                                    if (requirementContainer.ContainsKey(211601) && requirementContainer[211601])
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.IsRTFFormat = false;
                                        ModelHelper.CaptureRequirement(211601, @"[In Receiving a RopFastTransferSourceCopyMessages ROP Request] Implementation does output the message body, and the body of the Embedded Message object, in their original format, If the BestBody flag of the CopyFlags field is set. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                    }

                                    if (requirementContainer.ContainsKey(499003) && requirementContainer[499003])
                                    {
                                        abstractFastTransferStream.AbstractMessageContent.IsRTFFormat = false;
                                        ModelHelper.CaptureRequirement(499003, @"[In Appendix A: Product Behavior] Implementation does support this flag [BestBody flag] [in RopFastTransferSourceCopyMessages ROP]. (<5> Section 2.2.3.1.1.3.1: Microsoft Exchange Server 2007 and Microsoft Exchange Server 2010 follow this behavior.)");
                                    }
                                }
                            }

                            // If the folder permission is set to None.
                            if (currentFolder.FolderPermission == PermissionLevels.None || currentFolder.FolderPermission == PermissionLevels.FolderVisible)
                            {
                                if (currentDownloadContext.CopyMessageCopyFlag == RopFastTransferSourceCopyMessagesCopyFlags.Move)
                                {
                                    if (requirementContainer.ContainsKey(2631) && requirementContainer[2631])
                                    {
                                        // The server doesn't output any objects in a FastTransfer stream that the client does not have permissions to delete
                                        abstractFastTransferStream.AbstractMessageList.AbsMessage.AbsMessageContent.IsNoPermissionMessageNotOut = true;
                                        ModelHelper.CaptureRequirement(
                                            2631,
                                            @"[In Receiving a RopFastTransferSourceCopyMessages ROP Request] Implementation does not output any objects in a FastTransfer stream that the client does not have permissions to delete, If the Move flag of the CopyFlags field is set for a download operation. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                    }
                                }

                                if (requirementContainer.ContainsKey(1168) && requirementContainer[1168])
                                {
                                    // The server doesn't have the permissions necessary to access PidTagEcWarning if the folder permission is set to None.
                                    abstractFastTransferStream.AbstractMessageList.IsPidTagEcWarningOut = true;
                                    ModelHelper.CaptureRequirement(
                                        1168,
                                        @"[In messageList Element] Implementation does output MetaTagEcWarning meta-property (section 2.2.4.1.5.2) if a client does not have the permissions necessary to access it, as specified in section 3.2.5.8.1. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }

                                if (requirementContainer.ContainsKey(34381) && requirementContainer[34381])
                                {
                                    // The server doesn't have the permissions necessary to access PidTagEcWarning if the folder permission is set to None.
                                    abstractFastTransferStream.AbstractMessageList.IsPidTagEcWarningOut = true;
                                    ModelHelper.CaptureRequirement(
                                        34381,
                                        @"[In Download] Implementation does output the MetaTagEcWarning meta-property (section 2.2.4.1.5.2) in a FastTransfer stream 
                                        if a permission check for an object fails, wherever allowed by its syntactical structure, to signal a client about 
                                        incomplete content. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                                }
                            }

                            break;
                        }

                    // In case of the FastTransferStreamType of download context include topFolder.
                    case FastTransferStreamType.TopFolder:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyFolder)
                            {
                                ModelHelper.CaptureRequirement(
                                    3331,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyFolder, ROP request buffer field conditions is always, Root element in the produced FastTransfer stream is topFolder.");
                            }

                            if (priorOperation == MS_OXCFXICS.PriorOperation.RopFastTransferDestinationConfigure && sourOperation == SourceOperation.CopyFolder)
                            {
                                ModelHelper.CaptureRequirement(
                                        602,
                                        @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyFolder, Root element in FastTransfer stream is topFolder.");
                            }

                            // If the logon folder is Ghosted folder
                            if (currentConnection.LogonFolderType == LogonFlags.Ghosted && requirementContainer.ContainsKey(1111) && requirementContainer[1111])
                            {
                                // The PidTagNewFXFolder meta-property MUST be output for the Ghosted folder
                                abstractFastTransferStream.AbstractTopFolder.AbsFolderContent.IsPidTagNewFXFolderOut = true;
                                ModelHelper.CaptureRequirement(
                                    1111,
                                    @"[In folderContent Element] [If there is a valid replica (1) of the public folder on the server and the folder content 
                                    has not replicated to the server yet, the folder content is not included in the FastTransfer stream as part of the 
                                    folderContent element] Implementation does not include any data following the MetaTagNewFXFolder meta-property in 
                                    the buffer returned by the RopFastTransferSourceGetBuffer ROP (section 2.2.3.1.1.5), although additional data can 
                                    be included in the FastTransfer stream. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                            }

                            // Identify whether the subFolder is existent or not.
                            bool isSubFolderExist = false;
                            AbstractFolder subFolder = new AbstractFolder();

                            // Search the folder container to find if the download folder has subFolder
                            foreach (AbstractFolder tempSubFolder in currentConnection.FolderContainer)
                            {
                                if (currentFolder.SubFolderIds.Contains(tempSubFolder.FolderIdIndex))
                                {
                                    isSubFolderExist = true;
                                    subFolder = tempSubFolder;
                                    break;
                                }
                            }

                            if (isSubFolderExist)
                            {
                                // Identify folder Permission is available.
                                if (subFolder.FolderPermission != PermissionLevels.None)
                                {
                                    if ((currentDownloadContext.CopyFolderCopyFlag & CopyFolderCopyFlags.CopySubfolders) == CopyFolderCopyFlags.CopySubfolders || (currentDownloadContext.CopyFolderCopyFlag & CopyFolderCopyFlags.Move) == CopyFolderCopyFlags.Move)
                                    {
                                        // The server recursively include the subFolders of the folder specified in the InputServerObject in the scope.
                                        abstractFastTransferStream.AbstractTopFolder.SubFolderInScope = true;
                                    }

                                    if ((currentDownloadContext.CopyFolderCopyFlag & CopyFolderCopyFlags.Move) == CopyFolderCopyFlags.Move && (currentDownloadContext.CopyFolderCopyFlag & CopyFolderCopyFlags.CopySubfolders) != CopyFolderCopyFlags.CopySubfolders)
                                    {
                                        ModelHelper.CaptureRequirement(
                                            3481,
                                            @"[In Receiving a RopFastTransferSourceCopyFolder ROP Request] If the Move flag of the CopyFlags field is set and the CopySubfolders flag is not set, the server MUST recursively include the subfolders of the folder specified in the InputServerObject field in the scope.");
                                    }
                                }
                            }

                            break;
                        }

                    // In case of the FastTransferStreamType of download context include messageList.
                    case FastTransferStreamType.attachmentContent:
                        {
                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyTo && currentDownloadContext.ObjectType == ObjectType.Attachment)
                            {
                                ModelHelper.CaptureRequirement(
                                    3328,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyTo, ROP request buffer field conditions is The InputServerObject field is an Attachment object<23>, Root element in the produced FastTransfer stream is attachmentContent.");
                            }

                            if (priorDownloadOperation == PriorDownloadOperation.RopFastTransferSourceCopyProperties && currentDownloadContext.ObjectType == ObjectType.Attachment)
                            {
                                ModelHelper.CaptureRequirement(
                                    3329,
                                    @"[In FastTransfer Streams in ROPs] When ROP that initiates an operation is RopFastTranserSourceCopyProperties, ROP request buffer field conditions is The InputServerObject field is an Attachment object<23>, Root element in the produced FastTransfer stream is attachmentContent.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentConnection.AttachmentContainer.Count > 0 && sourOperation == SourceOperation.CopyTo)
                            {
                                ModelHelper.CaptureRequirement(
                                    597,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyTo, if the value of the InputServerObject field is an Attachment object, Root element in FastTransfer stream is attachmentContent element.");
                            }

                            if (priorOperation == PriorOperation.RopFastTransferDestinationConfigure && currentDownloadContext.ObjectType == ObjectType.Attachment && sourOperation == SourceOperation.CopyProperties)
                            {
                                ModelHelper.CaptureRequirement(
                                    600,
                                    @"[In RopFastTransferDestinationConfigure ROP Request Buffer] [SourceOperation] When SourceOperation enumeration value is CopyProperties, if the value of the InputServerObject field is an Attachment object, Root element in FastTransfer stream is attachmentContent element.");
                            }

                            break;
                        }

                    default:
                        break;
                }
            }

            if (result == RopResult.Success)
            {
                // If the server returns success result, which means the RopFastTransferSourceGetBuffer ROP downloads the next portion of a FastTransfer stream successfully. Then this requirement can be captured.
                ModelHelper.CaptureRequirement(
                        532,
                        @"[In RopFastTransferSourceGetBuffer ROP] The RopFastTransferSourceGetBuffer ROP ([MS-OXCROPS] section 2.2.12.3) downloads the next portion of a FastTransfer stream that is produced by a previously configured download operation.");
            }

            return result;
        }

        /// <summary>
        ///  Initializes a FastTransfer operation for uploading content encoded in a client-provided FastTransfer stream into a mailbox
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="uploadContextHandleIndex">Configure handle's index.</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        [Rule(Action = ("FastTransferDestinationConfigure(serverId,objHandleIndex,option,copyFlag,out uploadContextHandleIndex)/result"))]
        public static RopResult FastTransferDestinationConfigure(int serverId, int objHandleIndex, SourceOperation option, FastTransferDestinationConfigureCopyFlags copyFlag, out int uploadContextHandleIndex)
        {
            // The construction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Initialize return value.
            RopResult result = RopResult.InvalidParameter;

            if (requirementContainer.ContainsKey(3492001))
            {
                if ((copyFlag == FastTransferDestinationConfigureCopyFlags.Invalid) &&
                    requirementContainer[3492001] == true)
                {
                    // FastTransferDestinationConfigureCopyFlags is invalid parameter and exchange server version is not ExchangeServer2007 .
                    uploadContextHandleIndex = -1;
                    ModelHelper.CaptureRequirement(
                                  3492001,
                                  @"[In Appendix A: Product Behavior] If unknown flags in the CopyFlags field are set, implementation does fail the operation. <36> Section 3.2.5.8.2.1: Exchange 2010, Exchange 2013 and Exchange 2016 fail the ROP [RopFastTransferDestinationConfigure ROP] if unknown bit flags in the CopyFlags field are set.");

                    return result;
                }
            }

            if (option == SourceOperation.CopyProperties || option == SourceOperation.CopyTo || option == SourceOperation.CopyFolder || option == SourceOperation.CopyMessages)
            {
                priorOperation = MS_OXCFXICS.PriorOperation.RopFastTransferDestinationConfigure;
            }

            sourOperation = option;

            // Create a new Upload context.
            AbstractUploadInfo uploadInfo = new AbstractUploadInfo();

            // Set value for upload context.
            uploadContextHandleIndex = AdapterHelper.GetHandleIndex();
            uploadInfo.UploadHandleIndex = uploadContextHandleIndex;
            ConnectionData changeConnection = connections[serverId];
            connections.Remove(serverId);

            // Add the new Upload context to UploadContextContainer.
            changeConnection.UploadContextContainer = changeConnection.UploadContextContainer.Add(uploadInfo);
            connections.Add(serverId, changeConnection);
            result = RopResult.Success;
            ModelHelper.CaptureRequirement(
                   581,
                   @"[In RopFastTransferDestinationConfigure ROP] The RopFastTransferDestinationConfigure ROP ([MS-OXCROPS] section 2.2.12.1) initializes a FastTransfer operation for uploading content encoded in a client-provided FastTransfer stream into a mailbox.");
            if (requirementContainer.ContainsKey(3492002))
            {
                if ((copyFlag == FastTransferDestinationConfigureCopyFlags.Invalid) &&
                    requirementContainer[3492002] == true)
                {
                    // Exchange 2007 ignore unknown values of the CopyFlags field.
                    ModelHelper.CaptureRequirement(
                                  3492002,
                                  @"[In Appendix A: Product Behavior] If unknown flags in the CopyFlags field are set, implementation does not fail the operation. <37> Section 3.2.5.8.2.1: Exchange 2007 ignore unknown values of the CopyFlags field.");

                    return result;
                }
            }

            return result;
        }

        /// <summary>
        ///  Uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">A fastTransfer stream object handle index.</param>
        /// <param name="transferDataIndex">Transfer data index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("FastTransferDestinationPutBuffer(serverId,uploadContextHandleIndex,transferDataIndex)/result"))]
        public static RopResult FastTransferDestinationPutBuffer(int serverId, int uploadContextHandleIndex, int transferDataIndex)
        {
            // The construction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // serverHandleIndex is Invalid Parameter.
            if (uploadContextHandleIndex < 0 || transferDataIndex <= 0)
            {
                return RopResult.InvalidParameter;
            }

            ModelHelper.CaptureRequirement(
                614,
                 @"[In RopFastTransferDestinationPutBuffer ROP] The RopFastTransferDestinationPutBuffer ROP ([MS-OXCROPS] section 2.2.12.2) uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.");

            return RopResult.Success;
        }

        /// <summary>
        /// Tell the server of another server's version.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="serverHandleIndex">Server object handle index in handle container.</param>
        /// <param name="otherServerId">Another server's id.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        [Rule(Action = ("TellVersion(serverId,serverHandleIndex,otherServerId)/result"))]
        public static RopResult TellVersion(int serverId, int serverHandleIndex, int otherServerId)
        {
            // The construction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));
            if (serverHandleIndex < 0)
            {
                // serverHandleIndex is Invalid Parameter.
                return RopResult.InvalidParameter;
            }
            else
            {
                ModelHelper.CaptureRequirement(
                    572,
                    @"[In RopTellVersion ROP] The RopTellVersion ROP ([MS-OXCROPS] section 2.2.12.8) is used to provide the version of one server to another server that is participating in the server-to-client-to-server upload, as specified in section 3.3.4.2.1.");

                return RopResult.Success;
            }
        }

        /// <summary>
        /// Tell the server of another server's version.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">Folder Handle index.</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count returned from server.</param>
        /// <returns>Indicate the result of this rop operation.</returns>
        [Rule(Action = ("GetContentsTable(serverId,folderHandleIndex,deleteFlags,out rowCount)/result"))]
        public static RopResult GetContentsTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount)
        {
            // The construction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Get current ConnectionData value.
            ConnectionData currentConnection = connections[serverId];

            // Create a new currentDownloadContext.
            AbstractUploadInfo currentUploadContext = new AbstractUploadInfo();
            rowCount = -1;

            // Find the current Upload Context
            foreach (AbstractUploadInfo tempUploadContext in currentConnection.UploadContextContainer)
            {
                if (tempUploadContext.RelatedFastTransferOperation == EnumFastTransferOperation.SynchronizationImportDeletes)
                {
                    currentUploadContext = tempUploadContext;
                }
            }

            if (folderHandleIndex < 0)
            {
                rowCount = -1;

                // serverHandleIndex is Invalid Parameter.
                return RopResult.InvalidParameter;
            }
            else
            {
                if (deleteFlags == DeleteFlags.Initial)
                {
                    rowCount = 0;
                    return RopResult.Success;
                }
                else if (deleteFlags == DeleteFlags.HardDeleteCheck)
                {
                    rowCount = 0;
                }
                else if (deleteFlags == DeleteFlags.SoftDeleteCheck)
                {
                    rowCount = softDeleteMessageCount;
                    if (priorOperation == MS_OXCFXICS.PriorOperation.RopSynchronizationImportMessageMove)
                    {
                        rowCount = 1;
                    }
                }

                return RopResult.Success;
            }
        }

        /// <summary>
        /// Tell the server of another server's version.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">Folder Handle index.</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count returned from server.</param>
        /// <returns>Indicate the result of this rop operation.</returns>
        [Rule(Action = ("GetHierarchyTable(serverId,folderHandleIndex,deleteFlags,out rowCount)/result"))]
        public static RopResult GetHierarchyTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount)
        {
            // The construction conditions.
            Condition.IsTrue(connections.Count > 0);
            Condition.IsTrue(connections.Keys.Contains(serverId));

            // Get current ConnectionData value.
            ConnectionData currentConnection = connections[serverId];

            // Create a new currentDownloadContext.
            AbstractUploadInfo currentUploadContext = new AbstractUploadInfo();

            // Initialize the rowCount value.
            rowCount = -1;

            // Find the current Upload Context
            foreach (AbstractUploadInfo tempUploadContext in currentConnection.UploadContextContainer)
            {
                if (tempUploadContext.RelatedFastTransferOperation == EnumFastTransferOperation.SynchronizationImportDeletes)
                {
                    currentUploadContext = tempUploadContext;
                }
            }

            if (folderHandleIndex < 0)
            {
                rowCount = -1;

                // serverHandleIndex is Invalid Parameter.
                return RopResult.InvalidParameter;
            }
            else
            {
                if (deleteFlags == DeleteFlags.Initial)
                {
                    rowCount = 0;
                    return RopResult.Success;
                }
                else if (deleteFlags == DeleteFlags.HardDeleteCheck)
                {
                    if (requirementContainer.ContainsKey(90205002) && !requirementContainer[90205002])
                    {
                        rowCount = -1;
                    }
                    else
                    {
                        rowCount = 0;
                    }
                }
                else if (deleteFlags == DeleteFlags.SoftDeleteCheck)
                {
                    rowCount = softDeleteFolderCount;
                }

                return RopResult.Success;
            }
        }

        #endregion

        #region Others

        /// <summary>
        /// Validate if the given two buffer is equal
        /// </summary>
        /// <param name="operation">The Enumeration Fast Transfer Operation</param>
        /// <param name="firstBufferIndex">The first Buffer's index</param>
        /// <param name="secondBufferIndex">The second buffer's index</param>
        /// <returns>Returns true only if the two buffers are equal</returns>
        [Rule(Action = ("AreEqual(operation,firstBufferIndex,secondBufferIndex)/result"))]
        private static bool AreEqual(EnumFastTransferOperation operation, int firstBufferIndex, int secondBufferIndex)
        {
            // Identify whether the firstBuffer and the secondBuffer are equal or not.
            if (firstBufferIndex <= 0 || secondBufferIndex <= 0)
            {
                return false;
            }

            bool returnValue = true;
            return returnValue;
        }

        #endregion
    }
}