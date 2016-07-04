namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Traditional test cases of MS-OXCNOTIF.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// An instance of logonId.
        /// </summary>
        private const byte LogonId = 0;

        /// <summary>
        /// Whether the context used to send notification is connected.
        /// </summary>
        private bool isSenderContextConnected;

        /// <summary>
        /// Whether the context used to receive notification is connected.
        /// </summary>
        private bool isReceiverContextConnected;

        /// <summary>
        /// An instance of IOXCNOTIFAdapter.
        /// </summary>
        private IMS_OXCNOTIFAdapter cnotifAdapter;

        /// <summary>
        /// The value of the connect pointer which are used to receive notification.
        /// </summary>
        private IntPtr rpcContextForReceive;

        /// <summary>
        /// The value of the receiver context logon handle.
        /// </summary>
        private uint receiverContextLogonHandle;

        /// <summary>
        /// An instance of inboxFolderId.
        /// </summary>
        private ulong inboxFolderId;

        /// <summary>
        /// The message ID of the message which are used to trigger event.
        /// </summary>
        private ulong triggerMessageId;

        /// <summary>
        /// The message ID of the message which are used for locating.
        /// </summary>
        private ulong locatingMessageId;

        /// <summary>
        /// The new created folder ID
        /// </summary>
        private ulong newFolderId;

        /// <summary>
        /// The value of the connect pointer which are used to send notification.
        /// </summary>
        private IntPtr rpcContextForSend;

        /// <summary>
        /// The value of the sender context logon handle.
        /// </summary>
        private uint senderContextLogonHandle;

        /// <summary>
        /// An instance of outboxFolderId.
        /// </summary>
        private ulong outboxFolderId;

        /// <summary>
        /// The list of out handles in the ROP response.
        /// </summary>
        private List<List<uint>> responseSOHs = new List<List<uint>>();

        /// <summary>
        /// The logon response
        /// </summary>
        private RopLogonResponse logonRsp;

        #endregion

        #region Properties

        /// <summary>
        /// Gets a value indicating whether the connection used to receive notification is connected.
        /// </summary>
        protected bool IsReceiverContextConnected
        {
            get
            {
                return this.isReceiverContextConnected;
            }
        }

        /// <summary>
        /// Gets MS-OXCNOTIF adapter
        /// </summary>
        protected IMS_OXCNOTIFAdapter CNOTIFAdapter
        {
            get
            {
                return this.cnotifAdapter;
            }
        }

        /// <summary>
        /// Gets the receiver connect pointer
        /// </summary>
        protected IntPtr RpcContextForReceive
        {
            get
            {
                return this.rpcContextForReceive;
            }
        }

        /// <summary>
        /// Gets the receiver context logon handle
        /// </summary>
        protected uint ReceiverContextLogonHandle
        {
            get
            {
                return this.receiverContextLogonHandle;
            }
        }

        /// <summary>
        /// Gets the inbox folder ID
        /// </summary>
        protected ulong InboxFolderId
        {
            get
            {
                return this.inboxFolderId;
            }
        }

        /// <summary>
        /// Gets the message ID of the message which are used to trigger event
        /// </summary>
        protected ulong TriggerMessageId
        {
            get
            {
                return this.triggerMessageId;
            }
        }

        /// <summary>
        /// Gets the message ID of the message which are used for locating
        /// </summary>
        protected ulong LocatingMessageId
        {
            get
            {
                return this.locatingMessageId;
            }
        }

        /// <summary>
        /// Gets new created folder ID
        /// </summary>
        protected ulong NewFolderId
        {
            get
            {
                return this.newFolderId;
            }
        }
        #endregion

        #region Test Case Initialization and Cleanup

        /// <summary>
        /// Overrides TestClassBase's TestInitialize().
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.cnotifAdapter = Site.GetAdapter<IMS_OXCNOTIFAdapter>();
        }

        /// <summary>
        /// Notification initialize.
        /// </summary>
        protected void NotificationInitialize()
        {
            // Make the sender connection and log on to trigger the notification.
            this.isSenderContextConnected = this.cnotifAdapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.Site.Assert.IsTrue(this.isSenderContextConnected, "If is connected, the ReturnValue of its response is true(success)");
            this.logonRsp = this.cnotifAdapter.Logon();

            this.inboxFolderId = this.logonRsp.FolderIds[4];
            this.outboxFolderId = this.logonRsp.FolderIds[5];
            this.rpcContextForSend = this.cnotifAdapter.RPCContext;
            this.senderContextLogonHandle = this.cnotifAdapter.LogonHandle;

            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);

            uint newMessageHandle;
            uint newFolderHandle;

            // Create the a subfolder in Inbox folder.
            string folderName = Common.GenerateResourceName(Site, "NewSubFolder");
            RopCreateFolderResponse createFolderRsp = this.CreateFolder(inboxHandle, folderName, out newFolderHandle, FolderType.Genericfolder);
            this.newFolderId = createFolderRsp.FolderId;

            // Create the message which are used to trigger event.
            this.CreateMessage(this.inboxFolderId, out newMessageHandle);
            RopSaveChangesMessageResponse saveRsp = this.SaveMessage(newMessageHandle);
            this.triggerMessageId = saveRsp.MessageId;

            // Create the message which are used for locating.
            this.CreateMessage(this.inboxFolderId, out newMessageHandle);
            saveRsp = this.SaveMessage(newMessageHandle);
            this.locatingMessageId = saveRsp.MessageId;

            // Make the receiver connection and log on to register the notification.
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                this.cnotifAdapter.SwitchSessionContext();
            }

            this.isReceiverContextConnected = this.cnotifAdapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.Site.Assert.IsTrue(this.isReceiverContextConnected, "If is connected, the ReturnValue of its response is true(success)");
            this.rpcContextForReceive = this.cnotifAdapter.RPCContext;
            this.cnotifAdapter.Logon();
            this.receiverContextLogonHandle = this.cnotifAdapter.LogonHandle;
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup().
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
            bool transportIsMAPI = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http";
            if (!transportIsMAPI || (transportIsMAPI && Common.IsRequirementEnabled(1340, this.Site)))
            {
                this.CleanUpFolder(this.inboxFolderId);
                this.CleanUpFolder(this.outboxFolderId);

                // Disconnect the sender connect.
                this.isSenderContextConnected = this.cnotifAdapter.DoDisconnect();
                this.Site.Assert.IsTrue(this.isSenderContextConnected, "If is disconnected, the ReturnValue of its response is true(success)");

                // Switch to the receiver connect.
                this.SwitchRPCContextToSender();

                // Set the connect flag to true to make sure the disconnect does execute for the receiver connect.
                this.cnotifAdapter.IsConnected = this.isReceiverContextConnected;

                // Disconnect the receiver connect.
                this.isReceiverContextConnected = this.cnotifAdapter.DoDisconnect();
                this.Site.Assert.IsTrue(this.isReceiverContextConnected, "If is disconnected, the ReturnValue of its response is true(success)");
            }
        }

        #endregion

        #region Event Trigger Methods

        /// <summary>
        /// Trigger new mail event on server.
        /// </summary>
        protected void TriggerNewMailEvent()
        {
            // In order to get the notification from the server, need create event in one context and get from the other one. In the project one is the RPC context used to send notifications and the other is the RPC context used to receive notifications. This method used to change the RPC Context from receiver to sender.
            this.SwitchRPCContextToSender();
            uint messageHandle;
            this.CreateMessage(this.inboxFolderId, out messageHandle);
            PropertyTag[] recipientColumns;
            ModifyRecipientRow[] recipientRows;
            this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);
            this.ModifyRecipients(messageHandle, recipientColumns, recipientRows);
            uint inboxTableHandle;
            this.OpenFolder(this.InboxFolderId, out inboxTableHandle);
            uint contentTableHandle;

            // The message count before the new created message is submitted.
            uint msgCount = this.GetContentsTable(inboxTableHandle, out contentTableHandle, false).RowCount;
            this.SubmitMessage(messageHandle);

            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
            RopGetContentsTableResponse getContentsTableResponse;

            // Do a loop to make sure the new created message is submitted completely.
            while (retryCount > 0)
            {
                getContentsTableResponse = this.GetContentsTable(inboxTableHandle, out contentTableHandle, false);

                // The message is sent to the current user itself, so the message count will raise 2.
                if (getContentsTableResponse.RowCount == msgCount + 2)
                {
                    break;
                }

                System.Threading.Thread.Sleep(sleepTime);
                retryCount--;
            }

            Site.Assert.IsTrue(retryCount > 0, "The message should be submitted to server successfully by sending RopSubmitMessage request. Increase the retry count value defined in RetryCount property in ptfconfig file, and try again.");

            this.ReleaseHandle(messageHandle);

            // In order to get the notification from the server, need create event in one context and get from the other one. In the project one is the RPC context used to send notifications and the other is the RPC context used to receive notifications. This method used to change the RPC Context from sender to receiver.
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger an object create event.
        /// </summary>
        /// <returns>The created folder ID</returns>
        protected ulong TriggerObjectCreatedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);
            uint newFolderHandle;
            string folderName = Common.GenerateResourceName(Site, "NewCreateSubFolder");
            RopCreateFolderResponse createFolderResponse = this.CreateFolder(inboxHandle, folderName, out newFolderHandle, FolderType.Genericfolder);
            this.SwitchRPCContextToReceiver();
            return createFolderResponse.FolderId;
        }

        /// <summary>
        /// Trigger object deleted event.
        /// </summary>
        protected void TriggerObjectDeletedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);
            this.DeleteFolder(this.newFolderId, inboxHandle);
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger object moved folder event.
        /// </summary>
        protected void TriggerObjectMovedFolderEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);
            uint outboxHandle;
            this.OpenFolder(this.outboxFolderId, out outboxHandle);
            string moveName = Common.GenerateResourceName(Site, "MoveFolder");
            RopMoveFolderResponse copyFolderRsp = this.MoveFolder(this.newFolderId, moveName, inboxHandle, outboxHandle);
            Site.Assume.AreEqual<byte>(0, copyFolderRsp.PartialCompletion, "Moved folder completed.");
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger object move message event.
        /// </summary>
        protected void TriggerObjectMessageMoveEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);
            uint outboxHandle;
            this.OpenFolder(this.outboxFolderId, out outboxHandle);
            RopMoveCopyMessagesResponse moveRsp = this.MoveCopyMessage(inboxHandle, outboxHandle, this.triggerMessageId, false);
            Site.Assume.AreEqual<byte>(0, moveRsp.PartialCompletion, "Move message completed.");
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger object copied event.
        /// </summary>
        protected void TriggerObjectCopiedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxHandle;
            this.OpenFolder(this.inboxFolderId, out inboxHandle);
            uint outboxHandle;
            this.OpenFolder(this.outboxFolderId, out outboxHandle);
            RopMoveCopyMessagesResponse moveRsp = this.MoveCopyMessage(inboxHandle, outboxHandle, this.triggerMessageId, true);
            Site.Assume.AreEqual<byte>(0, moveRsp.PartialCompletion, "Message copied.");
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger object modified event.
        /// </summary>
        protected void TriggerObjectModifiedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxTableHandle;
            this.OpenFolder(this.inboxFolderId, out inboxTableHandle);
            uint messageHandle;

            this.OpenMessage(inboxTableHandle, this.inboxFolderId, this.triggerMessageId, out messageHandle);

            PropertyTag[] recipientColumns;
            ModifyRecipientRow[] recipientRows;
            this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);
            this.ModifyRecipients(messageHandle, recipientColumns, recipientRows);

            RopSaveChangesMessageResponse save2Rsp = this.SaveMessage(messageHandle);
            Site.Assume.AreEqual<uint>(0, save2Rsp.ReturnValue, "RopSaveChangesMessage operation performs successfully.");
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger search completed event.
        /// </summary>
        protected void TriggerSearchCompletedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxTableHandle;
            this.OpenFolder(this.inboxFolderId, out inboxTableHandle);
            uint searchFolderHandle;
            string folderName = Common.GenerateResourceName(Site, "NewSearchFolder");
            this.CreateFolder(inboxTableHandle, folderName, out searchFolderHandle, FolderType.Searchfolder);

            byte[] restrictionData = new byte[5];
            ushort pidTagMessageClassId = PropertyTags.All[PropertyNames.PidTagMessageClass].PropertyId;
            ushort typeOfPidTagMessageClass = PropertyTags.All[PropertyNames.PidTagMessageClass].PropertyType;
            restrictionData[0] = (byte)Restrictions.ExistRestriction;
            Array.Copy(BitConverter.GetBytes(typeOfPidTagMessageClass), 0, restrictionData, 1, sizeof(ushort));
            Array.Copy(BitConverter.GetBytes(pidTagMessageClassId), 0, restrictionData, 3, sizeof(ushort));

            RopSetSearchCriteriaResponse searchRsp = this.SetSearchCriteria(searchFolderHandle, new ulong[] { this.inboxFolderId }, restrictionData);

            Site.Assume.AreEqual<uint>(0, searchRsp.ReturnValue, "RopSetSearchCriteria operation performs successfully.");
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger table row added event.
        /// </summary>
        /// <returns>New created message ID</returns>
        protected ulong TriggerTableRowAddedEvent()
        {
            this.SwitchRPCContextToSender();
            uint messageHandle;
            this.CreateMessage(this.inboxFolderId, out messageHandle);
            RopSaveChangesMessageResponse saveRsp = this.SaveMessage(messageHandle);
            Site.Assume.AreEqual<uint>(0, saveRsp.ReturnValue, "RopSaveChangesMessage operation performs successfully.");
            this.SwitchRPCContextToReceiver();
            return saveRsp.MessageId;
        }

        /// <summary>
        /// Trigger table row deleted event.
        /// </summary>
        protected void TriggerTableRowDeletedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxTableHandle;
            this.OpenFolder(this.inboxFolderId, out inboxTableHandle);
            this.DeleteMessage(inboxTableHandle, this.triggerMessageId);
            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger table row changed event.
        /// </summary>
        protected void TriggerTableChangedEvent()
        {
            this.SwitchRPCContextToSender();
            uint inboxTableHandle;
            this.OpenFolder(this.inboxFolderId, out inboxTableHandle);
            for (int i = 0; i < 70; i++)
            {
                uint messageHandle;
                this.CreateMessage(this.inboxFolderId, out messageHandle);
                RopSaveChangesMessageResponse saveResponse = this.SaveMessage(messageHandle);
                this.ReleaseHandle(messageHandle);
                this.DeleteMessage(inboxTableHandle, saveResponse.MessageId);
            }

            this.SwitchRPCContextToReceiver();
        }

        /// <summary>
        /// Trigger table row modified event.
        /// </summary>
        protected void TriggerTableRowModifiedEvent()
        {
            this.TriggerObjectModifiedEvent();
        }
        #endregion

        #region Protected Methods

        /// <summary>
        /// Check whether support MS-OXCMAPIHTTP transport.
        /// </summary>
        protected void CheckWhetherSupportMAPIHTTP()
        {
            if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http") && (!Common.IsRequirementEnabled(1340, this.Site)))
            {
                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
        }

        /// <summary>
        /// Opens an existing folder in a mailbox.
        /// </summary>
        /// <param name="folderId">The Id of the folder which want to open.</param>
        /// <param name="tableHandle">Return the table handle that will be used in other operation.</param>
        /// <returns>The server response.</returns>
        protected RopOpenFolderResponse OpenFolder(ulong folderId, out uint tableHandle)
        {
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01,which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = 1;

            // Set FolderId to the 4th of logonResponse,this folder is to be opened.
            openFolderRequest.FolderId = folderId;

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                openFolderRequest,
                this.cnotifAdapter.LogonHandle,
                out this.responseSOHs);
            openFolderResponse = (RopOpenFolderResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            tableHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];
            return openFolderResponse;
        }

        /// <summary>
        /// Sets the properties visible on a table.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="propertyTags">Specifies the property values that are visible in table rows.</param>
        /// <returns>The server response.</returns>
        protected RopSetColumnsResponse SetColumns(uint tableHandle, PropertyTag[] propertyTags)
        {
            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;
            setColumnsRequest.InputHandleIndex = 0;
            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                setColumnsRequest,
                tableHandle,
                out this.responseSOHs);
            setColumnsResponse = (RopSetColumnsResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            return setColumnsResponse;
        }

        /// <summary>
        /// Retrieves rows from a table.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="rowCount">Specifies the number of requested rows.</param>
        /// <returns>The server response.</returns>
        protected RopQueryRowsResponse QueryRows(uint tableHandle, ushort rowCount)
        {
            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.5.4.1.
            queryRowsRequest.InputHandleIndex = 0;

            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;

            // Set ForwardRead to 0xff(TRUE), which specifies the direction to read rows(forwards),
            // as specified in [MS-OXCROPS] section 2.2.5.4.1.
            queryRowsRequest.ForwardRead = 1;

            // Set RowCount to 0x0032, which the number of requested rows,
            // as specified in [MS-OXCROPS] section 2.2.5.4.1.
            queryRowsRequest.RowCount = rowCount;

            IList<IDeserializable> ropResponses = this.cnotifAdapter.Process(
                queryRowsRequest,
                tableHandle,
                out this.responseSOHs);
            queryRowsResponse = (RopQueryRowsResponse)ropResponses[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                queryRowsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            return queryRowsResponse;
        }

        /// <summary>
        /// Gets the content table of a container.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="contentTableHandle">Content Table Handle.</param>
        /// <param name="disableNotification">Set to true if want to disable table notification</param>
        /// <returns>The server response</returns>
        protected RopGetContentsTableResponse GetContentsTable(uint tableHandle, out uint contentTableHandle, bool disableNotification)
        {
            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getContentsTableRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01,which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            getContentsTableRequest.OutputHandleIndex = 1;

            getContentsTableRequest.TableFlags = (byte)(disableNotification ? FolderTableFlags.NoNotifications : FolderTableFlags.None);

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                getContentsTableRequest,
                tableHandle,
                out this.responseSOHs);
            getContentsTableResponse = (RopGetContentsTableResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                getContentsTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            contentTableHandle = this.responseSOHs[0][getContentsTableResponse.OutputHandleIndex];
            return getContentsTableResponse;
        }

        /// <summary>
        /// Marks the current cursor position in a table.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <returns>The server response.</returns>
        protected RopCreateBookmarkResponse CreateBookmark(uint tableHandle)
        {
            RopCreateBookmarkRequest createBookmarkRequest;
            RopCreateBookmarkResponse createBookmarkResponse;

            createBookmarkRequest.RopId = (byte)RopId.RopCreateBookmark;
            createBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.11.1.
            createBookmarkRequest.InputHandleIndex = 0;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                createBookmarkRequest,
                tableHandle,
                out this.responseSOHs);
            createBookmarkResponse = (RopCreateBookmarkResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                createBookmarkResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return createBookmarkResponse;
        }

        /// <summary>
        /// Moves the cursor to a location specified relative to a user-defined bookmark.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="bookmark">Specifies the origin for the seek operation.</param>
        /// <param name="rowCount">Specifies the direction and the number of rows to seek.</param>
        /// <returns>The server response.</returns>
        protected RopSeekRowBookmarkResponse SeekRowBookmark(uint tableHandle, byte[] bookmark, int rowCount)
        {
            RopSeekRowBookmarkRequest seekRowBookmarkRequest;
            RopSeekRowBookmarkResponse seekRowBookmarkResponse;

            seekRowBookmarkRequest.RopId = (byte)RopId.RopSeekRowBookmark;
            seekRowBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.9.1.
            seekRowBookmarkRequest.InputHandleIndex = 0;

            // Set BookmarkSize, which specifies the size of the Bookmark field,
            // as specified in [MS-OXCROPS] section 2.2.5.9.1.
            seekRowBookmarkRequest.BookmarkSize = (ushort)bookmark.Length;

            // Set Bookmark, which specifies the origin for the seek operation,
            // as specified in [MS-OXCROPS] section 2.2.5.9.1.
            seekRowBookmarkRequest.Bookmark = bookmark;

            // Set RowCount, which specifies the direction and the number of rows to seek,
            // as specified in [MS-OXCROPS] section 2.2.5.9.1.
            seekRowBookmarkRequest.RowCount = rowCount;

            // Set WantRowMovedCount to 0xff(TRUE), which specifies the server returns the actual number of rows sought
            // in the response,as specified in [MS-OXCROPS] section 2.2.5.9.1.
            seekRowBookmarkRequest.WantRowMovedCount = 0xff;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                seekRowBookmarkRequest,
                tableHandle,
                out this.responseSOHs);
            seekRowBookmarkResponse = (RopSeekRowBookmarkResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                seekRowBookmarkResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return seekRowBookmarkResponse;
        }

        /// <summary>
        /// Moves the cursor to an approximate position in a table.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="numerator">Represents the numerator of the fraction identifying the table position to seek to.</param>
        /// <param name="denominator">Represents the denominator of the fraction identifying the table position to seek to.</param>
        /// <returns>The server response.</returns>
        protected RopSeekRowFractionalResponse SeekRowFractional(uint tableHandle, uint numerator, uint denominator)
        {
            RopSeekRowFractionalRequest seekRowFractionalRequest;
            RopSeekRowFractionalResponse seekRowFractionalResponse;

            seekRowFractionalRequest.RopId = (byte)RopId.RopSeekRowFractional;
            seekRowFractionalRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.10.1.
            seekRowFractionalRequest.InputHandleIndex = 0;

            // Set Numerator, which represents the numerator of the fraction identifying the table position to seek to,
            // as specified in [MS-OXCROPS] section 2.2.5.10.1.
            seekRowFractionalRequest.Numerator = numerator;

            // Set Denominator, which represents the denominator of the fraction identifying the table position to seek to,
            // as specified in [MS-OXCROPS] section 2.2.5.10.1.
            seekRowFractionalRequest.Denominator = denominator;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                seekRowFractionalRequest,
                tableHandle,
                out this.responseSOHs);
            seekRowFractionalResponse = (RopSeekRowFractionalResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                seekRowFractionalResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return seekRowFractionalResponse;
        }

        /// <summary>
        /// Moves the cursor to a specific position in a table.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="rowCount">Specifies the direction and the number of rows to seek.</param>
        /// <returns>The server response.</returns>
        protected RopSeekRowResponse SeekRow(uint tableHandle, int rowCount)
        {
            RopSeekRowRequest seekRowRequest;
            RopSeekRowResponse seekRowResponse;

            seekRowRequest.RopId = (byte)RopId.RopSeekRow;
            seekRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.8.1.
            seekRowRequest.InputHandleIndex = 0;

            seekRowRequest.Origin = (byte)Origin.Beginning;

            // Set RowCount, which specifies the direction and the number of rows to seek,
            // as specified in [MS-OXCROPS] section 2.2.5.8.1.
            seekRowRequest.RowCount = rowCount;

            // Set WantRowMovedCount to 0xff(TRUE),which specifies the server returns the actual number of rows moved
            // in the response,as specified in [MS-OXCROPS] section 2.2.5.8.1.
            seekRowRequest.WantRowMovedCount = 0xff;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                seekRowRequest,
                tableHandle,
                out this.responseSOHs);
            seekRowResponse = (RopSeekRowResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                seekRowResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return seekRowResponse;
        }

        /// <summary>
        /// This ROP expands a categorized row, see [MS-OXCTABL] section 2.2.2.17.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="maxRowCount">Maximum number of expanded rows to return data for.</param>
        /// <param name="categoryId">specifies the category to be expanded</param>
        /// <returns>The server response.</returns>
        protected RopExpandRowResponse ExpandRow(uint tableHandle, ushort maxRowCount, ulong categoryId)
        {
            RopExpandRowRequest expandRowRequest;
            RopExpandRowResponse expandRowResponse;

            expandRowRequest.RopId = (byte)RopId.RopExpandRow;
            expandRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored,as specified in [MS-OXCROPS] section 2.2.5.16.1.
            expandRowRequest.InputHandleIndex = 0;

            // Set MaxRowCount, which specifies the maximum number of expanded rows to return data for,
            // as specified in [MS-OXCROPS] section 2.2.5.16.1.
            expandRowRequest.MaxRowCount = maxRowCount;

            expandRowRequest.CategoryId = categoryId;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                expandRowRequest,
                tableHandle,
                out this.responseSOHs);
            expandRowResponse = (RopExpandRowResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                expandRowResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return expandRowResponse;
        }

        /// <summary>
        /// Call the RopRestrict ROP ([MS-OXCROPS] section 2.2.5.3) to establish a restriction on a table. 
        /// </summary>
        /// <param name="tableHandle">The table handle for this operation.</param>
        /// <param name="restriction">The restriction data</param>
        /// <returns>The server response.</returns>
        protected RopRestrictResponse RestrictTable(uint tableHandle, byte[] restriction)
        {
            RopRestrictRequest restrictRequest;
            RopRestrictResponse restrictResponse;

            restrictRequest.RopId = (byte)RopId.RopRestrict;
            restrictRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.5.3.1.
            restrictRequest.InputHandleIndex = 0;

            restrictRequest.RestrictFlags = (byte)AsynchronousFlags.TblAsync;

            // Set RestrictionDataSize, which specifies the length of the RestrictionData field,
            // as specified in [MS-OXCROPS] section 2.2.5.3.1.
            restrictRequest.RestrictionDataSize = (ushort)restriction.Length;

            // Set RestrictionData to null, which specifies there is no filter for limiting the view of a table to particular set of rows,
            // as specified in [MS-OXCDATA] section 2.12.
            restrictRequest.RestrictionData = restriction;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                restrictRequest,
                tableHandle,
                out this.responseSOHs);
            restrictResponse = (RopRestrictResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                restrictResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return restrictResponse;
        }

        /// <summary>
        /// Moves the cursor to a row in a table that matches specific search criteria.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <param name="restrictionData">The restriction specifies the filter for this operation.</param>
        /// <returns>The server response.</returns>
        protected RopFindRowResponse FindRow(uint tableHandle, byte[] restrictionData)
        {
            RopFindRowRequest findRowRequest;
            RopFindRowResponse findRowResponse;

            findRowRequest.RopId = (byte)RopId.RopFindRow;
            findRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.13.1.
            findRowRequest.InputHandleIndex = 0;

            findRowRequest.FindRowFlags = (byte)FindRowFlags.Forwards;

            // Set RestrictionDataSize, which specifies the length of the RestrictionData field,
            // as specified in [MS-OXCROPS] section 2.2.5.13.1.
            findRowRequest.RestrictionDataSize = (ushort)restrictionData.Length;
            findRowRequest.RestrictionData = restrictionData;
            findRowRequest.Origin = (byte)Origin.Beginning;

            // Set BookmarkSize, which specifies the size of the Bookmark field,
            // as specified in [MS-OXCROPS] section 2.2.5.13.1.
            findRowRequest.BookmarkSize = 0;

            // Set Bookmark, which specifies the bookmark to use as the origin,
            // as specified in [MS-OXCROPS] section 2.2.5.13.1.
            findRowRequest.Bookmark = null;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                findRowRequest,
                tableHandle,
                out this.responseSOHs);
            findRowResponse = (RopFindRowResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                findRowResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return findRowResponse;
        }

        /// <summary>
        /// Gets a list of columns in a table.  
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <returns>The server response.</returns>
        protected RopQueryColumnsAllResponse QueryColumnsAll(uint tableHandle)
        {
            RopQueryColumnsAllRequest queryColumnsAllRequest;
            RopQueryColumnsAllResponse queryColumnsAllResponse;

            queryColumnsAllRequest.RopId = (byte)RopId.RopQueryColumnsAll;
            queryColumnsAllRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.5.12.1.
            queryColumnsAllRequest.InputHandleIndex = 0;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                queryColumnsAllRequest,
                tableHandle,
                out this.responseSOHs);
            queryColumnsAllResponse = (RopQueryColumnsAllResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                queryColumnsAllResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return queryColumnsAllResponse;
        }

        /// <summary>
        /// Gets the cursor position.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <returns>The server response.</returns>
        protected RopQueryPositionResponse QueryPosition(uint tableHandle)
        {
            RopQueryPositionRequest queryPositionRequest;
            RopQueryPositionResponse queryPositionResponse;

            queryPositionRequest.RopId = (byte)RopId.RopQueryPosition;
            queryPositionRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored,
            // as specified in [MS-OXCROPS] section 2.2.5.7.1.
            queryPositionRequest.InputHandleIndex = 0;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                queryPositionRequest,
                tableHandle,
                out this.responseSOHs);
            queryPositionResponse = (RopQueryPositionResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                queryPositionResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return queryPositionResponse;
        }

        /// <summary>
        /// Resets a table to its original state, see [MS-OXCTABL] section 2.2.2.16.
        /// </summary>
        /// <param name="tableHandle">Represents the table handle for this operation.</param>
        /// <returns>The server response.</returns>
        protected RopResetTableResponse ResetTable(uint tableHandle)
        {
            RopResetTableRequest resetTableRequest;
            RopResetTableResponse resetTableResponse;

            resetTableRequest.RopId = (byte)RopId.RopResetTable;
            resetTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.5.15.1.
            resetTableRequest.InputHandleIndex = 0;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                resetTableRequest,
                tableHandle,
                out this.responseSOHs);
            resetTableResponse = (RopResetTableResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                resetTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return resetTableResponse;
        }

        /// <summary>
        /// Open a message in the folder
        /// </summary>
        /// <param name="folderHandle">The folder handle</param>
        /// <param name="folderId">The folder ID</param>
        /// <param name="messageId">The message ID</param>
        /// <param name="messageHandle">The message handle</param>
        /// <returns>The server response</returns>
        protected RopOpenMessageResponse OpenMessage(uint folderHandle, ulong folderId, ulong messageId, out uint messageHandle)
        {
            RopOpenMessageRequest openMessageRequest;
            RopOpenMessageResponse openMessageResponse;

            openMessageRequest.RopId = (byte)RopId.RopOpenMessage;
            openMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.6.1.1.
            openMessageRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored, as specified in [MS-OXCROPS] section 2.2.6.1.1.
            openMessageRequest.OutputHandleIndex = 1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used,
            // as specified in [MS-OXCROPS] section 2.2.6.1.1.
            openMessageRequest.CodePageId = 0x0FFF;

            // Set FolderId to that of created folder,which identifies the parent folder,
            // as specified in [MS-OXCROPS] section 2.2.6.1.1.
            openMessageRequest.FolderId = folderId;

            openMessageRequest.OpenModeFlags = (byte)MessageOpenModeFlags.ReadWrite;

            // Set MessageId to that of created message,which identifies the message to be opened,
            // as specified in [MS-OXCROPS] section 2.2.6.1.1.
            openMessageRequest.MessageId = messageId;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                openMessageRequest,
                folderHandle,
                out this.responseSOHs);
            openMessageResponse = (RopOpenMessageResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                openMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            messageHandle = this.responseSOHs[0][openMessageRequest.OutputHandleIndex];
            return openMessageResponse;
        }

        /// <summary>
        /// This ROP gets specific properties of a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">This field specifies the properties requested.</param>
        /// <returns>The structure of RopGetPropertiesSpecificResponse.</returns>
        protected RopGetPropertiesSpecificResponse GetPropertiesSpecific(uint objHandle, PropertyTag[] propertyTags)
        {
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest;
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;

            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.6.1.1.
            getPropertiesSpecificRequest.InputHandleIndex = 0;

            // The maximum size allowed for a property value returned
            getPropertiesSpecificRequest.PropertySizeLimit = 0;

            // Set the value to make the return string properties in Unicode.
            getPropertiesSpecificRequest.WantUnicode = 0x01;

            // Set the value the properties required.
            if (propertyTags != null)
            {
                getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                getPropertiesSpecificRequest.PropertyTagCount = 0x00;
            }

            getPropertiesSpecificRequest.PropertyTags = propertyTags;
            IList<IDeserializable> response = this.cnotifAdapter.Process(
              getPropertiesSpecificRequest,
              objHandle,
              out this.responseSOHs);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                getPropertiesSpecificResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return getPropertiesSpecificResponse;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Delete all message and subfolder of specified folder
        /// </summary>
        /// <param name="folderId">The folder ID</param>
        private void CleanUpFolder(ulong folderId)
        {
            uint tableHandle, contentTableHanlde;
            this.OpenFolder(folderId, out tableHandle);

            // Indicates the current message count in the specified folder.
            uint rowCount;

            // The times to try cleaning up the specified folder.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
            do
            {
                this.HardDeleteMessagesAndSubfolders(tableHandle);
                System.Threading.Thread.Sleep(sleepTime);
                rowCount = this.GetContentsTable(tableHandle, out contentTableHanlde, false).RowCount;
                retryCount--;
            }
            while (rowCount != 0 && retryCount >= 0);
            Site.Assert.AreEqual<uint>(0, rowCount, "The specified folder should be cleaned up.");
        }

        /// <summary>
        /// This method creates Sample RecipientColumns and Sample RecipientRows.
        /// </summary>
        /// <param name="recipientColumns">Sample RecipientColumns</param>
        /// <param name="recipientRows">Sample RecipientRows</param>
        private void CreateSampleRecipientColumnsAndRecipientRows(out PropertyTag[] recipientColumns, out ModifyRecipientRow[] recipientRows)
        {
            // Step 1: Create Sample RecipientColumns.

            // The following sample data is from MS-OXCMSG 4.7.1.
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            sampleRecipientColumns[0] = PropertyTags.All[PropertyNames.PidTagObjectType];
            sampleRecipientColumns[1] = PropertyTags.All[PropertyNames.PidTagDisplayType];
            sampleRecipientColumns[2] = PropertyTags.All[PropertyNames.PidTagAddressBookDisplayNamePrintable];
            sampleRecipientColumns[3] = PropertyTags.All[PropertyNames.PidTagSmtpAddress];
            sampleRecipientColumns[4] = PropertyTags.All[PropertyNames.PidTagSendInternetEncoding];
            sampleRecipientColumns[5] = PropertyTags.All[PropertyNames.PidTagDisplayTypeEx];
            sampleRecipientColumns[6] = PropertyTags.All[PropertyNames.PidTagRecipientDisplayName];
            sampleRecipientColumns[7] = PropertyTags.All[PropertyNames.PidTagRecipientFlags];
            sampleRecipientColumns[8] = PropertyTags.All[PropertyNames.PidTagRecipientTrackStatus];
            sampleRecipientColumns[9] = PropertyTags.All[PropertyNames.PidTagRecipientResourceState];
            sampleRecipientColumns[10] = PropertyTags.All[PropertyNames.PidTagRecipientOrder];
            sampleRecipientColumns[11] = PropertyTags.All[PropertyNames.PidTagRecipientEntryId];
            recipientColumns = sampleRecipientColumns;

            // Step 2:Configure a StandardPropertyRow: propertyRow.
            PropertyValue[] propertyValueArray = new PropertyValue[12];

            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            // PidTagObjectType
            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006);

            // PidTagDisplayType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000);

            // PidTa7BitDisplayName
            propertyValueArray[2].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0");

            // PidTagSmtpAddress
            propertyValueArray[3].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0");

            // PidTagSendInternetEncoding
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000);

            // PidTagDisplayTypeEx
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000);

            // PidTagRecipientDisplayName
            propertyValueArray[6].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0");

            // PidTagRecipientFlags
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001);

            // PidTagRecipientTrackStatus
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientResourceState
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientOrder
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000);

            // The following sample data(0x007c and the subsequent 124(0x7c) binary)
            // is copied from section 4.7.1 of MS-OXCMSG
            byte[] sampleData = 
            {                     
                                  0x7c, 0x00, 0x00, 0x00, 0x00, 0x00, 0xdc, 0xa7, 0x40, 0xc8,
                                  0xc0, 0x42, 0x10, 0x1a, 0xb4, 0xb9, 0x08, 0x00, 0x2b, 0x2f,
                                  0xe1, 0x82, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                  0x2f, 0x6f, 0x3d, 0x46, 0x69, 0x72, 0x73, 0x74, 0x20, 0x4f,
                                  0x72, 0x67, 0x61, 0x6e, 0x69, 0x7a, 0x61, 0x74, 0x69, 0x6f,
                                  0x6e, 0x2f, 0x6f, 0x75, 0x3d, 0x45, 0x78, 0x63, 0x68, 0x61,
                                  0x6e, 0x67, 0x65, 0x20, 0x41, 0x64, 0x6d, 0x69, 0x6e, 0x69,
                                  0x73, 0x74, 0x72, 0x61, 0x74, 0x69, 0x76, 0x65, 0x20, 0x47,
                                  0x72, 0x6f, 0x75, 0x70, 0x20, 0x28, 0x46, 0x59, 0x44, 0x49,
                                  0x42, 0x4f, 0x48, 0x46, 0x32, 0x33, 0x53, 0x50, 0x44, 0x4c,
                                  0x54, 0x29, 0x2f, 0x63, 0x6e, 0x3d, 0x52, 0x65, 0x63, 0x69,
                                  0x70, 0x69, 0x65, 0x6e, 0x74, 0x73, 0x2f, 0x63, 0x6e, 0x3d,
                                  0x75, 0x73, 0x65, 0x72, 0x32, 0x00 
            };

            // PidTagRecipientEntryId
            propertyValueArray[11].Value = sampleData;

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = (byte)PropertyRowFlag.FlaggedPropertyRow,
                PropertyValues = propertyValues
            };

            // For propertyRow.Flag
            int tempLengthForTest = 1;
            foreach (PropertyValue pv in propertyRow.PropertyValues)
            {
                tempLengthForTest = tempLengthForTest + pv.Value.Length;
            }

            // Step 3:Create Sample RecipientRows.
            RecipientRow recipientRow = new RecipientRow
            {
                // 0101 1001 0000 0110 S,D, Type=SMTP,I,U,E
                RecipientFlags = (ushort)(RecipientFlags.SMTP | RecipientFlags.S | RecipientFlags.D | RecipientFlags.I | RecipientFlags.U | RecipientFlags.E),

                // Set DisplayName,which specifies the Email Address of the recipient,as specified in [MS-OXCDATA]
                // section 2.9.3.2,this field is present because D is Set.
                DisplayName = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0"),

                // Set EmailAddress,which specifies the Email Address of the recipient,
                // as specified in [MS-OXCDATA] section 2.9.3.2.
                EmailAddress =
                    Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" +
                                              Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0"),

                // Set SimpleDisplayName,which specifies the Email Address of the recipient,
                // as specified in [MS-OXCDATA] section 2.9.3.2.
                SimpleDisplayName =
                    Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "\0"),

                // Set RecipientColumnCount,which specifies the number of columns from the RecipientColumns field
                // that are included in RecipientProperties,as specified in [MS-OXCDATA] section 2.9.3.2.
                RecipientColumnCount = 0x000C,
                RecipientProperties = propertyRow
            };

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                // Set RowId,which value specifies the ID of the recipient,as specified in [MS-OXCROPS] section 2.2.6.5.1.1.
                RowId = 0,
                RecipientType = (byte)RecipientType.PrimaryRecipient,

                // Set RecipientRowSize,which specifies the size of the RecipientRow field,
                // as specified in [MS-OXCROPS] section 2.2.6.5.1.1.
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };

            ModifyRecipientRow[] sampleModifyRecipientRows = new ModifyRecipientRow[1];
            sampleModifyRecipientRows[0] = modifyRecipientRow;
            recipientRows = sampleModifyRecipientRows;
        }

        /// <summary>
        /// Create a new folder on the server
        /// </summary>
        /// <param name="parentFolderHandle">The parent folder handle</param>
        /// <param name="name">The new folder name</param>
        /// <param name="newFolderHandle">Return the new folder handle</param>
        /// <param name="folderType">Specifies the folder type</param>
        /// <returns>The server response</returns>
        private RopCreateFolderResponse CreateFolder(uint parentFolderHandle, string name, out uint newFolderHandle, FolderType folderType)
        {
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createFolderRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01,which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = 1;

            createFolderRequest.FolderType = (byte)folderType;

            // Set UseUnicodeStrings to 0x0(FALSE),which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = 0;

            // Set OpenExisting to 0xFF,which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = 1;

            // Set Reserved to 0x0,this field is reserved and MUST be set to 0.
            createFolderRequest.Reserved = 0;

            // Set DisplayName,which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(name + "\0");

            // Set Comment,which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(name + "\0");

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                createFolderRequest,
                parentFolderHandle,
                out this.responseSOHs);
            createFolderResponse = (RopCreateFolderResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            newFolderHandle = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];
            return createFolderResponse;
        }

        /// <summary>
        /// Move or copy a message on the server
        /// </summary>
        /// <param name="fromFolderHandle">The source folder handle</param>
        /// <param name="toFolderHandle">The destination folder handle</param>
        /// <param name="messageId">The message ID which to operate</param>
        /// <param name="wantCopy">Set to true if want to copy the message, false to move</param>
        /// <returns>The server response</returns>
        private RopMoveCopyMessagesResponse MoveCopyMessage(uint fromFolderHandle, uint toFolderHandle, ulong messageId, bool wantCopy)
        {
            RopMoveCopyMessagesRequest moveCopyMessagesRequest;
            RopMoveCopyMessagesResponse moveCopyMessagesResponse;

            moveCopyMessagesRequest.RopId = (byte)RopId.RopMoveCopyMessages;
            moveCopyMessagesRequest.LogonId = TestSuiteBase.LogonId;

            // Set SourceHandleIndex to 0x0,which specifies the location in the Server object handle table
            // where the handle for the source Server object is stored.
            moveCopyMessagesRequest.SourceHandleIndex = 0;

            // Set DestHandleIndex to 0x01,which specifies the location in the Server object handle table
            // where the handle for the destination Server object is stored.
            moveCopyMessagesRequest.DestHandleIndex = 1;
            ulong[] messageIds = new ulong[] { messageId };

            // Set MessageIdCount to the length of messageIds,which specifies the size of the MessageIds field.
            moveCopyMessagesRequest.MessageIdCount = (ushort)messageIds.Length;

            // Set MessageIds to messageIds,which specify which messages to move or copy.
            moveCopyMessagesRequest.MessageIds = messageIds;

            // Set WantAsynchronous to 0x00(FALSE),which specifies the operation is to be executed Synchronously.
            moveCopyMessagesRequest.WantAsynchronous = 0;

            // Set WantCopy to 0xFF(TRUE),which specifies the operation is a copy.
            moveCopyMessagesRequest.WantCopy = (byte)(wantCopy ? 0xff : 0);

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                moveCopyMessagesRequest,
                new uint[] { fromFolderHandle, toFolderHandle },
                out this.responseSOHs);
            moveCopyMessagesResponse = (RopMoveCopyMessagesResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                moveCopyMessagesResponse.ReturnValue,
               "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return moveCopyMessagesResponse;
        }

        /// <summary>
        /// Release the handle resource.
        /// </summary>
        /// <param name="handle">The handle need to be released.</param>
        private void ReleaseHandle(uint handle)
        {
            RopReleaseRequest releaseRequest;
            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;
            releaseRequest.InputHandleIndex = 0;

            this.cnotifAdapter.Process(
                releaseRequest,
                handle,
                out this.responseSOHs);
        }

        /// <summary>
        /// Hard delete all messages and subfolder in specified folder
        /// </summary>
        /// <param name="folderHandle">The folder handle</param>
        private void HardDeleteMessagesAndSubfolders(uint folderHandle)
        {
            RopHardDeleteMessagesAndSubfoldersRequest deleteRequest;
            RopHardDeleteMessagesAndSubfoldersResponse deleteResponse;

            deleteRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
            deleteRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            deleteRequest.InputHandleIndex = 0;

            // Set WantAsynchronous to 0x00(FALSE),which specifies the operation is to be executed Synchronously.
            deleteRequest.WantAsynchronous = 0;

            // Set to true to delete associated information(FAI)
            deleteRequest.WantDeleteAssociated = 0xFF;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                deleteRequest,
                folderHandle,
                out this.responseSOHs);
            deleteResponse = (RopHardDeleteMessagesAndSubfoldersResponse)response[0];

            // Return value 1125 means ecNoDelSubmitMsg, Deleting a message that has been submitted for sending is not permitted.
            bool returnValue = (deleteResponse.ReturnValue == 0) || (deleteResponse.ReturnValue == 1125);

            Site.Assert.IsTrue(
                returnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
        }

        /// <summary>
        /// Delete specified message in the folder
        /// </summary>
        /// <param name="folderHandle">The folder handle</param>
        /// <param name="messageId">The message ID</param>
        private void DeleteMessage(uint folderHandle, ulong messageId)
        {
            RopDeleteMessagesRequest deleteMessagesRequest;
            RopDeleteMessagesResponse deleteMessagesResponse;

            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            deleteMessagesRequest.InputHandleIndex = 0;

            // Set WantAsynchronous to 0x00(FALSE),which specifies the operation is to be executed Synchronously.
            deleteMessagesRequest.WantAsynchronous = 0;

            // Set NotifyNonRead to 0x00,which specifies the server doesn't generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = 0;
            ulong[] messageIds = new ulong[] { messageId };
            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                deleteMessagesRequest,
                folderHandle,
                out this.responseSOHs);
            deleteMessagesResponse = (RopDeleteMessagesResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                deleteMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
        }

        /// <summary>
        /// Sets the search criteria on a search folder
        /// </summary>
        /// <param name="searchFolderHandle">The folder handle</param>
        /// <param name="folderIds">The folders which will be searched</param>
        /// <param name="restrictionData">The restriction data</param>
        /// <returns>The server response</returns>
        private RopSetSearchCriteriaResponse SetSearchCriteria(uint searchFolderHandle, ulong[] folderIds, byte[] restrictionData)
        {
            RopSetSearchCriteriaRequest setSearchCriteriaRequest;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse;

            setSearchCriteriaRequest.RopId = (byte)RopId.RopSetSearchCriteria;
            setSearchCriteriaRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            setSearchCriteriaRequest.InputHandleIndex = 0;

            // Set RestrictionDataSize to 0x0005,which specifies the length of the RestrictionData field.
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)restrictionData.Length;

            setSearchCriteriaRequest.RestrictionData = restrictionData;

            // Set FolderIdCount to the length of FolderIds,which specifies the number of IDs in the FolderIds field.
            setSearchCriteriaRequest.FolderIdCount = (ushort)folderIds.Length;

            // Set FolderIds to that of logonResponse,which contains identifiers that specify which folders are searched. 
            setSearchCriteriaRequest.FolderIds = folderIds;

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                setSearchCriteriaRequest,
                searchFolderHandle,
                out this.responseSOHs);
            setSearchCriteriaResponse = (RopSetSearchCriteriaResponse)response[0];

            this.Site.Assert.AreEqual<uint>(
               0,
               setSearchCriteriaResponse.ReturnValue,
               "If ROP succeeds, the ReturnValue of its response is 0(success)");

            return setSearchCriteriaResponse;
        }

        /// <summary>
        /// Move the folder into another folder
        /// </summary>
        /// <param name="folderId">The folder ID which will be moved</param>
        /// <param name="newFolderName">The new folder name</param>
        /// <param name="sourceFolderHandle">The handle of the parent folder of which will be moved</param>
        /// <param name="destFolderHandle">The destination folder handle</param>
        /// <returns>The server response</returns>
        private RopMoveFolderResponse MoveFolder(ulong folderId, string newFolderName, uint sourceFolderHandle, uint destFolderHandle)
        {
            RopMoveFolderRequest moveFolderRequest;
            RopMoveFolderResponse moveFolderResponse;

            moveFolderRequest.RopId = (byte)RopId.RopMoveFolder;
            moveFolderRequest.LogonId = TestSuiteBase.LogonId;
            List<uint> handles = new List<uint>
            {
                sourceFolderHandle, destFolderHandle
            };

            // Set SourceHandleIndex to 0x00,which  specifies the location in the Server object handle table
            // where the handle for the source Server object is stored.
            moveFolderRequest.SourceHandleIndex = 0;

            // Set DestHandleIndex to 0x01,which index specifies the location in the Server object handle table
            // where the handle for the destination Server object is stored.
            moveFolderRequest.DestHandleIndex = 1;

            // Set WantAsynchronous to 0x00(FALSE),which specifies the operation is to be executed Synchronously.
            moveFolderRequest.WantAsynchronous = 0;

            // Set UseUnicode to 0x00(FALSE),which specifies the NewFolderName field does not contain Unicode characters or multi-byte characters.
            moveFolderRequest.UseUnicode = 0;

            moveFolderRequest.FolderId = folderId;

            // Set NewFolderName,which specifies the name for the new moved folder.
            moveFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(newFolderName + "\0");

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                moveFolderRequest,
                handles,
                out this.responseSOHs);
            moveFolderResponse = (RopMoveFolderResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                moveFolderResponse.ReturnValue,
               "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return moveFolderResponse;
        }

        /// <summary>
        /// Delete a folder in the mailbox
        /// </summary>
        /// <param name="folderId">The folder ID</param>
        /// <param name="parentFolderHandle">The parent folder handle</param>
        private void DeleteFolder(ulong folderId, uint parentFolderHandle)
        {
            RopDeleteFolderRequest deleteFolderRequest;
            RopDeleteFolderResponse deleteFolderResponse;

            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            deleteFolderRequest.InputHandleIndex = 0;

            deleteFolderRequest.DeleteFolderFlags = (byte)(DeleteFolderFlags.DeleteHardDelete | DeleteFolderFlags.DelFolders | DeleteFolderFlags.DelMessages);

            // Set FolderId to targetFolderId,this folder is to be deleted.
            deleteFolderRequest.FolderId = folderId;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                deleteFolderRequest,
                parentFolderHandle,
                out this.responseSOHs);
            deleteFolderResponse = (RopDeleteFolderResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");
        }

        /// <summary>
        /// Save the changes made on the message
        /// </summary>
        /// <param name="messageHandle">The message handle</param>
        /// <returns>The server response</returns>
        private RopSaveChangesMessageResponse SaveMessage(uint messageHandle)
        {
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = 0;

            // Set ResponseHandleIndex to 0x01,which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.ResponseHandleIndex = 1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                saveChangesMessageRequest,
                messageHandle,
                out this.responseSOHs);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            return saveChangesMessageResponse;
        }

        /// <summary>
        /// Submit the message
        /// </summary>
        /// <param name="messageHandle">The message handle</param>
        private void SubmitMessage(uint messageHandle)
        {
            RopSubmitMessageRequest submitMessageRequest;
            RopSubmitMessageResponse submitMessageResponse;

            submitMessageRequest.RopId = (byte)RopId.RopSubmitMessage;
            submitMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.7.1.1.
            submitMessageRequest.InputHandleIndex = 0;

            submitMessageRequest.SubmitFlags = (byte)SubmitFlags.None;

            IList<IDeserializable> responseMessages = this.cnotifAdapter.Process(
                submitMessageRequest,
                messageHandle,
                out this.responseSOHs);
            submitMessageResponse = (RopSubmitMessageResponse)responseMessages[0];
            this.Site.Assert.AreEqual<uint>(
               0,
               submitMessageResponse.ReturnValue,
               "If ROP succeeds, the ReturnValue of its response is 0(success)");
        }

        /// <summary>
        /// Create a new message in the folder
        /// </summary>
        /// <param name="folderId">The folder ID</param>
        /// <param name="messageHandle">The new message handle</param>
        private void CreateMessage(ulong folderId, out uint messageHandle)
        {
            RopCreateMessageRequest createMessageRequest;
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.6.2.1.
            createMessageRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle,
            // for the output Server object will be stored, as specified in [MS-OXCROPS] section 2.2.6.2.1
            createMessageRequest.OutputHandleIndex = 1;

            // Set CodePageId to 0x0FFF,which specified the code page of Logon object will be used,
            // as specified in [MS-OXCROPS] section 2.2.6.2.1.
            createMessageRequest.CodePageId = 0x0FFF;

            // Set FolderId to the 5th of logonResponse(INBOX),which identifies the parent folder,
            // as specified in [MS-OXCROPS] section 2.2.6.2.1.
            createMessageRequest.FolderId = folderId;

            // Set AssociatedFlag to 0x00,which specified this message is not a folder associated information (FAI) message,
            // as specified in [MS-OXCROPS] section 2.2.6.2.1.
            createMessageRequest.AssociatedFlag = 0;

            IList<IDeserializable> responses = this.cnotifAdapter.Process(
                createMessageRequest,
                this.cnotifAdapter.LogonHandle,
                out this.responseSOHs);
            createMessageResponse = (RopCreateMessageResponse)responses[0];

            this.Site.Assert.AreEqual<uint>(
                0,
                createMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
            messageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// Modify the recipients on the message
        /// </summary>
        /// <param name="messageHandle">The message handle</param>
        /// <param name="recipientColumns">The properties tags want to set</param>
        /// <param name="recipientRows">The recipient rows</param>
        private void ModifyRecipients(uint messageHandle, PropertyTag[] recipientColumns, ModifyRecipientRow[] recipientRows)
        {
            RopModifyRecipientsRequest modifyRecipientsRequest;
            RopModifyRecipientsResponse modifyRecipientsResponse;

            modifyRecipientsRequest.RopId = (byte)RopId.RopModifyRecipients;
            modifyRecipientsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.6.5.1.
            modifyRecipientsRequest.InputHandleIndex = 0;

            // Set ColumnCount, which specifies the number of rows in the RecipientRows field,
            // as specified in [MS-OXCROPS] section 2.2.6.5.1.
            modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;

            // Set RecipientColumns to that created above,which specifies the property values that can be included
            // for each recipient,as specified in [MS-OXCROPS] section 2.2.6.5.1.
            modifyRecipientsRequest.RecipientColumns = recipientColumns;

            // Set RowCount, which specifies the number of rows in the RecipientRows field,
            // as specified in [MS-OXCROPS] section 2.2.6.5.1.
            modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;

            // Set RecipientRows to that created above,which is a list of ModifyRecipientRow structures,
            // as specified in [MS-OXCROPS] section 2.2.6.5.1.
            modifyRecipientsRequest.RecipientRows = recipientRows;

            IList<IDeserializable> response = this.cnotifAdapter.Process(
                modifyRecipientsRequest,
                messageHandle,
                out this.responseSOHs);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)response[0];

            this.Site.Assert.AreEqual<uint>(
                0,
                modifyRecipientsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
        }

        #endregion

        #region Switch Between Sender and Receiver Connections

        /// <summary>
        /// In order to get the notification from the server, need create event in one RPC connection context 
        /// and get the event notification from the other RPC connection context. 
        /// This method only used to change the RPC Context from the receiver RPC connection context to the sender RPC connection context.
        /// When use MAPIHTTP as transport the connection context won't be changed.
        /// </summary>
        private void SwitchRPCContextToSender()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                this.cnotifAdapter.SwitchSessionContext();
            }
            else
            {
                this.cnotifAdapter.RPCContext = this.rpcContextForSend;
                this.cnotifAdapter.LogonHandle = this.senderContextLogonHandle;
            }
        }

        /// <summary>
        /// In order to get the notification from the server, need create event in one RPC connection context 
        /// and get the event notification from the other RPC connection context.  
        /// This method only used to change the RPC Context from the sender RPC connection context to the receiver RPC connection context.
        /// When use MAPIHTTP as transport the connection context won't be changed.
        /// </summary>
        private void SwitchRPCContextToReceiver()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http")
            {
                this.cnotifAdapter.SwitchSessionContext();
            }
            else
            {
                this.cnotifAdapter.RPCContext = this.rpcContextForReceive;
                this.cnotifAdapter.LogonHandle = this.receiverContextLogonHandle;
            }
        }

        #endregion
    }
}