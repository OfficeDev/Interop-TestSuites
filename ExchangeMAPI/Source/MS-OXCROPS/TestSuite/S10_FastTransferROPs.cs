namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Fast Transfer ROPs. 
    /// </summary>
    [TestClass]
    public class S10_FastTransferROPs : TestSuiteBase
    {
        #region Class Initialization and Cleanup

        /// <summary>
        /// Class initialize.
        /// </summary>
        /// <param name="testContext">The session context handle</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferSourceCopyMessages.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC01_TestRopFastTransferSourceCopyMessages()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCreateMessage request to create a message
            #region RopCreateMessage success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;

            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Create a message in INBOX
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSaveChangesMessage request to save changes
            #region RopSaveChangesMessage success response

            // Save message 
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");
            ulong messageId = saveChangesMessageResponse.MessageId;

            #endregion

            // Step 3: Send the RopOpenFolder request to open folder containing created message
            #region RopOpenFoldersuccess response

            // Open the folder(Inbox) containing the created message
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Inbox will be opened
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");
            uint folderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageId;

            #endregion

            // Step 4: Send the RopFastTransferSourceCopyMessages request to verify success response
            #region RopFastTransferSourceCopyMessages success response

            RopFastTransferSourceCopyMessagesRequest fastTransferSourceCopyMessagesRequest;
            RopFastTransferSourceCopyMessagesResponse fastTransferSourceCopyMessagesResponse;

            fastTransferSourceCopyMessagesRequest.RopId = (byte)RopId.RopFastTransferSourceCopyMessages;
            fastTransferSourceCopyMessagesRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceCopyMessagesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferSourceCopyMessagesRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            fastTransferSourceCopyMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            fastTransferSourceCopyMessagesRequest.MessageIds = messageIds;
            fastTransferSourceCopyMessagesRequest.CopyFlags = (byte)RopFastTransferSourceCopyMessagesCopyFlags.BestBody;
            fastTransferSourceCopyMessagesRequest.SendOptions = (byte)SendOptions.ForceUnicode;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopFastTransferSourceCopyMessages request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceCopyMessagesRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferSourceCopyMessagesResponse = (RopFastTransferSourceCopyMessagesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferSourceCopyMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferSourceGetBuffer.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC02_TestRopFastTransferSourceGetBuffer()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCreateMessage request to create a message
            #region RopCreateMessage success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;

            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Create a message in INBOX
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSaveChangesMessage request to save changes
            #region RopSaveChangesMessage success response

            // Save message 
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");
            ulong messageId = saveChangesMessageResponse.MessageId;

            #endregion

            // Step 3: Send the RopOpenFolder request to open folder containing created message
            #region RopOpenFoldersuccess response

            // Open the folder(Inbox) containing the created message
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Inbox will be opened
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");
            uint folderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageId;

            #endregion

            // Step 4: Send the RopFastTransferSourceCopyMessages request to verify success response
            #region RopFastTransferSourceCopyMessages success response

            RopFastTransferSourceCopyMessagesRequest fastTransferSourceCopyMessagesRequest;
            RopFastTransferSourceCopyMessagesResponse fastTransferSourceCopyMessagesResponse;

            fastTransferSourceCopyMessagesRequest.RopId = (byte)RopId.RopFastTransferSourceCopyMessages;
            fastTransferSourceCopyMessagesRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceCopyMessagesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferSourceCopyMessagesRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            fastTransferSourceCopyMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            fastTransferSourceCopyMessagesRequest.MessageIds = messageIds;
            fastTransferSourceCopyMessagesRequest.CopyFlags = (byte)RopFastTransferSourceCopyMessagesCopyFlags.BestBody;
            fastTransferSourceCopyMessagesRequest.SendOptions = (byte)SendOptions.ForceUnicode;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopFastTransferSourceCopyMessages request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceCopyMessagesRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferSourceCopyMessagesResponse = (RopFastTransferSourceCopyMessagesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferSourceCopyMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");
            uint downloadContextHandle = responseSOHs[0][fastTransferSourceCopyMessagesResponse.OutputHandleIndex];

            #endregion

            // Step 5: Send the RopFastTransferSourceGetBuffer request to verify success response
            #region RopFastTransferSourceGetBuffer success response

            RopFastTransferSourceGetBufferRequest fastTransferSourceGetBufferRequest;
            RopFastTransferSourceGetBufferResponse fastTransferSourceGetBufferResponse;

            fastTransferSourceGetBufferRequest.RopId = (byte)RopId.RopFastTransferSourceGetBuffer;
            fastTransferSourceGetBufferRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceGetBufferRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // The server determines the buffer size based on the residual size of the RPC buffer
            fastTransferSourceGetBufferRequest.BufferSize = (ushort)MS_OXCROPSAdapter.BufferSize;

            // This value specifies the maximum size limit when the server determines the buffer size.
            // as specified in [MS-OXCROPS].
            fastTransferSourceGetBufferRequest.MaximumBufferSize = TestSuiteBase.MaximumBufferSize;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopFastTransferSourceGetBuffer request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceGetBufferRequest,
                downloadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferSourceGetBufferResponse = (RopFastTransferSourceGetBufferResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferSourceGetBufferResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferDestinationConfigure and RopFastTransferDestinationPutBuffer.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC03_TestRopFastTransferDestinationPutBuffer()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopOpenFolder request to OpenFolder 
            #region RopOpenFolder success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Open a folder first, then create a subfolder under the opened folder
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // This handle will be used as input handle in RopCreateFolder
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopCreateFolder request to create a subfolder under the opened folder
            #region RopCreateFolder success response

            // Create a subfolder in opened folder
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint targetFolderHandle = responseSOHs[0][createFolderResponse.OutputHandleIndex];

            #endregion

            // Step 3: Send the RopFastTransferDestinationConfigure request to verify success response
            #region RopFastTransferDestinationConfigure success response

            RopFastTransferDestinationConfigureRequest fastTransferDestinationConfigureRequest;
            RopFastTransferDestinationConfigureResponse fastTransferDestinationConfigureResponse;

            fastTransferDestinationConfigureRequest.RopId = (byte)RopId.RopFastTransferDestinationConfigure;
            fastTransferDestinationConfigureRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferDestinationConfigureRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferDestinationConfigureRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            fastTransferDestinationConfigureRequest.SourceOperation = (byte)SourceOperation.CopyTo;

            // The client identifies the FastTransfer operation being configured as a logical part of a larger object move operation
            fastTransferDestinationConfigureRequest.CopyFlags = (byte)RopFastTransferDestinationConfigureCopyFlags.Move;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopFastTransferDestinationConfigure request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferDestinationConfigureRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferDestinationConfigureResponse = (RopFastTransferDestinationConfigureResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferDestinationConfigureResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint fastTranferUploadContextHandle = responseSOHs[0][fastTransferDestinationConfigureResponse.OutputHandleIndex];

            #endregion

            // Step 4: Send the RopFastTransferDestinationPutBuffer request to verify success response
            #region RopFastTransferDestinationPutBuffer success response

            RopFastTransferDestinationPutBufferRequest fastTransferDestinationPutBufferRequest;

            // Set to a marker value (StartTopFld) refer to the table in MS-OXCFXICS.
            byte[] transferData = { 0x03, 0x00, 0x09, 0x40 };
            fastTransferDestinationPutBufferRequest.RopId = (byte)RopId.RopFastTransferDestinationPutBuffer;
            fastTransferDestinationPutBufferRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferDestinationPutBufferRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferDestinationPutBufferRequest.TransferDataSize = (ushort)transferData.Length;
            fastTransferDestinationPutBufferRequest.TransferData = transferData;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopFastTransferDestinationPutBuffer request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferDestinationPutBufferRequest,
                fastTranferUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferSourceCopyFolder.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC04_TestRopFastTransferSourceCopyFolder()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopOpenFolder request to open folder
            #region RopFastTransferDestinationPutBuffer success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Open a folder first, then create a subfolder under the opened folder
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Open root folder
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // This handle will be used as input handle in RopCreateFolder
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopCreateFolder request to create subfolder under opened folder
            #region RopCreateFolder success response

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0, which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint targetFolderHandle = responseSOHs[0][createFolderResponse.OutputHandleIndex];

            #endregion

            // Step 3: Send the RopFastTransferSourceCopyFolder request to create subfolder under opened folder
            #region RopFastTransferSourceCopyFolder success response

            RopFastTransferSourceCopyFolderRequest fastTransferSourceCopyFolderRequest;
            RopFastTransferSourceCopyFolderResponse fastTransferSourceCopyFolderResponse;

            fastTransferSourceCopyFolderRequest.RopId = (byte)RopId.RopFastTransferSourceCopyFolder;
            fastTransferSourceCopyFolderRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceCopyFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferSourceCopyFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            fastTransferSourceCopyFolderRequest.CopyFlags = (byte)RopFastTransferSourceCopyFolderCopyFlags.CopySubfolders;
            fastTransferSourceCopyFolderRequest.SendOptions = (byte)SendOptions.ForceUnicode;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopFastTransferSourceCopyFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceCopyFolderRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferSourceCopyFolderResponse = (RopFastTransferSourceCopyFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferSourceCopyFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopTellVersion.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC05_TestRopTellVersion()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to the private mailbox.
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            RopTellVersionRequest tellVersionRequest;
            ushort[] version = new ushort[3];

            tellVersionRequest.RopId = (byte)RopId.RopTellVersion;

            tellVersionRequest.LogonId = TestSuiteBase.LogonId;
            tellVersionRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            tellVersionRequest.Version = version;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopTellVersion request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                tellVersionRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
        }

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferSourceCopyTo.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC06_TestRopFastTransferSourceCopyTo()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCreateMessage request to create a message
            #region RopCreateMessage success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;

            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Create a message in INBOX
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSaveChangesMessage request to save changes
            #region RopSaveChangesMessage success response

            // Save message 
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");

            #endregion

            // Step 3: Send the RopFastTransferSourceCopyTo request to verify success response
            #region RopFastTransferSourceCopyTo response

            RopFastTransferSourceCopyToRequest fastTransferSourceCopyToRequest;
            RopFastTransferSourceCopyToResponse fastTransferSourceCopyToResponse;

            PropertyTag[] messagePropertyTags = this.CreateMessageSamplePropertyTags();
            fastTransferSourceCopyToRequest.RopId = (byte)RopId.RopFastTransferSourceCopyTo;
            fastTransferSourceCopyToRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceCopyToRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferSourceCopyToRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // This value specifies the level at which the copy is occurring, which is specified in [MS-OXCROPS]
            // Non-Zero: exclude all descendant sub-objects from being copied
            fastTransferSourceCopyToRequest.Level = TestSuiteBase.LevelOfZero;

            fastTransferSourceCopyToRequest.CopyFlags = (uint)RopFastTransferSourceCopyToCopyFlags.BestBody;
            fastTransferSourceCopyToRequest.SendOptions = (byte)SendOptions.Unicode;
            fastTransferSourceCopyToRequest.PropertyTagCount = (ushort)messagePropertyTags.Length;
            fastTransferSourceCopyToRequest.PropertyTags = messagePropertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopFastTransferSourceCopyTo request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceCopyToRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            fastTransferSourceCopyToResponse = (RopFastTransferSourceCopyToResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                fastTransferSourceCopyToResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopFastTransferSourceCopyProperties.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S10_TC07_TestRopFastTransferSourceCopyProperties()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCreateMessage request to create a message
            #region RopCreateMessage success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;

            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Create a message in INBOX
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSaveChangesMessage request to save changes
            #region RopSaveChangesMessage success response

            // Save message 
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success).");

            #endregion

            // Step 3: Send the RopFastTransferSourceCopyTo request to verify success response
            #region RopFastTransferSourceCopyTo response

            RopFastTransferSourceCopyPropertiesRequest fastTransferSourceCopyPropertiesRequest;

            PropertyTag[] messagePropertyTags = this.CreateMessageSamplePropertyTags();
            fastTransferSourceCopyPropertiesRequest.RopId = (byte)RopId.RopFastTransferSourceCopyProperties;
            fastTransferSourceCopyPropertiesRequest.LogonId = TestSuiteBase.LogonId;
            fastTransferSourceCopyPropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            fastTransferSourceCopyPropertiesRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // This value specifies the level at which the copy is occurring, which is specified in [MS-OXCROPS].
            fastTransferSourceCopyPropertiesRequest.Level = TestSuiteBase.LevelOfNonZero;

            // Move: the client identifies the FastTransfer operation being configured as a logical part of a larger object move operation.
            fastTransferSourceCopyPropertiesRequest.CopyFlags = (byte)RopFastTransferSourceCopyPropertiesCopyFlags.Move;

            fastTransferSourceCopyPropertiesRequest.SendOptions = (byte)SendOptions.Unicode;
            fastTransferSourceCopyPropertiesRequest.PropertyTagCount = (ushort)messagePropertyTags.Length;
            fastTransferSourceCopyPropertiesRequest.PropertyTags = messagePropertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopFastTransferSourceCopyProperties request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                fastTransferSourceCopyPropertiesRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        #endregion

        #region Common method

        /// <summary>
        /// PropertyTag for the message.
        /// </summary>
        /// <returns>Array of PropertyTag</returns>
        private PropertyTag[] CreateMessageSamplePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[1];
            PropertyTag tag = new PropertyTag
            {
                PropertyId = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyId,
                PropertyType = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyType
            };

            // PidTagLastModificationTime

            // PtypTime
            propertyTags[0] = tag;

            return propertyTags;
        }

        #endregion
    }
}