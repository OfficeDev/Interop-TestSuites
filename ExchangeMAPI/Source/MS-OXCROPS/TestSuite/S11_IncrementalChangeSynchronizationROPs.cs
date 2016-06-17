namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Incremental Change Synchronization ROPs.
    /// </summary>
    [TestClass]
    public class S11_IncrementalChangeSynchronizationROPs : TestSuiteBase
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
        /// This method tests the ROP buffers of RopSynchronizationConfigure.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC01_TestRopSynchronizationConfigure()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopSynchronizationConfigure request and verify the success response.
            #region RopSynchronizationConfigure success response

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Call CreateGenericFolderUnderRootFolder method to create a generic folder under the root folder.
            uint targetFolderHandle = this.CreateGenericFolderUnderRootFolder(ref logonResponse);

            // Call CreateFolderSamplePropertyTags method to create property tag samples for folder.
            PropertyTag[] propertyTags = this.CreateFolderSamplePropertyTags();

            RopSynchronizationConfigureRequest synchronizationConfigureRequest;
            RopSynchronizationConfigureResponse synchronizationConfigureResponse;

            // Construct RopSynchronizationConfigure request.
            synchronizationConfigureRequest.RopId = (byte)RopId.RopSynchronizationConfigure;
            
            synchronizationConfigureRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationConfigureRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationConfigureRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationConfigureRequest.SynchronizationType = (byte)SynchronizationType.Contents;
            synchronizationConfigureRequest.SendOptions = (byte)SendOptions.Unicode;
            synchronizationConfigureRequest.SynchronizationFlags = (byte)SynchronizationFlag.Unicode;

            // Set RestrictionDataSize to 0x0000 to avoid the complex RestrictionData.
            synchronizationConfigureRequest.RestrictionDataSize = TestSuiteBase.RestrictionDataSize2;
            synchronizationConfigureRequest.RestrictionData = null;

            // Eid: A server MUST include PidTagFolderId(for hierarchy synchronization) or PidTagMid(for contents synchronization)
            // into a folder change or message change header.
            synchronizationConfigureRequest.SynchronizationExtraFlags = (byte)SynchronizationExtraFlag.Eid;
            synchronizationConfigureRequest.PropertyTagCount = (ushort)propertyTags.Length;
            synchronizationConfigureRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSynchronizationConfigure request.");

            // Send the RopSynchronizationConfigure request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationConfigureRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationConfigureResponse = (RopSynchronizationConfigureResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationConfigureResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 2: Send the RopSynchronizationConfigure request and verify the failure response.
            #region RopSynchronizationConfigure failure response

            synchronizationConfigureRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSynchronizationConfigure request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationConfigureRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationImportMessageChange.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC02_TestRopSynchronizationImportMessageChange()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopSynchronizationImportMessageChange request and verify the success response.
            #region RopSynchronizationImportMessageChange success response

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Call CreateSyncUploadContext method to create an upload context.
            uint synchronizationUploadContextHandle = this.CreateSyncUploadContext(ref logonResponse);

            // Call CreateSamplePropertyValues method to create property value samples.
            TaggedPropertyValue[] propertyValues = this.CreateSamplePropertyValues();

            RopSynchronizationImportMessageChangeRequest synchronizationImportMessageChangeRequest;
            RopSynchronizationImportMessageChangeResponse synchronizationImportMessageChangeResponse;

            // Construct RopSynchronizationImportMessageChange request.
            synchronizationImportMessageChangeRequest.RopId = (byte)RopId.RopSynchronizationImportMessageChange;
            
            synchronizationImportMessageChangeRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationImportMessageChangeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationImportMessageChangeRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationImportMessageChangeRequest.ImportFlag = (byte)ImportFlag.Normal;
            synchronizationImportMessageChangeRequest.PropertyValueCount = (ushort)propertyValues.Length;
            synchronizationImportMessageChangeRequest.PropertyValues = propertyValues;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSynchronizationImportMessageChange request.");

            // Send the RopSynchronizationImportMessageChange request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportMessageChangeRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationImportMessageChangeResponse = (RopSynchronizationImportMessageChangeResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationImportMessageChangeResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 1: Send the RopSynchronizationImportMessageChange request and verify the failure response.
            #region RopSynchronizationImportMessageChange failure response

            synchronizationImportMessageChangeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSynchronizationImportMessageChange request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportMessageChangeRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationImportReadStateChanges.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC03_TestRopSynchronizationImportReadStateChanges()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Create a message.
            #region Create message

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            // Construct the RopCreateMessage request.
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            
            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            createMessageRequest.FolderId = logonResponse.FolderIds[4];
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request and verify the success response.
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

            // Step 2: Save message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            // Construct the RopSaveChangesMessage request.
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request and verify the success response.
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

            // Step 3: Open folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 4: Configure a synchronization upload context.
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            // Construct the RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationOpenCollector request.");

            // Send the RopSynchronizationOpenCollector request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint synchronizationUploadContextHandle = responseSOHs[0][synchronizationOpenCollectorResponse.OutputHandleIndex];

            #endregion

            // Step 5: Send the RopSynchronizationImportReadStateChanges request.
            #region RopSynchronizationImportReadStateChanges response

            RopSynchronizationImportReadStateChangesRequest synchronizationImportReadStateChangesRequest =
                new RopSynchronizationImportReadStateChangesRequest();
            MessageReadState[] messageReadStates = new MessageReadState[1];
            MessageReadState messageReadState = new MessageReadState
            {
                MarkAsRead = Convert.ToByte(TestSuiteBase.Zero)
            };

            // Send the RopLongTermIdFromId request to convert a short-term ID into a long-term ID.
            #region RopLongTermIdFromId response

            RopLongTermIdFromIdRequest ropLongTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse;

            ropLongTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
            ropLongTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;
            ropLongTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropLongTermIdFromIdRequest.ObjectId = saveChangesMessageResponse.MessageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopLongTermIdFromId request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropLongTermIdFromIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropLongTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;

            #endregion

            byte[] messageID = new byte[22];
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, messageID, 0, 16);
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.GlobalCounter, 0, messageID, 16, 6);

            messageReadState.MessageId = messageID;
            messageReadState.MessageIdSize = 22;
            messageReadStates[0] = messageReadState;

            // Construct the RopSynchronizationImportReadStateChanges request.
            synchronizationImportReadStateChangesRequest.RopId = (byte)RopId.RopSynchronizationImportReadStateChanges;
            synchronizationImportReadStateChangesRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationImportReadStateChangesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationImportReadStateChangesRequest.MessageReadStates = messageReadStates;
            synchronizationImportReadStateChangesRequest.MessageReadStateSize = (ushort)messageReadStates[0].Size();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSynchronizationImportReadStateChanges request to invoke success response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportReadStateChangesRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // Test RopSynchronizationImportReadStateChanges failure response.
            synchronizationImportReadStateChangesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSynchronizationImportReadStateChanges request to invoke failure response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportReadStateChangesRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationImportHierarchyChange.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC04_TestRopSynchronizationImportHierarchyChange()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Open folder.
            #region Open folder

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create subfolder.
            #region Create subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            // Construct the RopCreateFolder request.
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            // Send the RopCreateFolder request and verify the success response.
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

            TaggedPropertyValue[] hierarchyValues = this.CreateSampleHierarchyValues(createFolderResponse.FolderId);
            TaggedPropertyValue[] propertyValues = this.CreateSampleFolderPropertyValues();

            #endregion

            // Step 3: Configure a synchronization upload context.
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            // Construct the RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSynchronizationOpenCollector request.");

            // Send the RopSynchronizationOpenCollector request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint synchronizationUploadContextHandle = responseSOHs[0][synchronizationOpenCollectorResponse.OutputHandleIndex];

            #endregion

            // Step 4: Send RopSynchronizationImportHierarchyChange request.
            #region RopSynchronizationImportHierarchyChange response

            RopSynchronizationImportHierarchyChangeRequest synchronizationImportHierarchyChangeRequest;
            RopSynchronizationImportHierarchyChangeResponse synchronizationImportHierarchyChangeResponse;

            // Construct the RopSynchronizationImportHierarchyChange request.
            synchronizationImportHierarchyChangeRequest.RopId = (byte)RopId.RopSynchronizationImportHierarchyChange;
            synchronizationImportHierarchyChangeRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationImportHierarchyChangeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationImportHierarchyChangeRequest.HierarchyValueCount = (ushort)hierarchyValues.Length;
            synchronizationImportHierarchyChangeRequest.HierarchyValues = hierarchyValues;
            synchronizationImportHierarchyChangeRequest.PropertyValueCount = (ushort)propertyValues.Length;
            synchronizationImportHierarchyChangeRequest.PropertyValues = propertyValues;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationImportHierarchyChange request to invoke success response.");

            // Send the RopSynchronizationImportHierarchyChange request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportHierarchyChangeRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationImportHierarchyChangeResponse = (RopSynchronizationImportHierarchyChangeResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationImportHierarchyChangeResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Send the RopSynchronizationImportHierarchyChange request and verify the failure response.
            synchronizationImportHierarchyChangeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationImportHierarchyChange request to invoke success response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportHierarchyChangeRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationImportDeletes.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC05_TestRopSynchronizationImportDeletes()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Open folder.
            #region Open folder

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Configure a synchronization upload context.
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            // Construct the RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSynchronizationOpenCollector request.");

            // Send the RopSynchronizationOpenCollector request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint synchronizationUploadContextHandle = responseSOHs[0][synchronizationOpenCollectorResponse.OutputHandleIndex];

            #endregion

            // Step 3: Create message.
            #region Create message

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            // Construct the RopCreateMessage request.
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = TestSuiteBase.LogonId;
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            createMessageRequest.FolderId = logonResponse.FolderIds[4];
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request and verify the success response.
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

            // Step 4: Save message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            // Construct the RopSaveChangesMessage request.
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request and verify the success response.
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

            // Send the RopSynchronizationImportDeletes request and verify the response.
            #region RopSynchronizationImportDeletes response

            RopSynchronizationImportDeletesRequest ropSynchronizationImportDeletesRequest = new RopSynchronizationImportDeletesRequest();
            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[1];
            TaggedPropertyValue propertyValue = new TaggedPropertyValue();

            // Send the RopLongTermIdFromId request to convert a short-term ID into a long-term ID.
            #region RopLongTermIdFromId response

            RopLongTermIdFromIdRequest ropLongTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse;

            // Construct the RopLongTermIdFromId request.
            ropLongTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
            ropLongTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;
            ropLongTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropLongTermIdFromIdRequest.ObjectId = saveChangesMessageResponse.MessageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetPropertyIdsFromNames request.");

            // Send the RopGetPropertyIdsFromNames request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropLongTermIdFromIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropLongTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;

            #endregion

            byte[] sampleValue = new byte[24];
            propertyValue.PropertyTag.PropertyId = this.propertyDictionary[PropertyNames.PidTagTemplateData].PropertyId;
            propertyValue.PropertyTag.PropertyType = this.propertyDictionary[PropertyNames.PidTagTemplateData].PropertyType;

            // The combination of first two bytes (0x0016) indicates the length of value field.
            sampleValue[0] = 0x16;
            sampleValue[1] = 0x00;
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, sampleValue, 2, 16);
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.GlobalCounter, 0, sampleValue, 18, 6);

            propertyValue.Value = sampleValue;
            propertyValues[0] = propertyValue;

            // Construct the RopSynchronizationImportDeletes request.
            ropSynchronizationImportDeletesRequest.RopId = (byte)RopId.RopSynchronizationImportDeletes;
            ropSynchronizationImportDeletesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropSynchronizationImportDeletesRequest.LogonId = TestSuiteBase.LogonId;
            ropSynchronizationImportDeletesRequest.IsHierarchy = Convert.ToByte(TestSuiteBase.Zero);
            ropSynchronizationImportDeletesRequest.PropertyValueCount = (ushort)propertyValues.Length;
            ropSynchronizationImportDeletesRequest.PropertyValues = propertyValues;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationImportDeletes request to invoke success response.");

            // Send the RopSynchronizationImportDeletes request to get the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropSynchronizationImportDeletesRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // Send the RopSynchronizationImportDeletes request to get the failure response.
            ropSynchronizationImportDeletesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationImportDeletes request to invoke failure response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropSynchronizationImportDeletesRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationImportMessageMove.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC06_TestRopSynchronizationImportMessageMove()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Open folder.
            #region Open folder

            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create the first subfolder in opened folder.
            #region Create the first subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            // Construct RopCreateFolder request.
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            // Send the RopCreateFolder request to the server and verify the success response.
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

            // Step 3: Create the second subfolder in opened folder
            #region Create the second subfolder

            RopCreateFolderResponse createSecondFolderResponse;

            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopCreateFolder request.");

            // Send the RopCreateFolder request to create the second subfolder.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createSecondFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint secondFolderHandle = responseSOHs[0][createSecondFolderResponse.OutputHandleIndex];

            #endregion

            // Step 4: Configure a synchronization upload context
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorMsgRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorMsgResponse;

            // Construct RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorMsgRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorMsgRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorMsgRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorMsgRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorMsgRequest.IsContentsCollector = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationOpenCollector request.");

            // Send the RopSynchronizationOpenCollector request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorMsgRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorMsgResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorMsgResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint synchronizationUploadContextHandleMsg = responseSOHs[0][synchronizationOpenCollectorMsgResponse.OutputHandleIndex];

            #endregion

            // Step 5: Send the RopSynchronizationImportMessageChange request to import new messages 
            // or full changes to existing messages into the server replica.
            #region RopSynchronizationImportMessageChange

            RopSynchronizationImportMessageChangeRequest synchronizationImportMessageChangeRequest;
            RopSynchronizationImportMessageChangeResponse synchronizationImportMessageChangeResponse;

            // Call CreateSamplePropertyValues method to create property value samples.
            TaggedPropertyValue[] propertyValues = this.CreateSamplePropertyValues();

            // Construct the RopSynchronizationImportMessageChange request.
            synchronizationImportMessageChangeRequest.RopId = (byte)RopId.RopSynchronizationImportMessageChange;
            synchronizationImportMessageChangeRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationImportMessageChangeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationImportMessageChangeRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationImportMessageChangeRequest.ImportFlag = (byte)ImportFlag.Normal;
            synchronizationImportMessageChangeRequest.PropertyValueCount = (ushort)propertyValues.Length;
            synchronizationImportMessageChangeRequest.PropertyValues = propertyValues;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSynchronizationImportMessageChange request.");

            // Send the RopSynchronizationImportMessageChange request and get its output handle.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationImportMessageChangeRequest,
                synchronizationUploadContextHandleMsg,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationImportMessageChangeResponse = (RopSynchronizationImportMessageChangeResponse)response;
            uint targetMessageHandle = responseSOHs[0][synchronizationImportMessageChangeResponse.OutputHandleIndex];

            #endregion

            // Step 6: Save message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            // Construct the RopSaveChangesMessage request.
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;
            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request to the server and verify the success response.
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

            // Step 7: Send the RopSynchronizationImportMessageMove request and verify the response.
            #region RopSynchronizationImportMessageMove response.

            RopSynchronizationImportMessageMoveRequest importMessageMoveRequest = new RopSynchronizationImportMessageMoveRequest();
            RopSynchronizationImportMessageMoveResponse importMessageMoveResponse;

            // Construct the RopSynchronizationImportMessageMove request.
            #region Construct the RopSynchronizationImportMessageMove request

            importMessageMoveRequest.RopId = (byte)RopId.RopSynchronizationImportMessageMove;
            importMessageMoveRequest.LogonId = TestSuiteBase.LogonId;
            importMessageMoveRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            importMessageMoveRequest.SourceFolderIdSize = 22;
            byte[] value = new byte[22];

            // Send the RopLongTermIdFromId request to convert the short-term ID into a long-term ID.
            #region convert the short-term ID into the long-term ID

            RopLongTermIdFromIdRequest ropLongTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse;

            // Construct the RopLongTermIdFromId request.
            ropLongTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
            ropLongTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;
            ropLongTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropLongTermIdFromIdRequest.ObjectId = createFolderResponse.FolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSaveChangesMessage request.");

            // Send the RopLongTermIdFromId request to convert the short-term ID into a long-term ID.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropLongTermIdFromIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropLongTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, value, 0, 16);
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.GlobalCounter, 0, value, 16, 6);

            #endregion

            importMessageMoveRequest.SourceFolderId = value;
            importMessageMoveRequest.SourceMessageIdSize = 22;
            byte[] value1 = new byte[22];

            Array.Copy(propertyValues[0].Value, 2, value1, 0, 22);
            importMessageMoveRequest.SourceMessageId = value1;
            byte[] value3 = new byte[22];

            // Send the RopGetLocalReplicaIds to reserve a range of IDs to be used by a local replica.
            #region Reserve a range of IDs

            RopGetLocalReplicaIdsRequest ropGetLocalReplicaIdsRequest;
            RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse;

            // Construct the RopGetLocalReplicaIds request.
            ropGetLocalReplicaIdsRequest.IdCount = 2;
            ropGetLocalReplicaIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropGetLocalReplicaIdsRequest.LogonId = TestSuiteBase.LogonId;
            ropGetLocalReplicaIdsRequest.RopId = (byte)RopId.RopGetLocalReplicaIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopGetLocalReplicaIds request.");

            // Send the RopGetLocalReplicaIds to reserve a range of IDs to be used by a local replica.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropGetLocalReplicaIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropGetLocalReplicaIdsResponse = (RopGetLocalReplicaIdsResponse)response;
            Array.Copy(ropGetLocalReplicaIdsResponse.ReplGuid, 0, value3, 0, 16);
            Array.Copy(ropGetLocalReplicaIdsResponse.GlobalCount, 0, value3, 16, 6);

            #endregion

            importMessageMoveRequest.DestinationMessageIdSize = 22;
            importMessageMoveRequest.DestinationMessageId = value3;

            // PidTagChangeKey
            byte[] bytesForChangeNumber = new byte[20];
            byte[] guid = Guid.NewGuid().ToByteArray();
            Array.Copy(guid, 0, bytesForChangeNumber, 0, 16);
            importMessageMoveRequest.ChangeNumberSize = 20;
            importMessageMoveRequest.ChangeNumber = bytesForChangeNumber;
            importMessageMoveRequest.PredecessorChangeListSize = 23;
            byte[] bytesForPredecessorChangeList = 
            {
                0x16, 0x19, 0xD7, 0xFB, 0x0F,
                0x06, 0x16, 0xA1, 0x41, 0xBF,
                0xF6, 0x91, 0xC7, 0x63, 0xDA,
                0xA8, 0x66, 0x00, 0x00, 0x00,
                0x78, 0x4D, 0x1C
            };
            importMessageMoveRequest.PredecessorChangeList = bytesForPredecessorChangeList;

            // Configure a synchronization upload context
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            // Construct the RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSynchronizationOpenCollector request.");

            // Send the RopSynchronizationOpenCollector request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                secondFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint synchronizationUploadContextHandle = responseSOHs[0][synchronizationOpenCollectorResponse.OutputHandleIndex];

            #endregion

            #endregion

            // RopSynchronizationImportMessageMove success response.
            #region RopSynchronizationImportMessageMove success response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSynchronizationImportMessageMove request to invoke success response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                importMessageMoveRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            importMessageMoveResponse = (RopSynchronizationImportMessageMoveResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                importMessageMoveResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // RopSynchronizationImportMessageMove failure response.
            #region RopSynchronizationImportMessageMove failure response

            importMessageMoveRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSynchronizationImportMessageMove request to invoke failure response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                importMessageMoveRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationOpenCollector.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC07_TestRopSynchronizationOpenCollector()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 2: Call CreateGenericFolderUnderRootFolder method to open a folder and create a generic folder under the opened folder.
            uint targetFolderHandle = this.CreateGenericFolderUnderRootFolder(ref logonResponse);

            // Step 3: Construct the RopSynchronizationOpenCollector request.
            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationOpenCollector request.");

            // Step 4: Send the RopSynchronizationOpenCollector request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationGetTransferState.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC08_TestRopSynchronizationGetTransferState()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 2: Call CreateSyncUploadContext method to create an upload context.
            uint synchronizationUploadContextHandle = this.CreateSyncUploadContext(ref logonResponse);

            // Step 3: Construct the RopSynchronizationGetTransferState request.
            RopSynchronizationGetTransferStateRequest synchronizationGetTransferStateRequest;
            RopSynchronizationGetTransferStateResponse synchronizationGetTransferStateResponse;

            synchronizationGetTransferStateRequest.RopId = (byte)RopId.RopSynchronizationGetTransferState;
            synchronizationGetTransferStateRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationGetTransferStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationGetTransferStateRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the CreateSyncUploadContext request.");

            // Step 4: Send the CreateSyncUploadContext request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationGetTransferStateRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationGetTransferStateResponse = (RopSynchronizationGetTransferStateResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationGetTransferStateResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Step 5: Send the CreateSyncUploadContext request and verify the success response.
            synchronizationGetTransferStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the CreateSyncUploadContext request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationGetTransferStateRequest,
                synchronizationUploadContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSynchronizationUploadStateStreamBegin, 
        /// RopSynchronizationUploadStateStreamContinue and RopSynchronizationUploadStateStreamEnd.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC09_TestRopSynchronizationUploadStateStreamEnd()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2:Call CreateSyncUploadContext method to create an upload context.");

            // Step 2: Call CreateSyncUploadContext method to create an upload context.
            uint synchronizationContextHandle = this.CreateSyncUploadContext(ref logonResponse);

            // Step 3: Send the RopSynchronizationUploadStateStreamBegin request.
            #region RopSynchronizationUploadStateStreamBegin response.

            // Construct the RopSynchronizationUploadStateStreamBegin request.
            RopSynchronizationUploadStateStreamBeginRequest synchronizationUploadStateStreamBeginRequest;
            synchronizationUploadStateStreamBeginRequest.RopId = (byte)RopId.RopSynchronizationUploadStateStreamBegin;
            
            synchronizationUploadStateStreamBeginRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationUploadStateStreamBeginRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationUploadStateStreamBeginRequest.StateProperty =
               this.GetStatePropertyByIds(
               this.propertyDictionary[PropertyNames.PidTagCnsetSeen].PropertyId,
               this.propertyDictionary[PropertyNames.PidTagCnsetSeen].PropertyType);

            // Set TransferBufferSize, which is defined by tester.
            synchronizationUploadStateStreamBeginRequest.TransferBufferSize = 0x00000013;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSynchronizationUploadStateStreamBegin request.");

            // Send the RopSynchronizationUploadStateStreamBegin request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationUploadStateStreamBeginRequest,
                synchronizationContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion

            // Step 4: Send the RopSynchronizationUploadStateStreamContinue request.
            #region RopSynchronizationUploadStateStreamContinue response

            // Construct the RopSynchronizationUploadStateStreamContinue request.
            RopSynchronizationUploadStateStreamContinueRequest synchronizationUploadStateStreamContinueRequest;
            byte[] sampleStreamData = 
            { 
                0x11, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00 
            };
            synchronizationUploadStateStreamContinueRequest.RopId = (byte)RopId.RopSynchronizationUploadStateStreamContinue;
            synchronizationUploadStateStreamContinueRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationUploadStateStreamContinueRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationUploadStateStreamContinueRequest.StreamDataSize = (uint)sampleStreamData.Length;
            synchronizationUploadStateStreamContinueRequest.StreamData = sampleStreamData;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSynchronizationUploadStateStreamContinue request.");

            // Send the RopSynchronizationUploadStateStreamContinue request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationUploadStateStreamContinueRequest,
                synchronizationContextHandle, 
                ref this.response, 
                ref this.rawData, 
                RopResponseType.SuccessResponse);

            #endregion

            // Step 5: Send the RopSynchronizationUploadStateStreamEnd request.
            #region RopSynchronizationUploadStateStreamEnd response

            // Construct the RopSynchronizationUploadStateStreamEnd request.
            RopSynchronizationUploadStateStreamEndRequest synchronizationUploadStateStreamEndRequest;
            synchronizationUploadStateStreamEndRequest.RopId = (byte)RopId.RopSynchronizationUploadStateStreamEnd;
            synchronizationUploadStateStreamEndRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationUploadStateStreamEndRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSynchronizationUploadStateStreamEnd request.");

            // Send the RopSynchronizationUploadStateStreamEnd request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationUploadStateStreamEndRequest,
                synchronizationContextHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopGetLocalReplicaIds and RopSetLocalReplicaMidsetDeleted.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S11_TC10_TestRopSetLocalReplicaMidsetDeleted()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopGetLocalReplicaIds request.
            #region RopGetLocalReplicaIds response

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopGetLocalReplicaIdsRequest getLocalReplicaIdsRequest;
            RopGetLocalReplicaIdsResponse getLocalReplicaIdsResponse;

            // Construct the RopGetLocalReplicaIds request.
            getLocalReplicaIdsRequest.RopId = (byte)RopId.RopGetLocalReplicaIds;
            
            getLocalReplicaIdsRequest.LogonId = TestSuiteBase.LogonId;
            getLocalReplicaIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set IdCount, which specifies the number of IDs to reserve.
            getLocalReplicaIdsRequest.IdCount = 0x00000003;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetLocalReplicaIds request.");

            // Send the RopGetLocalReplicaIds request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getLocalReplicaIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getLocalReplicaIdsResponse = (RopGetLocalReplicaIdsResponse)response;

            #endregion

            // Step 2: Set LongTermIdRange.
            #region Set LongTermIdRange

            LongTermIdRange longTermIdRange = new LongTermIdRange();
            byte[] valueMax = new byte[24];
            byte[] valueMin = new byte[24];
            Array.Copy(getLocalReplicaIdsResponse.ReplGuid, 0, valueMax, 0, 16);
            Array.Copy(getLocalReplicaIdsResponse.ReplGuid, 0, valueMin, 0, 16);

            Array.Copy(getLocalReplicaIdsResponse.GlobalCount, 0, valueMin, 16, 6);
            Array.Copy(getLocalReplicaIdsResponse.GlobalCount, 0, valueMax, 16, 6);

            longTermIdRange.MaxLongTermId = valueMax;
            longTermIdRange.MinLongTermId = valueMin;

            LongTermIdRange[] longTermIdRanges = new LongTermIdRange[1];
            longTermIdRanges[0] = longTermIdRange;

            #endregion

            // Step 3: Open a folder and get its handle.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 4: Send the RopSetLocalReplicaMidsetDeleted request.
            #region RopSetLocalReplicaMidsetDeleted response

            RopSetLocalReplicaMidsetDeletedRequest setLocalReplicaMidsetDeletedRequest = new RopSetLocalReplicaMidsetDeletedRequest
            {
                RopId = (byte)RopId.RopSetLocalReplicaMidsetDeleted,
                LogonId = TestSuiteBase.LogonId,
                InputHandleIndex = TestSuiteBase.InputHandleIndex0,
                DataSize = 52,
                LongTermIdRangeCount = 1,
                LongTermIdRanges = longTermIdRanges
            };

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetLocalReplicaMidsetDeleted request to invoke success response.");

            // Send the RopSetLocalReplicaMidsetDeleted request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setLocalReplicaMidsetDeletedRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // Test RopSetLocalReplicaMidsetDeleted failure response.
            getLocalReplicaIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetLocalReplicaMidsetDeleted request to invoke failure response.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getLocalReplicaIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        #endregion

        #region Common methods

        /// <summary>
        /// Create an upload context.
        /// </summary>
        /// <param name="logonResponse">A logon response.</param>
        /// <returns>An upload context handle.</returns>
        private uint CreateSyncUploadContext(ref RopLogonResponse logonResponse)
        {
            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 3nd of logonResponse, which means it will open the root folder.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request in CreateSyncUploadContext method.");

            // Send the RopOpenFolder request and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a subfolder in opened folder.
            #region Create subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            // Construct the RopCreateFolder request.
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request in CreateSyncUploadContext method.");

            // Send the RopCreateFolder request and verify success response.
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

            // Step 3: Configure a synchronization upload context.
            #region Configure a synchronization upload context

            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse;

            // Construct the RopSynchronizationOpenCollector request.
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = TestSuiteBase.LogonId;
            synchronizationOpenCollectorRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            synchronizationOpenCollectorRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            synchronizationOpenCollectorRequest.IsContentsCollector = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSynchronizationOpenCollector request in CreateSyncUploadContext method.");

            // Send the RopSynchronizationOpenCollector request and verify success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                synchronizationOpenCollectorRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                synchronizationOpenCollectorResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 4: Return the handle of synchronization upload context.
            uint synchronizationUploadContextHandle = responseSOHs[0][synchronizationOpenCollectorResponse.OutputHandleIndex];
            return synchronizationUploadContextHandle;
        }

        /// <summary>
        /// Create property tag samples for folder. 
        /// </summary>
        /// <returns>A property tag array.</returns>
        private PropertyTag[] CreateFolderSamplePropertyTags()
        {
            PropertyTag[] propertyTags = new PropertyTag[5];
            PropertyTag tag = new PropertyTag
            {
                PropertyId = this.propertyDictionary[PropertyNames.PidTagAccess].PropertyId,
                PropertyType = this.propertyDictionary[PropertyNames.PidTagAccess].PropertyType
            };

            // PtypInteger32
            propertyTags[0] = tag;

            // PidTagAccessLevel
            tag.PropertyId = this.propertyDictionary[PropertyNames.PidTagAccessLevel].PropertyId;
            
            // PtypInteger32
            tag.PropertyType = this.propertyDictionary[PropertyNames.PidTagAccessLevel].PropertyType;
            propertyTags[1] = tag;

            // PidTagChangeKey
            tag.PropertyId = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyId;
            
            // PtypBinary
            tag.PropertyType = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyType;
            propertyTags[2] = tag;

            // PidTagCreationTime
            tag.PropertyId = this.propertyDictionary[PropertyNames.PidTagCreationTime].PropertyId;
            
            // PtypTime
            tag.PropertyType = this.propertyDictionary[PropertyNames.PidTagCreationTime].PropertyType;
            propertyTags[3] = tag;

            // PidTagLastModificationTime
            tag.PropertyId = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyId;
            
            // PtypTime
            tag.PropertyType = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyType;
            propertyTags[4] = tag;

            return propertyTags;
        }

        /// <summary>
        /// Create property value samples.
        /// </summary>
        /// <returns>A property value array.</returns>
        private TaggedPropertyValue[] CreateSamplePropertyValues()
        {
            // Step 1: Send the RopGetLocalReplicaIds request to reserve a range of IDs to be used by a local replica.
            #region RopGetLocalReplicaIds

            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[4];
            TaggedPropertyValue propertyValue;

            // PidTagSourceKey
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagSourceKey].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagSourceKey].PropertyType
                }
            };

            byte[] sample = new byte[24];
            
            // The combination of first two bytes (0x0016) indicates the length of value field.
            sample[0] = 0x16;
            sample[1] = 0x00;

            RopGetLocalReplicaIdsRequest ropGetLocalReplicaIdsRequest;
            RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse;

            ropGetLocalReplicaIdsRequest.IdCount = 2;
            ropGetLocalReplicaIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropGetLocalReplicaIdsRequest.LogonId = TestSuiteBase.LogonId;
            ropGetLocalReplicaIdsRequest.RopId = (byte)RopId.RopGetLocalReplicaIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetLocalReplicaIds request in CreateSamplePropertyValues method.");

            // Send the RopGetLocalReplicaIds request to reserve a range of IDs to be used by a local replica.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropGetLocalReplicaIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropGetLocalReplicaIdsResponse = (RopGetLocalReplicaIdsResponse)response;

            #endregion

            // Step 2: Construct the property value samples.
            #region Construct the property value samples

            Array.Copy(ropGetLocalReplicaIdsResponse.ReplGuid, 0, sample, 2, 16);
            Array.Copy(ropGetLocalReplicaIdsResponse.GlobalCount, 0, sample, 18, 6);
            propertyValue.Value = sample;
            propertyValues[0] = propertyValue;

            // PidTagLastModificationTime
            propertyValue = new TaggedPropertyValue();
            byte[] sampleForPidTagLastModificationTime = { 154, 148, 234, 120, 114, 202, 202, 1 };

            propertyValue.PropertyTag.PropertyId = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyId;
            propertyValue.PropertyTag.PropertyType = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyType;
            propertyValue.Value = sampleForPidTagLastModificationTime;
            propertyValues[1] = propertyValue;
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyType
                }
            };

            byte[] sampleForPidTagChangeKey = new byte[22];
            byte[] guid = Guid.NewGuid().ToByteArray();

            // The combination of first two bytes (0x0014) indicates the length of value field.
            sampleForPidTagChangeKey[0] = 0x14;
            sampleForPidTagChangeKey[1] = 0x00;

            Array.Copy(guid, 0, sampleForPidTagChangeKey, 2, 16);
            propertyValue.Value = sampleForPidTagChangeKey;

            propertyValues[2] = propertyValue;

            // PidTagPredecessorChangeList
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagPredecessorChangeList].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagPredecessorChangeList].PropertyType
                }
            };

            byte[] sampleForPidTagPredecessorChangeList = 
            { 
                0x17, 0x00, 0x16, 0x19, 0xD7,
                0xFB, 0x0F, 0x06, 0x16, 0xA1,
                0x41, 0xBF, 0xF6, 0x91, 0xC7,
                0x63, 0xDA, 0xA8, 0x66, 0x00,
                0x00, 0x00, 0x78, 0x4D, 0x1C 
            };
            propertyValue.Value = sampleForPidTagPredecessorChangeList;
            propertyValues[3] = propertyValue;
            return propertyValues;

            #endregion
        }

        /// <summary>
        /// Create hierarchy value samples.
        /// </summary>
        /// <param name="parentFolderId">Parent folder id.</param>
        /// <returns>The property values of hierarchy</returns>
        private TaggedPropertyValue[] CreateSampleHierarchyValues(ulong parentFolderId)
        {
            TaggedPropertyValue[] hierarchyValues = new TaggedPropertyValue[6];
            TaggedPropertyValue propertyValue = new TaggedPropertyValue();

            // PidTagParentSourceKey
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagParentSourceKey].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagParentSourceKey].PropertyType
                }
            };

            // Send RopLongTermIdFromId request to convert a short-term ID into a long-term ID.
            #region RopLongTermIdFromId response

            RopLongTermIdFromIdRequest ropLongTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse;

            ropLongTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;

            ropLongTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;
            ropLongTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropLongTermIdFromIdRequest.ObjectId = parentFolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetLocalReplicaIds request in CreateSampleHierarchyValues method.");

            // Send RopLongTermIdFromId request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropLongTermIdFromIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            ropLongTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;
            byte[] sampleForPidTagParentSourceKey = new byte[24];

            // The combination of first two bytes (0x0016) indicates the length of value field.
            sampleForPidTagParentSourceKey[0] = 0x16;
            sampleForPidTagParentSourceKey[1] = 0x00;

            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, sampleForPidTagParentSourceKey, 2, 16);
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.GlobalCounter, 0, sampleForPidTagParentSourceKey, 18, 6);

            #endregion

            propertyValue.Value = sampleForPidTagParentSourceKey;
            hierarchyValues[0] = propertyValue;

            // PidTagSourceKey
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagSourceKey].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagSourceKey].PropertyType
                }
            };
            byte[] sampleForPidTagSourceKey = new byte[24];

            // The combination of first two bytes (0x0016) indicates the length of value field.
            sampleForPidTagSourceKey[0] = 0x16;
            sampleForPidTagSourceKey[1] = 0x00;

            // Send the RopGetLocalReplicaIds request to reserve a range of IDs to be used by a local replica.
            #region RopGetLocalReplicaIds response

            RopGetLocalReplicaIdsRequest ropGetLocalReplicaIdsRequest;
            RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse;

            ropGetLocalReplicaIdsRequest.IdCount = 1;
            ropGetLocalReplicaIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            ropGetLocalReplicaIdsRequest.LogonId = TestSuiteBase.LogonId;
            ropGetLocalReplicaIdsRequest.RopId = (byte)RopId.RopGetLocalReplicaIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetLocalReplicaIds request in CreateSampleHierarchyValues method.");

            // Send RopGetLocalReplicaIds request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropGetLocalReplicaIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropGetLocalReplicaIdsResponse = (RopGetLocalReplicaIdsResponse)response;

            Array.Copy(ropGetLocalReplicaIdsResponse.ReplGuid, 0, sampleForPidTagSourceKey, 2, 16);
            Array.Copy(ropGetLocalReplicaIdsResponse.GlobalCount, 0, sampleForPidTagSourceKey, 18, 6);

            #endregion

            propertyValue.Value = sampleForPidTagSourceKey;
            hierarchyValues[1] = propertyValue;

            // PidTagLastModificationTime
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagLastModificationTime].PropertyType
                }
            };
            byte[] sampleForPidTagLastModificationTime = { 154, 148, 234, 120, 114, 202, 202, 1 };
            propertyValue.Value = sampleForPidTagLastModificationTime;
            hierarchyValues[2] = propertyValue;

            // PidTagChangeKey
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagChangeKey].PropertyType
                }
            };
            byte[] sampleForPidTagChangeKey = new byte[22];
            byte[] guid = Guid.NewGuid().ToByteArray();

            // The combination of first two bytes (0x0014) indicates the length of value field.
            sampleForPidTagChangeKey[0] = 0x14;
            sampleForPidTagChangeKey[1] = 0x00;

            Array.Copy(guid, 0, sampleForPidTagChangeKey, 2, 16);
            propertyValue.Value = sampleForPidTagChangeKey;
            hierarchyValues[3] = propertyValue;

            // PidTagPredecessorChangeList
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagPredecessorChangeList].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagPredecessorChangeList].PropertyType
                }
            };
            byte[] sampleForPidTagPredecessorChangeList = 
            { 
                0x17, 0x00, 0x16, 0x19, 0xD7,
                0xFB, 0x0F, 0x06, 0x16, 0xA1,
                0x41, 0xBF, 0xF6, 0x91, 0xC7,
                0x63, 0xDA, 0xA8, 0x66, 0x00,
                0x00, 0x00, 0x78, 0x4D, 0x1C 
            };
            propertyValue.Value = sampleForPidTagPredecessorChangeList;
            hierarchyValues[4] = propertyValue;

            // PidTagDisplayName
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagDisplayName].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagDisplayName].PropertyType
                }
            };

            byte[] sampleForPidTagDisplayName = new byte[Encoding.Unicode.GetByteCount(DisplayNameAndCommentForNonSearchFolder + "\0")];
            Array.Copy(
                Encoding.Unicode.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0"),
                0,
                sampleForPidTagDisplayName,
                0,
                Encoding.Unicode.GetByteCount(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0"));
            propertyValue.Value = sampleForPidTagDisplayName;
            hierarchyValues[5] = propertyValue;

            return hierarchyValues;
        }

        /// <summary>
        /// Create property values for a folder.
        /// </summary>
        /// <returns>A property value array.</returns>
        private TaggedPropertyValue[] CreateSampleFolderPropertyValues()
        {
            TaggedPropertyValue[] folderPropertyValues = new TaggedPropertyValue[1];
            TaggedPropertyValue taggedPropertyValue;

            // PidTagFolderType
            taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagFolderType].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagFolderType].PropertyType
                }
            };

            // FOLDER_GENERIC
            byte[] sample = { 0x01, 0x00, 0x00, 0x00 };
            taggedPropertyValue.Value = sample;
            folderPropertyValues[0] = taggedPropertyValue;

            return folderPropertyValues;
        }

        /// <summary>
        /// Create a generic folder under the root folder.
        /// </summary>
        /// <param name="logonResponse">A logon response.</param>
        /// <returns>The object handle of the created folder.</returns>
        private uint CreateGenericFolderUnderRootFolder(ref RopLogonResponse logonResponse)
        {
            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            // Construct the RopOpenFolder request.
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request in CreateGenericFolderUnderRootFolder method.");

            // Send the RopOpenFolder request to the server and verify the success response.
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

            // Get the handle of opened folder, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a subfolder in opened folder.
            #region Create subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            // Construct the RopCreateFolder request.
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request in CreateGenericFolderUnderRootFolder method.");

            // Send the RopCreateFolder request to the server and verify the success response.
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

            #endregion

            // Step 3: Get and return the handle of the created folder.
            uint targetFolderHandle = responseSOHs[0][createFolderResponse.OutputHandleIndex];
            return targetFolderHandle;
        }

        #endregion
    }
}