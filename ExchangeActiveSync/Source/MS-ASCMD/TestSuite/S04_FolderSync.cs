namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the FolderSync command.
    /// </summary>
    [TestClass]
    public class S04_FolderSync : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is used to verify FolderSync command, all folders MUST be returned to the client when initial folder synchronization is done with a synchronization key of 0(zero).
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC01_FolderSync_SyncKey0()
        {
            // Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            Site.Assert.AreEqual<byte>(
                 (byte)1,
                 folderSyncResponse.ResponseData.Status,
                 "The server should return a status code 1 in the FolderSync command response to indicate success.");

            Site.Assert.IsNotNull(
                 folderSyncResponse.ResponseData.SyncKey,
                 "The server should return a non-null SyncKey in the FolderSync command response to indicate success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R113");

            // The Count element is great than 0 in the FolderSync command response to indicate folders are returned to the client.
            // Verify MS-ASCMD requirement: MS-ASCMD_R113
            Site.CaptureRequirementIfIsTrue(
                folderSyncResponse.ResponseData.Changes.Count > 0,
                113,
                @"[In FolderSync] The FolderSync command synchronizes the collection hierarchy.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R119");

            // The Count element is great than 0 in the FolderSync command response to indicate folders are returned to the client.
            // Verify MS-ASCMD requirement: MS-ASCMD_R119
            Site.CaptureRequirementIfIsTrue(
                folderSyncResponse.ResponseData.Changes.Count > 0,
                119,
                @"[In FolderSync] All folders MUST be returned to the client when initial folder synchronization is done with a synchronization key of 0 (zero).");

            bool isVerifyR5416 = false;
            foreach (FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                string name = add.DisplayName;
                string serverId = add.ServerId;
                foreach (FolderSyncChangesAdd addNew in folderSyncResponse.ResponseData.Changes.Add)
                {
                    if (serverId != addNew.ServerId)
                    {
                        if (name != addNew.DisplayName)
                        {
                            isVerifyR5416 = true;
                        }
                        else
                        {
                            isVerifyR5416 = false;
                            break;
                        }
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5416");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5416
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5416,
                5416,
                @"[In DisplayName(FolderSync)] Subfolder display names MUST be unique for a sample of N (default N=10) within a folder.");

            // Folder has been synced successfully, server returns a non-null SyncKey.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R116");

            // Verify MS-ASCMD requirement: MS-ASCMD_R116
            Site.CaptureRequirementIfIsNotNull(
                folderSyncResponse.ResponseData.SyncKey,
                116,
                @"[In FolderSync] The synchronization key is returned in the SyncKey element of the response if the FolderSync command succeeds.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4074");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4074
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)1,
                folderSyncResponse.ResponseData.Status,
                4074,
                @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4578");

            // Folder has been synced successfully, server returns a non-null SyncKey.
            Site.CaptureRequirementIfIsNotNull(
                folderSyncResponse.ResponseData.SyncKey,
                4578,
                @"[In SyncKey(FolderSync)] After successful folder synchronization, the server MUST send a synchronization key to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5008");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5008
            // Folder has been synced successfully, server returns a non-null SyncKey.
            Site.CaptureRequirementIfIsNotNull(
                folderSyncResponse.ResponseData.SyncKey,
                5008,
                @"[In Synchronizing a Folder Hierarchy] The server responds with a new folderhierarchy:SyncKey element value and provides a list of all the folders in the user's mailbox.");
        }

        /// <summary>
        /// This test case is used to verify FolderSync command, if there are no changes since the last folders synchronization, a Count element value of 0 (zero) is returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC02_FolderSync_NoChanged()
        {
            // The client calls FolderSync command to synchronize the collection hierarchy if no changes occurred for folder.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest(this.LastFolderSyncKey);
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2120");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2120
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                folderSyncResponse.ResponseData.Changes.Count,
                2120,
                @"[In Count] If there are no changes since the last folder synchronization, a Count element value of 0 (zero) is returned.");
        }

        /// <summary>
        /// This test case is used to verify FolderSync command, if any changes have occurred on the server, the count is not equal to 0.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC03_FolderSync_Changed()
        {
            #region Change a new DeviceID and call FolderSync command.
            this.CMDAdapter.ChangeDeviceID("NewDeviceID");
            this.RecordDeviceInfoChanged();
            string folderName = Common.GenerateResourceName(Site, "FolderSync");
            FolderSyncResponse folderSyncResponseForNewDeviceID = this.FolderSync();
            #endregion

            #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(folderSyncResponseForNewDeviceID.ResponseData.SyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            Site.Assert.AreEqual<byte>((byte)1, folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Change the DeviceId back and call method FolderSync to synchronize the collection hierarchy.
            this.CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest(folderSyncResponseForNewDeviceID.ResponseData.SyncKey);
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);
            foreach (Response.FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.DisplayName == folderName)
                {
                    this.User1Information.UserCreatedFolders.Clear();
                    TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, add.ServerId);
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5024");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5024
            Site.CaptureRequirementIfAreEqual<uint>(
                (uint)1,
                folderSyncResponse.ResponseData.Changes.Count,
                5024,
                @"[In Synchronizing a Folder Hierarchy] [FolderSync sequence for folder hierarchy synchronization, order 2:] If any changes have occurred on the server, the new, deleted, or changed folders are returned to the client.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify FolderSync command, if client resynchronizes the existing folder hierarchy, ServerId values do not change. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC04_FolderSync_Resynchronizes()
        {
            // The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
            string folderName = Common.GenerateResourceName(Site, "FolderSync");
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            Site.Assert.AreEqual<byte>((byte)1, folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);

            // The client calls FolderSync method to synchronize the collection hierarchy, then server returns latest folder SyncKey.
            FolderSyncResponse folderSyncResponse = this.FolderSync();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5012");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5012
            Site.CaptureRequirementIfAreEqual(
                folderCreateResponse.ResponseData.ServerId, 
                TestSuiteBase.GetCollectionId(folderSyncResponse, folderName),
                5012, 
                @"[In Synchronizing a Folder Hierarchy] Existing folderhierarchy:ServerId values do not change when the client resynchronizes.");
        }

        /// <summary>
        /// This test case is used to verify FolderSync command, if the SyncKey is an empty string, the status is equal to 9.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC05_FolderSync_Status9()
        {
            // Call method FolderSync with an empty SyncKey to synchronize the collection hierarchy.
            FolderSyncRequest folderSyncRequest = new FolderSyncRequest { RequestData = { SyncKey = string.Empty } };
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4068");

            // If the client sent a malformed or mismatched synchronization key, the server should return a status code 9 in the FolderSync command response.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4068
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)9,
                folderSyncResponse.ResponseData.Status,
                4068,
                @"[In Status(FolderSync)] If the command fails, the Status element contains a code that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4070");

            // If the client sent a malformed or mismatched synchronization key, the server should return a status code 9 in the FolderSync command response.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4070
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)9,
                folderSyncResponse.ResponseData.Status,
                4070,
                @"[In Status(FolderSync)] If one collection fails, a failure status MUST be returned for all collections.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4083");

            // The server should return a status code 9 in the FolderSync command response to indicate the client sent a malformed or mismatched synchronization key.
            // If the SyncKey is an empty string, the status is equal to 9.
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)9,
                folderSyncResponse.ResponseData.Status,
                4083,
                @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 9 is] The client sent a malformed or mismatched synchronization key [, or the synchronization state is corrupted on the server].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4581");

            // The server should return a status code 9 in the FolderSync command response to indicate the client sent a malformed or mismatched synchronization key.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4581
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)9,
                folderSyncResponse.ResponseData.Status,
                4581,
                @"[In SyncKey(FolderSync)] The server MUST return a Status element (section 2.2.3.162.4) value of 9 if the value of the SyncKey element does not match the value of the synchronization key on the server.");
        }

        /// <summary>
        /// This test case is used to verify FolderSync command, if the SyncKey is invalid, the status is equal to 10.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC06_FolderSync_Status10()
        {
            // Call method FolderSync to synchronize the collection hierarchy with a null SyncKey.
            FolderSyncRequest folderSyncRequest = new FolderSyncRequest { RequestData = { SyncKey = null } };
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4086");

            // If the SyncKey is invalid, the status is equal to 10.
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)10,
                folderSyncResponse.ResponseData.Status,
                4086,
                @"[In Status(FolderSync)] [When the scope is Global], [the cause of the status value 10 is] The client sent a FolderSync command request that contains a semantic or syntactic error.");
        }

        /// <summary>
        /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully after adding a folder.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC07_FolderSync_AddFolder()
        {
            #region The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderSync"), "0");
            Site.Assert.AreEqual<byte>((byte)1, folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            bool isVerifyR5860 = false;
            foreach (FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.ServerId == folderCreateResponse.ResponseData.ServerId)
                {
                    isVerifyR5860 = true;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5860");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5860
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5860,
                5860,
                @"[In Add(FolderSync)] [The Add element] creates a new folder on the client.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify FolderSync command synchronizes the updated folder successfully.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC08_FolderSync_UpdateFolder()
        {
            #region Call method FolderCreate command to create a new folder as a child folder of the specified parent folder.
            string folderName = Common.GenerateResourceName(Site, "FolderSync");
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            Site.Assert.AreEqual<byte>((byte)1, folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Change DeviceID
            this.CMDAdapter.ChangeDeviceID("NewDeviceID");
            this.RecordDeviceInfoChanged();
            string folderSyncKey = folderCreateResponse.ResponseData.SyncKey;
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse foldersyncResponseForNewDeviceID = this.FolderSync();
            string changeDeviceIDFolderId = TestSuiteBase.GetCollectionId(foldersyncResponseForNewDeviceID, folderName);
            Site.Assert.IsFalse(string.IsNullOrEmpty(changeDeviceIDFolderId), "If the new folder created by FolderCreate command, server should return a server ID for the new created folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5020");

            // If client sends the FolderSync request successfully, the server must send a synchronization key to the client in a response.
            Site.CaptureRequirementIfIsTrue(
                (foldersyncResponseForNewDeviceID.ResponseData.SyncKey != null) && (foldersyncResponseForNewDeviceID.ResponseData.SyncKey != folderSyncKey),
                5020,
                @"[In Synchronizing a Folder Hierarchy] [FolderSync sequence for folder hierarchy synchronization, order 1:] The server responds with [the folder hierarchy and] a new folderhierarchy:SyncKey value.");

            #endregion

            #region Call method FolderUpdate to rename a folder.
            string folderUpdateName = Common.GenerateResourceName(Site, "FolderUpdate");
            FolderUpdateRequest folderUpdateRequest = Common.CreateFolderUpdateRequest(foldersyncResponseForNewDeviceID.ResponseData.SyncKey, changeDeviceIDFolderId, folderUpdateName, "0");
            this.CMDAdapter.FolderUpdate(folderUpdateRequest);
            #endregion

            #region Restore DeviceID and call FolderSync command.
            this.CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));

            // The client calls FolderSync command to synchronize the collection hierarchy with original device id.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest(folderSyncKey);
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);
            bool isUpdated = false;
            foreach (Response.FolderSyncChangesUpdate update in folderSyncResponse.ResponseData.Changes.Update)
            {
                if (update.DisplayName == folderUpdateName)
                {
                    isUpdated = true;
                    break;
                }
            }

            Site.Assert.IsTrue(isUpdated, "Rename successfully");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully after deleting a folder.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC09_FolderSync_DeleteFolder()
        {
            #region The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
            string folderName = Common.GenerateResourceName(Site, "FolderSync");
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            Site.Assert.AreEqual<byte>((byte)1, folderCreateResponse.ResponseData.Status, "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            #endregion

            #region Changes DeviceID.
            this.CMDAdapter.ChangeDeviceID("NewDeviceID");
            this.RecordDeviceInfoChanged();
            #endregion

            #region Calls method FolderSync to synchronize the collection hierarchy.
            FolderSyncRequest folderSyncRequestForNewDeviceID = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSyncResponseForNewDeviceID = this.CMDAdapter.FolderSync(folderSyncRequestForNewDeviceID);

            // Gets the server ID for new folder after change DeviceID.
            string serverId = TestSuiteBase.GetCollectionId(folderSyncResponseForNewDeviceID, folderName);
            Site.Assert.IsNotNull(serverId, "Call method GetServerId to get a non-null ServerId to indicate success.");
            #endregion

            #region The client calls FolderDelete command to delete the created folder in step 2 from the server.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(folderSyncResponseForNewDeviceID.ResponseData.SyncKey, serverId);
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);
            Site.Assert.AreEqual<byte>((byte)1, folderDeleteResponse.ResponseData.Status, "The server should return a status code 1 in the FolderDelete command response to indicate success.");
            #endregion

            #region Restore DeviceID and call FolderSync command
            this.CMDAdapter.ChangeDeviceID(Common.GetConfigurationPropertyValue("DeviceID", this.Site));

            // The client calls FolderSync command to synchronize the collection hierarchy with original device id.
            FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest(folderCreateResponse.ResponseData.SyncKey);
            FolderSyncResponse folderSyncResponse = this.CMDAdapter.FolderSync(folderSyncRequest);
            Site.Assert.AreEqual<byte>(1, folderSyncResponse.ResponseData.Status, "Server should return status 1 in the FolderSync response to indicate success.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "Server should return Changes element in the FolderSync response after the collection hierarchy changed by call FolderDelete command.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes.Delete, "Server should return Changes element in the FolderSync response after the specified folder deleted.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5863");

            // The folderDeleteSuccess is true indicates the folder which deleted by FolderDelete command is deleted successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5863
            Site.CaptureRequirementIfIsNotNull(
                folderSyncResponse.ResponseData.Changes.Delete[0].ServerId,
                5863,
                @"[In Delete(FolderSync)] [The Delete element] specifies that a folder on the server was deleted since the last folder synchronization.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify FolderSync command synchronizes the folder hierarchy successfully, but does not synchronize the items in the collections themselves.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S04_TC10_FolderSync_NoSynchronizeItems()
        {
            #region User2 calls method SendMail to send MIME-formatted e-mail messages to user1.
            this.SwitchUser(this.User2Information);
            string subject = Common.GenerateResourceName(Site, "subject");
            SendMailResponse responseSendMail = this.SendPlainTextEmail(null, subject, this.User2Information.UserName, this.User1Information.UserName, null);
            Site.Assert.AreEqual(string.Empty, responseSendMail.ResponseDataXML, "If SendMail command executes successfully, server should return empty xml data");
            #endregion

            #region Switch to user1 mailbox and call FolderSync command
            this.SwitchUser(this.User1Information);
            this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);

            // Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5775");
            bool isVerifyR5775 = true;

            // FolderSync command does not synchronize the new email items, so FolderChangesAdd does not contain new email items.
            foreach (FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.DisplayName == subject && add.ParentId == this.User1Information.InboxCollectionId)
                {
                    isVerifyR5775 = false;
                    break;
                }
            }

            // Verify MS-ASCMD requirement: MS-ASCMD_R5775
            Site.CaptureRequirementIfIsTrue(
                isVerifyR5775,
                5775,
                @"[In FolderSync] But [FolderSync command] does not synchronize the items in the collections themselves.");
            #endregion
        }
        #endregion
    }
}