namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the FolderDelete command.
    /// </summary>
    [TestClass]
    public class S03_FolderDelete : TestSuiteBase
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
        /// This test case is used to verify if the FolderDelete command is successful, the status should be equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S03_TC01_FolderDelete_Success()
        {
            #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderDelete"), "0");
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderCreateResponse.ResponseData.Status),
                "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            #endregion

            #region Call method FolderDelete to delete the created folder from the server.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId);
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                "The server should return a status code 1 in the FolderDelete command response to indicate success.");
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderSyncResponse.ResponseData.Status),
                "The server should return a status code 1 in the FolderSync command response to indicate success.");

            bool folderDeleteSuccess = true;
            foreach (FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.ServerId == folderCreateResponse.ResponseData.ServerId)
                {
                    folderDeleteSuccess = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R97");

            // The folderDeleteSuccess is true indicates the folder which deleted by FolderDelete command is deleted successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R97
            Site.CaptureRequirementIfIsTrue(
                folderDeleteSuccess,
                97,
                @"[In FolderDelete] The FolderDelete command deletes a folder from the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R98");

            // The folderDeleteSuccess is true indicates the folder which deleted by FolderDelete command is deleted successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R98
            Site.CaptureRequirementIfIsTrue(
                folderDeleteSuccess,
                98,
                @"[In FolderDelete] The ServerId (section 2.2.3.151.2) of the folder is passed to the server in the FolderDelete command request (section 2.2.2.3), which deletes the collection with the matching identifier.");

            // The folderDeleteSuccess is true indicates the folder which deleted by FolderDelete command is deleted successfully.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4045");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4045
            Site.CaptureRequirementIfIsTrue(
                folderDeleteSuccess,
                4045,
                @"[In Status(FolderDelete)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5783");

            // If folder has been deleted successfully, the server must send a synchronization key to the client in a response.
            Site.CaptureRequirementIfIsNotNull(
               folderDeleteResponse.ResponseData.SyncKey,
               5783,
               @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] After a successful [FolderCreate command (section 2.2.2.2),] FolderDelete command (section 2.2.2.3) [, or FolderUpdate command (section 2.2.2.5)], the server MUST send a synchronization key to the client in a response.");
            #endregion
        }
        
        /// <summary>
        /// This test case is used to verify FolderDelete command, if the specified folder is a special system folder, the status in return value is 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S03_TC02_FolderDelete_Status3()
        {
            // Call method FolderDelete to delete a Calendar folder from the server.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(this.LastFolderSyncKey,this.User1Information.CalendarCollectionId);
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);

            Site.Assert.IsNotNull(folderDeleteResponse.ResponseData, "The FolderDelete element should not be null.");
            Site.Assert.IsNotNull(folderDeleteResponse.ResponseData.Status, "As child element of FolderDelete, the Status should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R99");

            // Since the FolderDelete and Status element are not null, server sends a response indicating the status of the deletion.
            // Verify MS-ASCMD requirement: MS-ASCMD_R99
            Site.CaptureRequirement(
                99,
                @"[In FolderDelete] The server then sends a response indicating the status of the deletion.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R101");

            // The server should return a status code 3 in the FolderDelete command response to indicate the specified folder is a special folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R101
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                101,
                @"[In FolderDelete] Attempting to delete a recipient information cache using this command results in a Status element (section 2.2.3.162.3) value of 3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4046");

            // The server should return a status code 3 in the FolderDelete command response to indicate the specified folder is a special folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4046
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4046,
                @"[In Status(FolderDelete)] [When the scope is] Item, [the meaning of the status value] 3 [is] The specified folder is a special system folder, such as the Inbox folder, Outbox folder, Contacts folder, Recipient information, or Drafts folder, and cannot be deleted by the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4047");

            // The server should return a status code 3 in the FolderDelete command response to indicate the specified folder is a special folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4047
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4047,
                @"[In Status(FolderDelete)] [When the scope is Item], [the cause of the status value 3 is] The client specified a special folder in a FolderDelete command request (section 2.2.2.3). special folders cannot be deleted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4040");

            // The server should return a status code 3 in the FolderDelete command response to indicate the specified folder is a special folder.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4040
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4040,
                @"[In Status(FolderDelete)] If the command failed, the Status element in the server response contains a code indicating the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5785");

            // If the FolderDelete command request fails, server returns a null SyncKey.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5785
            Site.CaptureRequirementIfIsNull(
                folderDeleteResponse.ResponseData.SyncKey,
                5785,
                @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] If the [FolderCreate command,] FolderDelete command [, or FolderUpdate command] is not successful, the server MUST NOT return a SyncKey element.");

            // The recipient information cache is not supported when the value of the MS-ASProtocolVersion header is set to 12.1. 
            // MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.
            if ("12.1" != Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site))
            {
                // Call method FolderDelete to delete a recipient information cache from the server.
                folderDeleteRequest = Common.CreateFolderDeleteRequest(this.LastFolderSyncKey, "RI");
                folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);

                Site.Assert.IsNotNull(folderDeleteResponse.ResponseData, "The FolderDelete element should not be null.");
                Site.Assert.IsNotNull(folderDeleteResponse.ResponseData.Status, "As child element of FolderDelete, the Status should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R100");

                // The server should return a status code 3 in the FolderDelete command response to indicate the specified folder is a special folder.
                // Verify MS-ASCMD requirement: MS-ASCMD_R100
                Site.CaptureRequirementIfAreEqual<int>(
                    3,
                    int.Parse(folderDeleteResponse.ResponseData.Status),
                    100,
                    @"[In FolderDelete] The FolderDelete command cannot be used to delete a recipient information cache.");
            }
        }
        
        /// <summary>
        /// This test case is used to verify FolderDelete command, if the specified folder does not exist, the status in return value is 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S03_TC03_FolderDelete_Status4()
        {
            // Call method FolderDelete to delete an invalid folder from the server.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(this.LastFolderSyncKey, "InvalidServerId");
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4048");

            // The server should return a status code 4 in the FolderDelete command response to indicate the specified folder does not exist.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4048
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4048,
                @"[In Status(FolderDelete)] [When the scope is] Item, [the meaning of the status value] 4 [is] The specified folder does not exist.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4049");

            // The server should return a status code 4 in the FolderDelete command response to indicate the specified folder does not exist.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4049
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4049,
                @"[In Status(FolderDelete)] [When the scope is Item], [the cause of the status value 4 is] The client specified a nonexistent folder in a FolderDelete command request.");
        }
        
        /// <summary>
        /// This test case is used to verify FolderDelete command, if the SyncKey is an empty string, the status in return value is 9.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S03_TC04_FolderDelete_Status9()
        {
            #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderDelete"), "0");
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(folderCreateResponse.ResponseData.Status),
                "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            #endregion

            #region Call method FolderDelete to delete a folder from the server, and set SyncKey value to an empty string.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(string.Empty, folderCreateResponse.ResponseData.ServerId);
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4060");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4060
            Site.CaptureRequirementIfAreEqual<int>(
                9,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4060,
                @"[In Status(FolderDelete)] [When the scope is Global], [the cause of the status value 9 is] The client sent a malformed or mismatched synchronization key [, or the synchronization state is corrupted on the server].");
            #endregion

            #region Call method FolderDelete to delete the created folder from the server.
            folderDeleteRequest = Common.CreateFolderDeleteRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId);
            folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderDeleteResponse.ResponseData.Status), "The created Folder should be deleted.");
            #endregion
        }
        
        /// <summary>
        /// This test case is used to verify FolderDelete command, if the request contains a semantic or syntactic error, the status in return value is 10.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S03_TC05_FolderDelete_Status10()
        {
            #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderDelete"), "0");
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "If the FolderCreate command creates a folder successfully, server should return a status code 1.");
            #endregion

            #region Call method FolderDelete without folder SyncKey to delete a folder from the server.
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(null, folderCreateResponse.ResponseData.ServerId);
            FolderDeleteResponse folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4063");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4063
            Site.CaptureRequirementIfAreEqual<int>(
                10,
                int.Parse(folderDeleteResponse.ResponseData.Status),
                4063,
                @"[In Status(FolderDelete)] [When the scope is Global], [the cause of the status value 10 is] The client sent a FolderCreate command request (section 2.2.2.3) that contains a semantic or syntactic error.");
            #endregion

            #region Call method FolderDelete to delete the created folder from the server.
            folderDeleteRequest = Common.CreateFolderDeleteRequest(folderCreateResponse.ResponseData.SyncKey, folderCreateResponse.ResponseData.ServerId);
            folderDeleteResponse = this.CMDAdapter.FolderDelete(folderDeleteRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderDeleteResponse.ResponseData.Status), "The server should return a status code 1 in the FolderDelete command response to indicate success.");
            #endregion
        }
        #endregion
    }
}