namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    /// <summary>
    /// This scenario is designed to test the FolderCreate command.
    /// </summary>
    [TestClass]
    public class S02_FolderCreate : TestSuiteBase
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
        /// This test case is used to verify if the FolderCreate command request is successful, ServerId element should be returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC01_FolderCreate_Success()
        {
            #region Call method FolderCreate to create a new folder as a child folder of the specified parent folder.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "If the FolderCreate command executes successfully, the Status in response should be 1.");
           
            // Record created folder collectionID.
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            bool folderAddSuccess = false;
            foreach (FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.ServerId == folderCreateResponse.ResponseData.ServerId)
                {
                    folderAddSuccess = true;
                    break;
                }
            }
            #endregion

            #region Capture FolderCreate command success related requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4011");

            // If the serverId of Add element equal with the serverId specified in the response of FolderCreate command, it indicates the FolderCreate command completed successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4011
            Site.CaptureRequirementIfIsTrue(
                folderAddSuccess,
                4011,
                @"[In Status(FolderCreate)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R79");

            // If the serverId of Add element equal with the serverId specified in the response of FolderCreate command, it indicates the FolderCreate command completed successfully.
            // Verify MS-ASCMD requirement: MS-ASCMD_R79
            Site.CaptureRequirementIfIsTrue(
                folderAddSuccess,
                79,
                @"[In FolderCreate] The FolderCreate command creates a new folder as a child folder of the specified parent folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3903");

            // Folder has been created successfully, server returns a non-null ServerId.
            Site.CaptureRequirementIfIsNotNull(
                folderCreateResponse.ResponseData.ServerId,
                3903,
                @"[In ServerId(FolderCreate)] The ServerId element MUST be returned if the FolderCreate command request was successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4569");

            // Folder has been created successfully, server must send a synchronization key to the client in a response.
            Site.CaptureRequirementIfIsNotNull(
                folderCreateResponse.ResponseData.SyncKey,
                4569,
                @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] After a successful FolderCreate command (section 2.2.2.2) [, FolderDelete command (section 2.2.2.3), or FolderUpdate command (section 2.2.2.5)], the server MUST send a synchronization key to the client in a response.");
            #endregion
        }
        
        /// <summary>
        /// This test case is used to verify if the FolderCreate command request fails, ServerId element should not be returned.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC02_FolderCreate_Fail()
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder without DisplayName element value.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, string.Empty, "0");
            Site.Assert.IsNotNull(folderCreateResponse.ResponseData.Status, "The Status element should be return.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3904");

            // If the FolderCreate command request fails, server returns a null ServerId.
            Site.CaptureRequirementIfIsNull(
                folderCreateResponse.ResponseData.ServerId,
                3904,
                @"[In ServerId(FolderCreate)] the element MUST NOT be returned if the FolderCreate command request fails.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4570");

            // If the FolderCreate command request fails, server returns a null SyncKey.
            Site.CaptureRequirementIfIsNull(
                folderCreateResponse.ResponseData.SyncKey,
                4570,
                @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] If the FolderCreate command [, FolderDelete command, or FolderUpdate command] is not successful, the server MUST NOT return a SyncKey element.");
        }
        
        /// <summary>
        /// This test case is used to verify FolderCreate command, if the folder name already exists, the status should be equal to 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC03_FolderCreate_Status2()
        {
            #region Call method FolderCreate to create a new folder as a child folder of the mailbox Root folder.
            string folderName = Common.GenerateResourceName(Site, "FolderCreate", 1);
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            FolderCreateResponse folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "If the FolderCreate command executes successfully, the Status in response should be 1.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Call method FolderCreate to create another new folder with same name as a child folder of the mailbox Root folder.
            folderCreateRequest = Common.CreateFolderCreateRequest(folderCreateResponse.ResponseData.SyncKey, (byte)FolderType.UserCreatedMail, folderName, "0");
            folderCreateResponse = this.CMDAdapter.FolderCreate(folderCreateRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4013");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4013
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4013,
                @"[In Status(FolderCreate)] [When the scope is Item], [the cause of the status value 2 is] The parent folder already contains a folder that has this name.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify FolderCreate command, if the parentId doesn't exist, the status should be equal to 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC04_FolderCreate_Status5()
        {
            // Set a parentFolderID that doesn't exist in request.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "InvalidParentId");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4019");

            // If the parent folder does not exist on the server, the value of the Status element should be 5.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4019
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4019,
                @"[In Status(FolderCreate)] [When the scope is Item], [the cause of the status value 5 is] The parent folder does not exist on the server, possibly because it has been deleted or renamed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4018");

            // If the parent folder does not exist on the server, the value of the Status element should be 5.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4018
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4018,
                @"[In Status(FolderCreate)] [When the scope is] Item, [the meaning of the status value] 5 [is] The specified parent folder was not found.");
        }
        
        /// <summary>
        /// This test case is used to verify FolderCreate command, if the SyncKey is invalid, the status should be equal to 9.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC05_FolderCreate_Status9()
        {
            // Call method FolderCreate to create a new folder with invalid folder SyncKey.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse("InvalidFolderSyncKey", (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4032");

            // If the client sent a malformed or mismatched synchronization key in FolderCreate request, the value of the Status element should be 9.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4032
            Site.CaptureRequirementIfAreEqual<int>(
                9,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4032,
                @"[In Status(FolderCreate)] [When the scope is Global], [the cause of the status value 9 is] The client sent a malformed or mismatched synchronization key, or the synchronization state is corrupted on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4573");

            // If the client sent a malformed or mismatched synchronization key in FolderCreate request, the value of the Status element should be 9.
            // Verify MS-ASCMD requirement: MS-ASCMD_R4573
            Site.CaptureRequirementIfAreEqual<int>(
                9,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4573,
                @"[In SyncKey(FolderCreate, FolderDelete, and FolderUpdate)] The server MUST return a Status element (section 2.2.3.162.4) value of 9 if the value of the SyncKey element does not match the value of the synchronization key on the server.");
        }
        
        /// <summary>
        /// This test case is used to verify FolderCreate command, if the request contains a semantic error, the status should be equal to 10.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC06_FolderCreate_Status10()
        {
            // Call method FolderCreate to create a new folder as a child folder of the specified parent folder without folder SyncKey.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(null, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), "0");
            Site.Assert.AreEqual<int>(10, int.Parse(folderCreateResponse.ResponseData.Status), "If the request contains a semantic error, the status should be equal to 10.");
            folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.Inbox, Common.GenerateResourceName(this.Site, "FolderCreate"), "0");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4035");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4035
            Site.CaptureRequirementIfAreEqual<int>(
                10,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4035,
                @"[In Status(FolderCreate)] [When the scope is Global], [the cause of the status value 10 is] The client sent a FolderCreate command request (section 2.2.2.2) that contains a semantic error, or the client attempted to create a default folder, such as the Inbox folder, Outbox folder, or Contacts folder.");
        }

        /// <summary>
        /// This test case is used to verify FolderCreate command, if the specified parent folder is the recipient information cache, the status should be equal to 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC07_FolderCreate_Status3()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recipient information cache is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Set the specified parent folder is the recipient information cache
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderCreate"), User1Information.RecipientInformationCacheCollectionId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R4016");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4016
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                int.Parse(folderCreateResponse.ResponseData.Status),
                4016,
                @"[In Status(FolderCreate)] [When the scope is Item], [the cause of the status value 3 is] The specified parent folder is the Recipient information folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R5769");

            // Since server return status 3, the FolderCreate command cannot be used to update a recipient information cache.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5769
            Site.CaptureRequirement(
                5769,
                @"[In FolderCreate] The FolderCreate command cannot be used to create [a recipient information cache or] a subfolder of a recipient information cache.");
        }

        /// <summary>
        /// This test case is used to verify FolderCreate command, if create a recipient information cache, the status should be equal to 3.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S02_TC08_FolderCreate_RecipientInformationCache_Status3()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recipient information cache is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            // Set the specified parent folder is the recipient information cache
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.RecipientInformationCache, Common.GenerateResourceName(Site, "FolderCreate"), User1Information.InboxCollectionId);

            Site.Assert.AreEqual<int>(3, int.Parse(folderCreateResponse.ResponseData.Status), "The status should be equal to 3 when the FolderCreate is used to create a recipient information cache.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "MS-ASCMD_R81");

            // Since server return status 3, the FolderCreate command cannot be used to update a recipient information cache.
            // Verify MS-ASCMD requirement: MS-ASCMD_R81
            Site.CaptureRequirement(
                81,
                @"[In FolderCreate] The FolderCreate command cannot be used to create a recipient information cache [or a subfolder of a recipient information cache].");
        }
        #endregion
    }
}