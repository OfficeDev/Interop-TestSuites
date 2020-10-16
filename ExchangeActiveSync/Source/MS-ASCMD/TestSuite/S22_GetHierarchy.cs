namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the GetHierarchy command.
    /// </summary>
    [TestClass]
    public class S22_GetHierarchy : TestSuiteBase
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

        /// <summary>
        /// Verify the requirement about GetHierarchy command.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S22_TC01_GetHierarchySuccess()
        {
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetHierarchy command is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetHierarchy command is not supported when the MS-ASProtocolVersion header is set to 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetHierarchy command is not supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The GetHierarchy command is not supported when the MS-ASProtocolVersion header is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region The client calls FolderCreate command to create a new folder as a child folder of the specified parent folder, then server returns ServerId for FolderCreate command.
            FolderCreateResponse folderCreateResponse = this.GetFolderCreateResponse(this.LastFolderSyncKey, (byte)FolderType.UserCreatedMail, Common.GenerateResourceName(Site, "FolderSync"), "0");
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderCreate command response to indicate success.");
            TestSuiteBase.RecordCaseRelativeFolders(this.User1Information, folderCreateResponse.ResponseData.ServerId);
            #endregion

            #region Call method FolderSync to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.FolderSync();
            Site.Assert.AreEqual<int>(1, int.Parse(folderCreateResponse.ResponseData.Status), "The server should return a status code 1 in the FolderSync command response to indicate success.");
            string sentItemFolderCollectionId = string.Empty;
            string deleteItemFolderCollectionId = string.Empty;

            foreach (Response.FolderSyncChangesAdd folderAdd in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (string.Compare(folderAdd.DisplayName, "Sent Items", true) == 0)
                {
                    sentItemFolderCollectionId = folderAdd.ServerId;
                }

                if (string.Compare(folderAdd.DisplayName, "Deleted Items", true) == 0)
                {
                    deleteItemFolderCollectionId = folderAdd.ServerId;
                }
            }
            #endregion

            #region Call method GetHierarchy to get the list of email folders from the server.
            GetHierarchyResponse getHierarchyResponse = this.CMDAdapter.GetHierarchy();

            bool isVerifiedR7505 = false;
            bool isVerifiedR7507 = false;
            foreach (Response.FoldersFolder folder in getHierarchyResponse.ResponseData.Folder)
            {
                if (folder.DisplayName.Equals("Sent Items"))
                {
                     isVerifiedR7505 = true;
                }

                if (folder.DisplayName.Equals("Deleted Items"))
                {
                    isVerifiedR7507 = true;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR7505,
                7505,
                @"[In GetHierarchy] The client can use the GetHierarchy command to obtain the collection ID of a folder, such as Sent Items folder [or Deleted Items folder], that cannot be deleted.");

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR7507,
                7507,
                @"[In GetHierarchy] The client can use the GetHierarchy command to obtain the collection ID of a folder, such as [Sent Items folder or] Deleted Items folder, that cannot be deleted.");

            // If R6030 have been verfied and sentItemFolderCollectionId is not null , then the client can obtain the collection ID of folder from ServerId element of previous FOlderSync 
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(sentItemFolderCollectionId) == false,
                6031,
                @"[In GetHierarchy] The collection ID is obtained from the ServerId element of a previous FolderSync [or FolderCreate] command.");

            bool isVerifiedR6032 = string.IsNullOrEmpty(folderCreateResponse.ResponseData.ServerId) == false;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR6032,
                6032,
                @"[In GetHierarchy] The collection ID is obtained from the ServerId element of a previous [FolderSync or] FolderCreate command.");

            bool isVerifiedR6025 = false;

            foreach (Response.FoldersFolder folder in getHierarchyResponse.ResponseData.Folder)
            {
                if (!string.IsNullOrEmpty(folder.ParentId))
                {
                    isVerifiedR6025 = true;
                }
                else
                {
                    isVerifiedR6025 = false;
                    break;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR6025,
                6025,
                @"[In GetHierarchy] Each folder's place within the folder hierarchy is indicated by its parent ID.");

            bool isVerifiedR6026 = false;
            foreach (Response.FoldersFolder folder in getHierarchyResponse.ResponseData.Folder)
            {
                // According Open Specification, if the type of the folder is not 1, 2, 3, 4, 5 and 6 then this folder is not a email folder.
                if (folder.Type == 1 || folder.Type == 2 || folder.Type == 3 || folder.Type == 4 || folder.Type == 5 || folder.Type == 6)
                {
                    isVerifiedR6026 = true;
                }
                else
                {
                    isVerifiedR6026 = false;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR6026,
                6026,
                @"[In GetHierarchy] The list of folders returned by the GetHierarchy command includes only email folders.");

            // If above requirements have been verified, the R6024 will be verified.
            this.Site.CaptureRequirement(
                6024,
                @"[In GetHierarchy] The GetHierarchy command gets the list of email folders from the server.");
            #endregion
        }
    }
}
