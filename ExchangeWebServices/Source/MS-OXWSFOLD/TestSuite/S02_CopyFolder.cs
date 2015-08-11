namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify CopyFolder operation.
    /// </summary>
    [TestClass]
    public class S02_CopyFolder : TestSuiteBase
    {
        #region Class initialize and clean up

        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases
        /// <summary>
        /// This test case verifies requirements related to copying folder operation via copy Drafts folder into Inbox folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S02_TC01_CopyFolder()
        {
            #region Copy the "drafts" folder to the inbox folder

            // Identify the folders to be copied.
            DistinguishedFolderIdType copiedFolderId = new DistinguishedFolderIdType();
            copiedFolderId.Id = DistinguishedFolderIdNameType.drafts;

            // CopyFolder request.
            CopyFolderType copyFolderRequest = this.GetCopyFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), copiedFolderId);

            // Copy the "drafts" folder.
            CopyFolderResponseType copyFolderResponse = this.FOLDAdapter.CopyFolder(copyFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderIdType folderId = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the copied folder's folder id.
            this.NewCreatedFolderIds.Add(folderId);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1852");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1852
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                copyFolderResponse.ResponseMessages.Items[0].ResponseClass,
                1852,
                @"[In CopyFolder Operation]A successful CopyFolder operation request returns a CopyFolderResponse element with the ResponseClass attribute of the CopyFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R185222");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R185222
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                copyFolderResponse.ResponseMessages.Items[0].ResponseCode,
                185222,
                @"[In CopyFolder Operation]A successful CopyFolder operation request returns a CopyFolderResponse element with the ResponseCode element of the CopyFolderResponse element set to ""NoError"".");

            #region Get the new copied folder

            // GetFolder request.
            GetFolderType getSubFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderId);

            // Get the specified folder.
            GetFolderResponseType getSubFolderResponse = this.FOLDAdapter.GetFolder(getSubFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getSubFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R211");

            // The copied folder can be gotten successfully through returned folder id, so the draft folder was copied.
            this.Site.CaptureRequirement(
                211,
                @"[In m:CopyFolderType Complex Type]The CopyFolderType complex type specifies a request message to copy folders in a server database.");
        }

        /// <summary>
        /// This test case verifies requirements related to copying multiple folders to Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S02_TC02_CopyMultipleFolders()
        {
            #region Copy the "drafts" and "deleteditems" folder to inbox

            // Set "drafts" and "deleteditems" folders' Id.
            DistinguishedFolderIdType copiedFolderId1 = new DistinguishedFolderIdType();
            copiedFolderId1.Id = DistinguishedFolderIdNameType.drafts;
            DistinguishedFolderIdType copiedFolderId2 = new DistinguishedFolderIdType();
            copiedFolderId2.Id = DistinguishedFolderIdNameType.deleteditems;

            // CopyFolder request.
            CopyFolderType copyFolderRequest = this.GetCopyFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), copiedFolderId1, copiedFolderId2);

            // Copy the "drafts" and "deleteditems" folder.
            CopyFolderResponseType copyFolderResponse = this.FOLDAdapter.CopyFolder(copyFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyFolderResponse, 2, this.Site);

            // Save copied folders' id.
            FolderIdType folderId1 = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            FolderIdType folderId2 = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[1]).Folders[0].FolderId;

            // Save the copied folders' id.
            this.NewCreatedFolderIds.Add(folderId1);
            this.NewCreatedFolderIds.Add(folderId2);

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to copying folder to and from a public folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S02_TC03_CopyPublicFolder()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(55501, this.Site), "Exchange 2007 and Exchange 2010 support the CopyFolder operation if either the source folder or the destination folder is a public folder");

            #region Create a new public folder in the public folder root

            // CreateFolder request.
            CreateFolderType createPublicFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.publicfoldersroot.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new public folder.
            CreateFolderResponseType createPublicFolderResponse = this.FOLDAdapter.CreateFolder(createPublicFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createPublicFolderResponse, 1, this.Site);

            // Save the new created public folder's folder id.
            FolderIdType newPublicFolderId = ((FolderInfoResponseMessageType)createPublicFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newPublicFolderId);

            #endregion

            #region Create a new folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Copy the public folder to the folder created in inbox

            // CopyFolder request.
            CopyFolderType copyPublicFolderRequest = this.GetCopyFolderRequest(newFolderId.Id, newPublicFolderId);

            // Copy the public folder.
            CopyFolderResponseType copyPublicFolderResponse = this.FOLDAdapter.CopyFolder(copyPublicFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyPublicFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderIdType copiedPublicFolderId = ((FolderInfoResponseMessageType)copyPublicFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the copied folder's folder id.
            this.NewCreatedFolderIds.Add(copiedPublicFolderId);

            #endregion

            #region Get the new copied public folder that in inbox

            // GetFolder request.
            GetFolderType getNewCopiedPulicFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, copiedPublicFolderId);

            // Get the new copied public folder.
            GetFolderResponseType getNewCopiedPublicFolderResponse = this.FOLDAdapter.GetFolder(getNewCopiedPulicFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getNewCopiedPublicFolderResponse, 1, this.Site);

            #endregion

            #region Copy the folder in inbox to the public folder created

            // CopyFolder request.
            CopyFolderType copyFolderRequest = this.GetCopyFolderRequest(newPublicFolderId.Id, newFolderId);

            // Copy the public folder.
            CopyFolderResponseType copyFolderResponse = this.FOLDAdapter.CopyFolder(copyFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderIdType copiedFolderId = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the copied folder's folder id.
            this.NewCreatedFolderIds.Add(copiedFolderId);

            #endregion

            #region Get the new copied folder that in root public folder

            // GetFolder request.
            GetFolderType getNewCopiedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, copiedFolderId);

            // Get the new copied folder.
            GetFolderResponseType getNewCopiedFolderResponse = this.FOLDAdapter.GetFolder(getNewCopiedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getNewCopiedFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R55501");

            // Folder can be copied successfully either when the source folder or destination folder is public folder ,so this requirement can be captured.
            this.Site.CaptureRequirement(
                55501,
                @"[In Appendix C: Product Behavior] Implementation does support the CopyFolder operation if either the source folder or the destination folder is a public folder.(Exchange Server 2007 and Exchange Server 2010 follow this behavior.)");
        }
        #endregion
    }
}