namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify EmptyFolder operation.
    /// </summary>
    [TestClass]
    public class S05_EmptyFolder : TestSuiteBase
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
        /// This test case verifies requirements related to EmptyFolder including deleting sub folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S05_TC01_EmptyFolder()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5664, this.Site), "Exchange Server 2007 and the initial release version of Exchange Server 2010 do not support EmptyFolder operation");

            #region Create a new item and a new folder with an item in the Inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "ToBeDeleteFolder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            string itemName1 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemInFolder = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName1);
            Site.Assert.IsNotNull(itemInFolder, "Item should be created successfully!");

            string itemName2 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in inbox.
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), DistinguishedFolderIdNameType.inbox.ToString(), itemName2);
            Site.Assert.IsNotNull(itemId, "Item should be created successfully!");

            // Variable to indicate whether the item2 is created properly.
            bool isItem2Created = this.FindItem(DistinguishedFolderIdNameType.inbox.ToString(), itemName2) != null;

            Site.Assert.IsTrue(isItem2Created, "The item should be created successfully in the specific folder.");

            #endregion

            #region Empty the inbox folder

            // Specify which folder will be emptied.
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;

            // Empty the specific folder
            EmptyFolderResponseType emptyFolderResponse = this.CallEmptyFolderOperation(folderId, DisposalType.HardDelete, true);

            // Check the response.
            Common.CheckOperationSuccess(emptyFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3474");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3474
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                emptyFolderResponse.ResponseMessages.Items[0].ResponseClass,
                3474,
                @"[In EmptyFolder Operation]A successful EmptyFolder operation request returns an EmptyFolderResponse element with the ResponseClass attribute of the EmptyFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R34744");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R34744
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                emptyFolderResponse.ResponseMessages.Items[0].ResponseCode,
                34744,
                @"[In EmptyFolder Operation]A successful EmptyFolder operation request returns an EmptyFolderResponse element with the ResponseCode element of the EmptyFolderResponse element set to ""NoError"".");

            #region Find the item in inbox to see whether it has been deleted

            // Verify if item under inbox exists.
            bool isInboxItemDeleted = this.IfItemDeleted(DistinguishedFolderIdNameType.inbox.ToString(), itemName2);

            #endregion

            #region Get the folder in inbox folder to verify whether it has been deleted

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specific folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, getFolderResponse.ResponseMessages.Items[0].ResponseClass, "Folder information should not be returned! ");

            // Variable to indicate whether the folder in inbox folder is deleted.
            bool isFolderDeleted = getFolderResponse.ResponseMessages.Items[0].ResponseCode == ResponseCodeType.ErrorItemNotFound;
            bool isItemDeleted = isInboxItemDeleted && this.IfItemDeleted(DistinguishedFolderIdNameType.inbox.ToString(), itemName1);

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R367");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R367.
            bool isVerifyR367 = isItemDeleted && isFolderDeleted;

            Site.Assert.IsTrue(
                isVerifyR367,
                "The expected result of deleting item should be \"true\", actual result is {0};\n" +
                "the expected result of deleting folder should be \"true\", actual result is {1}.\n ",
                isItemDeleted,
                isFolderDeleted);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR367,
                367,
                @"[In Elements]EmptyFolder specifies a request to empty folders in a mailbox in the server store.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R571");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R571.
            bool isVerifyR571 = isItemDeleted && isFolderDeleted;

            Site.Assert.IsTrue(
                isVerifyR571,
                "The expected result of deleting item should be \"true\", actual result is {0};\n" +
                "the expected result of deleting folder should be \"true\", actual result is {1}.\n ",
                isItemDeleted,
                isFolderDeleted);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR571,
                571,
                @"[In m:EmptyFolderType Complex Type]The EmptyFolderType complex type specifies a request message to empty a folder in a mailbox.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R380");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R380.
            // DeleteSubFolders has been set as true, if the subfolder is deleted, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isFolderDeleted,
                380,
                @"[In m:EmptyFolderType Complex Type]The DeleteSubFolders attribute is set to ""true"" if the subfolders are to be deleted.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5664");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5664
            bool isVerifiedR5664 = isItemDeleted && isFolderDeleted;

            Site.Assert.IsTrue(
               isVerifiedR5664,
               "The expected result of deleting item should be \"true\", actual result is {0};\n" +
               "the expected result of deleting folder should be \"true\", actual result is {1}.\n ",
               isItemDeleted,
               isFolderDeleted);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR5664,
                5664,
                @"[In Appendix C: Product Behavior] Implementation does include the EmptyFolder operation.(Exchange Server 2010 SP2 and above follow this behavior.)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R37801");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R37801
            bool isVerifiedR37801 = isItemDeleted && isFolderDeleted;

            Site.Assert.IsTrue(
               isVerifiedR37801,
               "The expected result of deleting item should be \"true\", actual result is {0};\n" +
               "the expected result of deleting folder should be \"true\", actual result is {1}.\n ",
               isItemDeleted,
               isFolderDeleted);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR37801,
                37801,
                @"[In m:EmptyFolderType Complex Type ]DeleteType which value is HardDelete specifies that an item or folder is permanently removed from the store.");
        }

        /// <summary>
        /// This test case verifies the requirements related to EmptyFolder without deleting its sub folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S05_TC02_EmptyFolderWithoutDeletingSubFolder()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5664, this.Site), "Exchange Server 2007 and the initial release version of Exchange Server 2010 do not support EmptyFolder operation");

            #region Create a new item and a new folder with an item in the Inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "ToBeDeleteFolder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);

            string itemName1 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemInFolder = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName1);
            Site.Assert.IsNotNull(itemInFolder, "Item should be created successfully!");
            string itemName2 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in inbox.
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), DistinguishedFolderIdNameType.inbox.ToString(), itemName2);
            Site.Assert.IsNotNull(itemId, "Item should be created successfully!");

            // Variable to indicate whether the item2 is created properly.
            bool isItem2Created = this.FindItem(DistinguishedFolderIdNameType.inbox.ToString(), itemName2) != null;

            Site.Assert.IsTrue(isItem2Created, "The item should be created successfully in the specific folder.");

            #endregion

            #region Empty the inbox folder.

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;

            // Empty the specific folder
            EmptyFolderResponseType emptyFolderResponse = this.CallEmptyFolderOperation(folderId, DisposalType.HardDelete, false);

            // Check the response.
            Common.CheckOperationSuccess(emptyFolderResponse, 1, this.Site);

            #endregion

            #region Get the folder in inbox folder to verify whether it has been deleted

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specific folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R381");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R381.
            Site.CaptureRequirementIfAreNotEqual<ResponseCodeType>(
                ResponseCodeType.ErrorItemNotFound,
                getFolderResponse.ResponseMessages.Items[0].ResponseCode,
                381,
                @"[In m:EmptyFolderType Complex Type][if the subfolders are not to be deleted], it is set to ""false"". ");
        }

        /// <summary>
        /// This test case verifies requirements related to EmptyFolder with disposal type set to "MoveToDeletedItems".
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S05_TC03_EmptyFolderMoveToDeletedItems()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5664, this.Site), "Exchange Server 2007 and the initial release version of Exchange Server 2010 do not support EmptyFolder operation");

            #region Create a new item and a new folder with an item in the Inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "ToBeDeleteFolder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            this.NewCreatedFolderIds.Add(newFolderId);

            string itemName1 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemInFolder = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName1);
            Site.Assert.IsNotNull(itemInFolder, "Item should be created successfully!");

            string itemName2 = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in inbox.
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), DistinguishedFolderIdNameType.inbox.ToString(), itemName2);
            Site.Assert.IsNotNull(itemId, "Item should be created successfully!");

            // Variable to indicate whether the item2 is created properly.
            bool isItem2Created = this.FindItem(DistinguishedFolderIdNameType.inbox.ToString(), itemName2) != null;

            Site.Assert.IsTrue(isItem2Created, "The item should be created successfully in the specific folder.");

            #endregion

            #region Empty the inbox folder

            // Specify which folder will be emptied.
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;

            // Empty the specific folder
            EmptyFolderResponseType emptyFolderResponse = this.CallEmptyFolderOperation(folderId, DisposalType.MoveToDeletedItems, true);

            // Check the response.
            Common.CheckOperationSuccess(emptyFolderResponse, 1, this.Site);

            #endregion

            #region Get the folder in inbox folder to verify whether it has been deleted

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specific folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            #region Find the item to see whether it has been deleted

            // Verify if item under inbox exists.
            ItemIdType itemIdAfterEmpty = this.FindItem(DistinguishedFolderIdNameType.deleteditems.ToString(), itemName2);
            bool isItemInDeletedItems = itemIdAfterEmpty != null;
            this.NewCreatedItemIds.Add(itemIdAfterEmpty);

            #endregion

            #region Get new created folder's parent folder

            GetFolderType getParentFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId);

            // Get the new created folder.
            GetFolderResponseType getParentFolderResponse = this.FOLDAdapter.GetFolder(getParentFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getParentFolderResponse, 1, this.Site);

            string folderDisplayName = ((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].DisplayName;

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R37802");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R37802
            bool isVerifiedR37802 = folderDisplayName.Equals("Deleted Items") && isItemInDeletedItems;

            Site.Assert.IsTrue(
                isVerifiedR37802,
                "Parent folder name after deleted expected to be \"Deleted Items\" and actual is {0};\n" +
                "Item in deleted items expected to be \"true\" and actual is {1};\n ",
                folderDisplayName,
                isItemInDeletedItems);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR37802,
                37802,
                @"[In m:EmptyFolderType Complex Type ]DeleteType which value is MoveToDeletedItems specifies that an item or folder is moved to the Deleted Items folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to empty folder failed.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S05_TC04_EmptyPublicFolderFailed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5664, this.Site), "Exchange Server 2007 and the initial release version of Exchange Server 2010 do not support EmptyFolder operation");

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

            #region Delete the created folder

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, newFolderId);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);

            // The folder has been deleted, so its folder id has disappeared.
            this.NewCreatedFolderIds.Remove(newFolderId);

            #endregion

            #region Empty the deleted folder
            EmptyFolderResponseType emptyFolderResponse = this.CallEmptyFolderOperation(newFolderId, DisposalType.HardDelete, true);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R34745");

            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                emptyFolderResponse.ResponseMessages.Items[0].ResponseClass,
                34745,
                @"[In EmptyFolder Operation]An unsuccessful EmptyFolder operation request returns an EmptyFolderResponse element with the ResponseClass attribute of the EmptyFolderResponseMessage element set to ""Error"".");
            #endregion
        }
        
        /// <summary>
        /// This test case verifies requirements related to soft empty a folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S05_TC05_SoftEmptyFolder()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5664, this.Site), "Exchange Server 2007 and the initial release version of Exchange Server 2010 do not support EmptyFolder operation");
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1462, this.Site), "Exchange 2007 does not include enumeration value recoverableitemsdeletions");

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

            #region Create an item

            string itemName = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName);
            Site.Assert.IsNotNull(itemId, "Item should be created successfully!");

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            #region Empty the created folder

            EmptyFolderResponseType emptyFolderResponse = this.CallEmptyFolderOperation(newFolderId, DisposalType.SoftDelete, true);
            Common.CheckOperationSuccess(emptyFolderResponse, 1, this.Site);

            #endregion

            #region Find the item
            ItemIdType findItemID = this.FindItem(DistinguishedFolderIdNameType.recoverableitemsdeletions.ToString(), itemName);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R37803");

            this.Site.CaptureRequirementIfIsNotNull(
                findItemID,
                37803,
                @"[In m:EmptyFolderType Complex Type ]DeleteType which value is SoftDelete specifies that an item or folder is moved to the dumpster if the dumpster is enabled.");

            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.ItemIds = new BaseItemIdType[] { findItemID };
            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);
            #endregion
        }
        #endregion
    }
}