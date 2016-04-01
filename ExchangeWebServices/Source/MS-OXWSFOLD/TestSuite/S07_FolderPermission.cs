namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to folder permission.
    /// </summary>
    [TestClass]
    public class S07_FolderPermission : TestSuiteBase
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
        /// This test case verifies requirements related to permission set to custom and all base permissions enabled.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC01_FolderPermissionCustomLevelAllPermissionEnabled()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and enable all permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].ReadItems = new PermissionReadAccessType();
            permissionSet.Permissions[0].ReadItems = PermissionReadAccessType.FullDetails;
            permissionSet.Permissions[0].ReadItemsSpecified = true;
            permissionSet.Permissions[0].CanCreateItems = true;
            permissionSet.Permissions[0].CanCreateItemsSpecified = true;
            permissionSet.Permissions[0].CanCreateSubFolders = true;
            permissionSet.Permissions[0].CanCreateSubFoldersSpecified = true;
            permissionSet.Permissions[0].IsFolderVisible = true;
            permissionSet.Permissions[0].IsFolderVisibleSpecified = true;
            permissionSet.Permissions[0].IsFolderContact = true;
            permissionSet.Permissions[0].IsFolderContactSpecified = true;
            permissionSet.Permissions[0].IsFolderOwner = true;
            permissionSet.Permissions[0].IsFolderOwnerSpecified = true;
            permissionSet.Permissions[0].IsFolderContact = true;
            permissionSet.Permissions[0].IsFolderContactSpecified = true;
            permissionSet.Permissions[0].EditItems = new PermissionActionType();
            permissionSet.Permissions[0].EditItems = PermissionActionType.All;
            permissionSet.Permissions[0].EditItemsSpecified = true;
            permissionSet.Permissions[0].DeleteItems = new PermissionActionType();
            permissionSet.Permissions[0].DeleteItems = PermissionActionType.All;
            permissionSet.Permissions[0].DeleteItemsSpecified = true;
            permissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a subfolder under the folder created in step 1 with User2's credential

            // CreateFolder request.
            CreateFolderType createFolderInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderInSharedMailboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderInSharedMailboxResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4302
            // Successful Creating sub folder indicates that UserId specifies a user identifier
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass,
                4302,
                @"[In t:BasePermissionType Complex Type]UserId specifies a user identifier.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4702");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4702
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass,
                4702,
                @"[In t:BasePermissionType Complex Type][CanCreateSubFolders]A value of true indicates that the client can create a sub folder. ");

            #region Edit items User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;
            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R46");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R46
            // CanCreateItems has been set to true in request and if item can be created this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                this.CanCreateItem,
                46,
                @"[In t:BasePermissionType Complex Type][CanCreateItems]A value of ""true"" indicates that the client can create an item. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R51003");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R51003
            // Value "All" has been set to EditItem, if items that user owns and doesn't own can be edited this requirement can be captured.
            bool isVerifiedR51003 = this.CanEditOwnedItem && this.CanEditNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR51003,
               "Can edit owned item expected to be \"true\" and actual is {0};\n" +
               "Can edit not owned item expected to be \"true\" and actual is {1};\n ",
               this.CanEditOwnedItem,
               this.CanEditNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR51003,
                51003,
                @"[In t:BasePermissionType Complex Type]The type of element EditItems is ""All"", which indicates that the user has permission to perform the action on all items in the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R52003");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R52003
            // Value "All" has been set to DeleteItem, if items that user owns and doesn't own can be deleted this requirement can be captured.
            bool isVerifiedR52003 = this.CanDeleteOwnedItem && this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR52003,
               "Can delete owned item expected to be \"true\" and actual is {0};\n" +
               "Can delete not owned item expected to be \"true\" and actual is {1};\n ",
               this.CanDeleteOwnedItem,
               this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR52003,
                52003,
                @"[In t:BasePermissionType Complex Type]The type of element DeleteItems is ""All"", which indicates that the user has permission to perform the action on all items in the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R98");

            // Permission is set and schema is verified in adapter so this requirement can be captured.
            this.Site.CaptureRequirement(
                98,
                @"[In t:FolderType Complex Type]The type of element PermissionSet is t:PermissionSetType (section 2.2.4.14).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R9802");

            // Permission is set and authorized user get related permissions so this requirement can be captured.
            this.Site.CaptureRequirement(
                9802,
                @"[In t:FolderType Complex Type]PermissionSet specifies all permissions that are configured for a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1182");

            // Permission is set and authorized user get related permissions so this requirement can be captured.
            this.Site.CaptureRequirement(
                1182,
                @"[In t:PermissionSetType Complex Type][Permissions] specifies a collection of permissions for a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R154");

            // Permission is set and authorized user get related permissions so this requirement can be captured.
            this.Site.CaptureRequirement(
                154,
                @"[In t:PermissionLevelType Simple Type]The value Custom means the user has custom access permissions on the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to validate folder permission level set to custom and all base permissions disabled.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC02_FolderPermissionCustomLevelAllPermissionDisabled()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and disable all permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].CanCreateItems = false;
            permissionSet.Permissions[0].CanCreateItemsSpecified = true;
            permissionSet.Permissions[0].CanCreateSubFolders = false;
            permissionSet.Permissions[0].CanCreateSubFoldersSpecified = true;
            permissionSet.Permissions[0].IsFolderVisible = false;
            permissionSet.Permissions[0].IsFolderVisibleSpecified = true;
            permissionSet.Permissions[0].IsFolderContact = false;
            permissionSet.Permissions[0].IsFolderContactSpecified = true;
            permissionSet.Permissions[0].IsFolderOwner = false;
            permissionSet.Permissions[0].IsFolderOwnerSpecified = true;
            permissionSet.Permissions[0].IsFolderContact = false;
            permissionSet.Permissions[0].IsFolderContactSpecified = true;
            permissionSet.Permissions[0].EditItems = new PermissionActionType();
            permissionSet.Permissions[0].EditItems = PermissionActionType.None;
            permissionSet.Permissions[0].EditItemsSpecified = true;
            permissionSet.Permissions[0].DeleteItems = new PermissionActionType();
            permissionSet.Permissions[0].DeleteItems = PermissionActionType.None;
            permissionSet.Permissions[0].DeleteItemsSpecified = true;
            permissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7802");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7802
            // Extended property is not set in "CreateFolder" operation.
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                7802,
                @"[In t:BaseFolderType Complex Type]This element [ExtendedProperty] is not present, server responses NO_ERROR.");

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a subfolder under the folder created in step 1 with User2's credential

            // CreateFolder request.
            CreateFolderType createFolderInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderInSharedMailboxRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                createFolderInSharedMailboxResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                createFolderInSharedMailboxResponse.ResponseMessages.Items.GetLength(0));

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R47001");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R47001
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass,
                47001,
                @"[In t:BasePermissionType Complex Type][CanCreateSubFolders]A value of ""false"" indicates that the client cannot create a sub folder. ");

            #region Edit items that User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;
            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4502");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4502
            this.Site.CaptureRequirementIfIsFalse(
                this.CanCreateItem,
                4502,
                @"[In t:BasePermissionType Complex Type][CanCreateItems]A value of ""false"" indicates that the client cannot create an item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R51001");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R51001
            // Value "None" has been set to EidtItem, if items that user owns and doesn't own can be edited this requirement can be captured.
            bool isVerifiedR51001 = !this.CanEditOwnedItem && !this.CanEditNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR51001,
               "Can edit owned item expected to be \"false\" and actual is {0};\n" +
               "Can edit not owned item expected to be \"false\" and actual is {1};\n ",
               this.CanEditOwnedItem,
               this.CanEditNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR51001,
                51001,
                @"[In t:BasePermissionType Complex Type] The type of element EditItems is ""None"", which indicates that the user does not have permission to perform the action on any items in the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R52001");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R52001
            bool isVerifyR52001 = !this.CanDeleteOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifyR52001,
                "Can delete owned item expected to be \"false\" and actual is {0};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {1};\n ",
                this.CanDeleteOwnedItem,
                this.CanDeleteNotOwnedItem);

            // Value "None" has been set to DeleteItem, if items that user owns and doesn't own can be deleted this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR52001,
                52001,
                @"[In t:BasePermissionType Complex Type] The type of element DeleteItems is ""None"", which indicates that the user does not have permission to perform the action on any items in the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to owner.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC03_FolderPermissionOwnerLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and enable Owner permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Owner;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            permissionSet.Permissions[0].UserId.DisplayName = Common.GetConfigurationPropertyValue("User2Name", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create a sub folder in the folder created in step 1 with User1's credential

            // CreateFolder request.
            CreateFolderType createFolderNotOwnedInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderNotOwnedInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderNotOwnedInSharedMailboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderNotOwnedInSharedMailboxResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdInSharedMailboxNotOwned = ((FolderInfoResponseMessageType)createFolderNotOwnedInSharedMailboxResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Get the new created folder in step 1 with User2's credential

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);
            getFolderRequest.FolderShape.AdditionalProperties = new BasePathToElementType[]
                    {
                        new PathToUnindexedFieldType()
                        {
                            FieldURI = UnindexedFieldURIType.folderPermissionSet
                        }
                    };

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType allFoldersInformation = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0];

            #endregion

            #region Update Folder created in step 2 with User2's credential

            // UpdateFolder request.
            UpdateFolderType updateSubFolderNotOwnedRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", newFolderIdInSharedMailboxNotOwned);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateSubFolderNotOwnedResponse = this.FOLDAdapter.UpdateFolder(updateSubFolderNotOwnedRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateSubFolderNotOwnedResponse, 1, this.Site);

            this.CanEditSubFolder = ResponseClassType.Success == updateSubFolderNotOwnedResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Get the new created Subfolder in step 2 with User2's credential

            // GetFolder request.
            GetFolderType getSubFolderNotOwnedRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdInSharedMailboxNotOwned);

            // Get the new created folder.
            GetFolderResponseType getSubFolderNotOwnedResopnse = this.FOLDAdapter.GetFolder(getSubFolderNotOwnedRequest);

            // Check the response.
            Common.CheckOperationSuccess(getSubFolderNotOwnedResopnse, 1, this.Site);

            this.CanReadSubFolder = ResponseClassType.Success == getSubFolderNotOwnedResopnse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Delete the created folder in step 2 with User2's credential

            // DeleteFolder request.
            DeleteFolderType deleteFolderInSharedMailBoxRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, newFolderIdInSharedMailboxNotOwned);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderInSharedMailBoxResponse = this.FOLDAdapter.DeleteFolder(deleteFolderInSharedMailBoxRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderInSharedMailBoxResponse, 1, this.Site);

            this.CanDeleteSubFolder = ResponseClassType.Success == deleteFolderInSharedMailBoxResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Create a subfolder under the folder created in step 1 with User2's credential

            // CreateFolder request.
            CreateFolderType createFolderInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderInSharedMailboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderInSharedMailboxResponse, 1, this.Site);

            this.CanCreateSubFolder = ResponseClassType.Success == createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Edit items that User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;
            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R146");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R146
            bool isVerifiedR146 = this.CanCreateSubFolder && this.CanCreateItem && this.CanReadOwnedItem && this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && this.CanEditNotOwnedItem && this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR146,
                "Can create subfolder expected to be \"true\" and actual is {0};\n" +
                "Can create item expected to be \"true\" and actual is {1};\n" +
                "Can read owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit owned item expected to be \"true\" and actual is {3};\n" +
                "Can delete owned item expected to be \"true\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"true\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"true\" and actual is {7};\n ",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR146,
                146,
                @"[In t:PermissionLevelType Simple Type]The value Owner means the user can create, read, edit, and delete all items in the folder and create subfolders. The user is both folder owner and folder contact.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R810107");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R810107
            bool isVerifiedR810107 = allFoldersInformation.Folders[0].EffectiveRights.Delete == true && this.CanDeleteOwnedItem && this.CanDeleteNotOwnedItem && this.CanDeleteSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR810107,
                "Delete to be \"true\" and actual is {0};\n" +
                "Can delete owned item expected to be \"true\" and actual is {1};\n" +
                "Can delete not owned item expected to be \"true\" and actual is {2};\n" +
                "Can delete subfolder item expected to be \"true\" and actual is {3};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Delete,
                this.CanDeleteOwnedItem,
                this.CanDeleteNotOwnedItem,
                this.CanDeleteSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR810107,
                810107,
                @"[In t:BaseFolderType Complex Type] Value ""true"" of the element Delete of EffectiveRights indicates a client can delete a folder or item.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8101011");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R8101011
            bool isVerifiedR8101011 = allFoldersInformation.Folders[0].EffectiveRights.Read == true && this.CanReadOwnedItem && this.CanReadNotOwnedItem && this.CanReadSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR8101011,
                "Read to be \"true\" and actual is {0};\n" +
                "Can read owned item expected to be \"true\" and actual is {1};\n" +
                "Can read not owned item expected to be \"true\" and actual is {2};\n" +
                "Can read subfolder item expected to be \"true\" and actual is {3};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Read,
                this.CanReadOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanReadSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR8101011,
                8101011,
                @"[In t:BaseFolderType Complex Type] Value ""true"" of the element Read of EffectiveRights indicates a client can read a folder or item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R810109");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R810109
            bool isVerifiedR810109 = allFoldersInformation.Folders[0].EffectiveRights.Modify == true && this.CanEditOwnedItem && this.CanEditNotOwnedItem && this.CanEditSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR810109,
                "Modify to be \"true\" and actual is {0};\n" +
                "Can edit owned item expected to be \"true\" and actual is {1};\n" +
                "Can edit not owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit subfolder item expected to be \"true\" and actual is {3};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Modify,
                this.CanEditOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanEditSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR810109,
                810109,
                @"[In t:BaseFolderType Complex Type] Value ""true"" of the element Modify of EffectiveRights indicates a client can modify a folder or item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1182");

            // Permission is set in create folder request and permissions are successfully applied, this requirement can be captured.
            this.Site.CaptureRequirement(
                1182,
                @"[In t:PermissionSetType Complex Type][Permissions] specifies a collection of permissions for a folder.");

            UserIdType defaultUserId = ((BasePermissionType)((FolderType)allFoldersInformation.Folders[0]).PermissionSet.Permissions[0]).UserId;
            UserIdType anonymousUserId = ((BasePermissionType)((FolderType)allFoldersInformation.Folders[0]).PermissionSet.Permissions[1]).UserId;
            UserIdType authorizedUserId = ((BasePermissionType)((FolderType)allFoldersInformation.Folders[0]).PermissionSet.Permissions[2]).UserId;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1298");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1298
            // SID is returned and schema is verified in adapter and the specified user has gotten specified permission, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                authorizedUserId.SID,
                "MS-OXWSCDATA",
                1298,
                @"[In t:UserIdType Complex Type] The element ""SID"" with type ""xs:string"" specifies the security descriptor definition language (SSDL) form of the security identifier (SID) for a user.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1299");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1299
            // PrimarySmtpAddress is returned and schema is verified in adapter and the specified user has gotten specified permission, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                authorizedUserId.PrimarySmtpAddress,
                "MS-OXWSCDATA",
                1299,
                @"[In t:UserIdType Complex Type] The element ""PrimarySmtpAddress"" with type ""xs:string"" specifies the primary SMTP address of an account.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1300");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1300
            // Display name is returned and schema is verified in adapter and the specified user has gotten specified permission, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                authorizedUserId.DisplayName,
                "MS-OXWSCDATA",
                1300,
                @"[In t:UserIdType Complex Type] The element ""DisplayName"" with type ""xs:string"" specifies the user name for display.");

            Site.Assert.AreEqual<DistinguishedUserType>(DistinguishedUserType.Default, defaultUserId.DistinguishedUser, "Default user's user id type should be default!");
            Site.Assert.AreEqual<DistinguishedUserType>(DistinguishedUserType.Anonymous, anonymousUserId.DistinguishedUser, "Anonymous user's user id type should be anonymous!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1301");

            // The account is a default user and anonymous user.
            this.Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1301,
                @"[In t:UserIdType Complex Type] [The element ""DistinguishedUser"" with type ""t:DistinguishedUserType"" specifies a value that identifies the Anonymous and Default user accounts for delegate access.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to PublishingEditor.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC04_FolderPermissionPublishingEditorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.PublishingEditor);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R147");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R147
            bool isVerifiedR147 = this.CanCreateSubFolder && this.CanCreateItem && this.CanReadOwnedItem && this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && this.CanEditNotOwnedItem && this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR147,
                "Can create subfolder expected to be \"true\" and actual is {0};\n" +
                "Can create item expected to be \"true\" and actual is {1};\n" +
                "Can read owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit owned item expected to be \"true\" and actual is {3};\n" +
                "Can delete owned item expected to be \"true\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"true\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"true\" and actual is {7};\n ",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR147,
                147,
                @"[In t:PermissionLevelType Simple Type]The value PublishingEditor means the user can create, read, edit, and delete all items in the folder, and create subfolders.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1231");

            // Permission Level is set and authorized user do have given permission, so this requirement can be captured.
            this.Site.CaptureRequirement(
                1231,
                @"[In t:PermissionType Complex Type][PermissionLevel] Specifies the combination of permissions that a user has on a folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to Editor.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC05_FolderPermissionEditorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.Editor);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R148");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R148
            bool isVerifiedR148 = !this.CanCreateSubFolder && this.CanCreateItem && this.CanReadOwnedItem && this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && this.CanEditNotOwnedItem && this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR148,
                "Can create subfolder expected to be \"false\" and actual is {0};\n" +
                "Can create item expected to be \"true\" and actual is {1};\n" +
                "Can read owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit owned item expected to be \"true\" and actual is {3};\n" +
                "Can delete owned item expected to be \"true\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"true\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"true\" and actual is {7};\n ",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR148,
                148,
                @"[In t:PermissionLevelType Simple Type]The value Editor means the user can create, read, edit, and delete all items in the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to PublishingAuthor.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC06_FolderPermissionPublishingAuthorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.PublishingAuthor);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R149");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R149
            bool isVerifiedR149 = this.CanCreateSubFolder && this.CanCreateItem && this.CanReadOwnedItem && this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && !this.CanEditNotOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR149,
                "Can create subfolder expected to be \"true\" and actual is {0};\n" +
                "Can create item expected to be \"true\" and actual is {1};\n" +
                "Can read owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit owned item expected to be \"true\" and actual is {3};\n" +
                "Can delete owned item expected to be \"true\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {7};\n",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR149,
                149,
                @"[In t:PermissionLevelType Simple Type]The value PublishingAuthor means the user can create and read all items in the folder, edit and delete only items that the user creates, and create subfolders.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to Author.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC07_FolderPermissionAuthorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.Author);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R150");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R150
            bool isVerifiedR150 = this.CanCreateItem && this.CanReadOwnedItem && this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && !this.CanEditNotOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR150,
                "Can create item expected to be \"true\" and actual is {0};\n" +
                "Can read owned item expected to be \"true\" and actual is {1};\n" +
                "Can edit owned item expected to be \"true\" and actual is {2};\n" +
                "Can delete owned item expected to be \"true\" and actual is {3};\n" +
                "Can read not owned item expected to be \"true\" and actual is {4};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {5};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {6};\n",
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR150,
                150,
                @"[In t:PermissionLevelType Simple Type]The value Author means the user can create and read all items in the folder, and edit and delete only items that the user creates.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to NonEditingAuthor.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC08_FolderPermissionNonEditingAuthorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.NoneditingAuthor);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R151");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R151
            bool isVerifiedR151 = !this.CanCreateSubFolder && this.CanCreateItem && this.CanReadOwnedItem && !this.CanEditOwnedItem && this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && !this.CanEditNotOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR151,
                "Can create subfolder expected to be \"false\" and actual is {0};\n" +
                "Can create item expected to be \"true\" and actual is {1};\n" +
                "Can read owned item expected to be \"true\" and actual is {2};\n" +
                "Can edit owned item expected to be \"false\" and actual is {3};\n" +
                "Can delete owned item expected to be \"true\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {7};\n",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR151,
                151,
                @"[In t:PermissionLevelType Simple Type]The value NoneditingAuthor means the user can create and read all items in the folder, and delete only items that the user creates.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to Reviewer.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC09_FolderPermissionReviewerLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.Reviewer);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R152");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R152
            bool isVerifiedR152 = !this.CanCreateSubFolder && !this.CanCreateItem && !this.CanReadOwnedItem && !this.CanEditOwnedItem && !this.CanDeleteOwnedItem && this.CanReadNotOwnedItem && !this.CanEditNotOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR152,
                "Can create subfolder expected to be \"false\" and actual is {0};\n" +
                "Can create item expected to be \"false\" and actual is {1};\n" +
                "Can read owned item expected to be \"false\" and actual is {2};\n" +
                "Can edit owned item expected to be \"false\" and actual is {3};\n" +
                "Can delete owned item expected to be \"false\" and actual is {4};\n" +
                "Can read not owned item expected to be \"true\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {7};\n",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR152,
                152,
                @"[In t:PermissionLevelType Simple Type]The value Reviewer means the user can read all items in the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to Contributor.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC10_FolderPermissionContributorLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and enable Contributor permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Contributor;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create a subfolder

            // CreateFolder request.
            CreateFolderType createFolderNotOwnedInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderNotOwnedInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderNotOwnedInSharedMailboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderNotOwnedInSharedMailboxResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderIdNotOwnedInSharedMailbox = ((FolderInfoResponseMessageType)createFolderNotOwnedInSharedMailboxResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Get the new created folder in step 1 with User2's credential

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType allFoldersInformation = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0];

            #endregion

            #region Update Folder created in step 2 with User2's credential

            // UpdateFolder request.
            UpdateFolderType updateSubFolderNotOwnedRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", newFolderIdNotOwnedInSharedMailbox);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateSubFolderNotOwnedResponse = this.FOLDAdapter.UpdateFolder(updateSubFolderNotOwnedRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                updateSubFolderNotOwnedResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                updateSubFolderNotOwnedResponse.ResponseMessages.Items.GetLength(0));

            this.CanEditSubFolder = ResponseClassType.Success == updateSubFolderNotOwnedResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Get the new created Subfolder in step 2 with User2's credential

            // GetFolder request.
            GetFolderType getSubFolderNotOwnedRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderIdNotOwnedInSharedMailbox);

            // Get the new created folder.
            GetFolderResponseType getSubFolderNotOwnedResopnse = this.FOLDAdapter.GetFolder(getSubFolderNotOwnedRequest);

            // Check the response.
            Common.CheckOperationSuccess(getSubFolderNotOwnedResopnse, 1, this.Site);

            this.CanReadSubFolder = ResponseClassType.Success == getSubFolderNotOwnedResopnse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Delete the subfolder created in step 2 with User2's credential

            // DeleteFolder request.
            DeleteFolderType deleteFolderInSharedMailBoxRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, newFolderIdNotOwnedInSharedMailbox);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderInSharedMailBoxResponse = this.FOLDAdapter.DeleteFolder(deleteFolderInSharedMailBoxRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                deleteFolderInSharedMailBoxResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                deleteFolderInSharedMailBoxResponse.ResponseMessages.Items.GetLength(0));

            this.CanDeleteSubFolder = ResponseClassType.Success == deleteFolderInSharedMailBoxResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Create a subfolder under the folder created in step 1 with User2's credential

            // CreateFolder request.
            CreateFolderType createFolderInSharedMailboxAfteSwitchRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderOwnedInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderInSharedMailboxAfteSwitchRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                createFolderOwnedInSharedMailboxResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                createFolderOwnedInSharedMailboxResponse.ResponseMessages.Items.GetLength(0));

            this.CanCreateSubFolder = ResponseClassType.Success == createFolderOwnedInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Edit items that User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;

            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R810108");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R810108
            bool isVerifiedR810108 = allFoldersInformation.Folders[0].EffectiveRights.Delete == false && !this.CanDeleteOwnedItem && !this.CanDeleteNotOwnedItem && !this.CanDeleteSubFolder;

            Site.Assert.IsTrue(
                isVerifiedR810108,
                "Delete expected to be \"false\" and actual is {0};\n" +
                "Can delete owned item expected to be \"false\" and actual is {1};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {2};\n" +
                "Can delete subfolder expected to be \"false\" and actual is {3};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Delete,
                this.CanDeleteOwnedItem,
                this.CanDeleteNotOwnedItem,
                this.CanDeleteSubFolder);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR810108,
                810108,
                @"[In t:BaseFolderType Complex Type] Value ""false"" of the element Delete of EffectiveRights indicates a client cannot delete a folder or item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8101010");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R8101010
            bool isVerifiedR8101010 = allFoldersInformation.Folders[0].EffectiveRights.Modify == false && !this.CanEditSubFolder && !this.CanEditOwnedItem && !this.CanEditNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR8101010,
                "Modify expected to be \"false\" and actual is {0};\n" +
                "Can edit subfolder expected to be \"false\" and actual is {1};\n" +
                "Can edit owned item expected to be \"false\" and actual is {2};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {3};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Modify,
                this.CanEditSubFolder,
                this.CanEditOwnedItem,
                this.CanEditNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR8101010,
                8101010,
                @"[In t:BaseFolderType Complex Type] Value ""false"" of the element Modify of EffectiveRights indicates a client can not modify a folder or item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8101012");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R8101010
            bool isVerifiedR8101012 = allFoldersInformation.Folders[0].EffectiveRights.Read == false && !this.CanReadOwnedItem && !this.CanReadNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR8101012,
                "Modify expected to be \"false\" and actual is {0};\n" +
                "Can read owned item expected to be \"false\" and actual is {1};\n" +
                "Can read not owned item expected to be \"false\" and actual is {2};\n ",
                allFoldersInformation.Folders[0].EffectiveRights.Read,
                this.CanReadOwnedItem,
                this.CanReadNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR8101012,
                8101012,
                @"[In t:BaseFolderType Complex Type] Value ""false"" of the element Read of EffectiveRights indicates a client can not read a folder or item.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission level set to None.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC11_FolderPermissionNoneLevel()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            this.ValidateFolderPermissionLevel(PermissionLevelType.None);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R145");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R145
            bool isVerifiedR145 = !this.CanCreateSubFolder && !this.CanCreateItem && !this.CanReadOwnedItem && !this.CanEditOwnedItem && !this.CanDeleteOwnedItem && !this.CanReadNotOwnedItem && !this.CanEditNotOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR145,
                "Can create subfolder expected to be \"false\" and actual is {0};\n" +
                "Can create item expected to be \"false\" and actual is {1};\n" +
                "Can read owned item expected to be \"false\" and actual is {2};\n" +
                "Can edit owned item expected to be \"false\" and actual is {3};\n" +
                "Can delete owned item expected to be \"false\" and actual is {4};\n" +
                "Can read not owned item expected to be \"false\" and actual is {5};\n" +
                "Can edit not owned item expected to be \"false\" and actual is {6};\n" +
                "Can delete not owned item expected to be \"false\" and actual is {7};\n ",
                this.CanCreateSubFolder,
                this.CanCreateItem,
                this.CanReadOwnedItem,
                this.CanEditOwnedItem,
                this.CanDeleteOwnedItem,
                this.CanReadNotOwnedItem,
                this.CanEditNotOwnedItem,
                this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR145,
                145,
                @"[In t:PermissionLevelType Simple Type]The value None means the user has no permissions on the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to setting all permissions or none.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC12_FolderPermissionWithOrWithoutAllSet()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a new folder without any permissions in the inbox folder

            // Configure permission set.
            PermissionSetType permissionSetFirst = new PermissionSetType();
            permissionSetFirst.Permissions = new PermissionType[1];
            permissionSetFirst.Permissions[0] = new PermissionType();
            permissionSetFirst.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSetFirst.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSetFirst.Permissions[0].UserId = new UserIdType();
            permissionSetFirst.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequestFirst = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSetFirst });

            // Create a new folder.
            CreateFolderResponseType createFolderResponseFirst = this.FOLDAdapter.CreateFolder(createFolderRequestFirst);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponseFirst, 1, this.Site);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R122302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R122302
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                122302,
                @"[In t:PermissionType Complex Type]This element [ReadItems] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R460102");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R460102
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                460102,
                @"[In t:BasePermissionType Complex Type]This element [CanCreateItems] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R470302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R470302
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                470302,
                @"[In t:BasePermissionType Complex Type]This element [CanCreateSubFolders] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R480302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R480302
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                480302,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderOwner] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R490302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R490302
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                490302,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderVisible] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R500302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R500302
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                500302,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderContact] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R510202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R510202
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                510202,
                @"[In t:BasePermissionType Complex Type]This element [EditItems] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R520202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R520202
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                520202,
                @"[In t:BasePermissionType Complex Type]This element [DeleteItems] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R120302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R120302
            // Since the UnknownEntries is not set and server responds NoError, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseCode,
                120302,
                @"[In t:PermissionSetType Complex Type]This element [UnknownEntries] is not present, server responses NO_ERROR.");

            // Save the new created folder's folder id.
            FolderIdType firstFolderId = ((FolderInfoResponseMessageType)createFolderResponseFirst.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(firstFolderId);

            #endregion

            #region Create a new folder with all permissions in the inbox folder

            // Configure permission set.
            PermissionSetType permissionSetSecond = new PermissionSetType();
            permissionSetSecond.Permissions = new PermissionType[1];
            permissionSetSecond.Permissions[0] = new PermissionType();
            permissionSetSecond.Permissions[0].ReadItems = new PermissionReadAccessType();
            permissionSetSecond.Permissions[0].ReadItems = PermissionReadAccessType.FullDetails;
            permissionSetSecond.Permissions[0].CanCreateItemsSpecified = true;
            permissionSetSecond.Permissions[0].CanCreateItems = false;
            permissionSetSecond.Permissions[0].CanCreateSubFoldersSpecified = true;
            permissionSetSecond.Permissions[0].CanCreateSubFolders = false;
            permissionSetSecond.Permissions[0].IsFolderOwnerSpecified = true;
            permissionSetSecond.Permissions[0].IsFolderOwner = false;
            permissionSetSecond.Permissions[0].IsFolderVisibleSpecified = true;
            permissionSetSecond.Permissions[0].IsFolderVisible = false;
            permissionSetSecond.Permissions[0].IsFolderContactSpecified = true;
            permissionSetSecond.Permissions[0].IsFolderContact = false;
            permissionSetSecond.Permissions[0].EditItemsSpecified = true;
            permissionSetSecond.Permissions[0].EditItems = PermissionActionType.All;
            permissionSetSecond.Permissions[0].DeleteItemsSpecified = true;
            permissionSetSecond.Permissions[0].DeleteItems = PermissionActionType.All;
            permissionSetSecond.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSetSecond.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSetSecond.Permissions[0].UserId = new UserIdType();
            permissionSetSecond.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequestSecond = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSetSecond });

            // Create a new folder.
            CreateFolderResponseType createFolderResponseSecond = this.FOLDAdapter.CreateFolder(createFolderRequestSecond);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponseSecond, 1, this.Site);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R980301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R980301
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponseSecond.ResponseMessages.Items[0].ResponseCode,
                980301,
                @"[In t:FolderType Complex Type]This element [PermissionSet] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R123202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R123202
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                 ResponseCodeType.NoError,
                createFolderResponseSecond.ResponseMessages.Items[0].ResponseCode,
                123202,
                @"[In t:PermissionType Complex Type]This element [PermissionLevel] is not present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R123201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R123201
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseSecond.ResponseMessages.Items[0].ResponseClass,
                123201,
                @"[In t:PermissionType Complex Type]This element [PermissionLevel] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R122301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R122301
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseSecond.ResponseMessages.Items[0].ResponseClass,
                122301,
                @"[In t:PermissionType Complex Type]This element [ReadItems] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R460101");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R460101
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                460101,
                @"[In t:BasePermissionType Complex Type]This element [CanCreateItems] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R470301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R470301
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                470301,
                @"[In t:BasePermissionType Complex Type]This element [CanCreateSubFolders] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R480301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R480301
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                480301,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderOwner] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R490301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R490301
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                490301,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderVisible] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R500301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R500301
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                500301,
                @"[In t:BasePermissionType Complex Type]This element [IsFolderContact] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R510201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R510201
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                510201,
                @"[In t:BasePermissionType Complex Type]This element [EditItems] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R520201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R520201
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseFirst.ResponseMessages.Items[0].ResponseClass,
                520201,
                @"[In t:BasePermissionType Complex Type]This element [DeleteItems] is present, server responses NO_ERROR.");

            // Save the new created folder's folder id.
            FolderIdType secondFolderId = ((FolderInfoResponseMessageType)createFolderResponseSecond.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(secondFolderId);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, secondFolderId);
            getFolderRequest.FolderShape.AdditionalProperties = new BasePathToElementType[]
                    {
                        new PathToUnindexedFieldType()
                        {
                            FieldURI = UnindexedFieldURIType.folderPermissionSet
                        }
                    };

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R45");

            // Since got folder successfully and schema is validated, so this requirement can be directly captured.
            this.Site.CaptureRequirement(
                45,
                @"[In t:BasePermissionType Complex Type]The type of element CanCreateItems is xs:boolean [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R47");

            // Since got folder successfully and schema is validated, so this requirement can be directly captured.
            this.Site.CaptureRequirement(
                47,
                @"[In t:BasePermissionType Complex Type]The type of element CanCreateSubFolders is xs:boolean.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R48");

            // Since got folder successfully and schema is validated, so this requirement can be directly captured.
            this.Site.CaptureRequirement(
                48,
                @"[In t:BasePermissionType Complex Type]The type of element IsFolderOwner is xs:boolean.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R49");

            // Since got folder successfully and schema is validated, so this requirement can be directly captured.
            this.Site.CaptureRequirement(
                49,
                @"[In t:BasePermissionType Complex Type]The type of element IsFolderVisible is xs:boolean.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R50");

            // Since got folder successfully and schema is validated, so this requirement can be directly captured.
            this.Site.CaptureRequirement(
                50,
                @"[In t:BasePermissionType Complex Type]The type of element IsFolderContact is xs:boolean.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission read item set to FullDetail.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC13_FolderPermissionReadItemFullDetail()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and enable ReadItem with full detail permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].ReadItems = new PermissionReadAccessType();
            permissionSet.Permissions[0].ReadItems = PermissionReadAccessType.FullDetails;
            permissionSet.Permissions[0].ReadItemsSpecified = true;
            permissionSet.Permissions[0].CanCreateItems = true;
            permissionSet.Permissions[0].CanCreateItemsSpecified = true;
            permissionSet.Permissions[0].IsFolderOwner = true;
            permissionSet.Permissions[0].IsFolderOwnerSpecified = true;
            permissionSet.Permissions[0].EditItems = new PermissionActionType();
            permissionSet.Permissions[0].EditItems = PermissionActionType.Owned;
            permissionSet.Permissions[0].EditItemsSpecified = true;
            permissionSet.Permissions[0].DeleteItems = new PermissionActionType();
            permissionSet.Permissions[0].DeleteItems = PermissionActionType.Owned;
            permissionSet.Permissions[0].DeleteItemsSpecified = true;
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Edit items that User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;

            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R158");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R158
            // If items that user owns and doesn't can both be read this requirement can be captured.
            bool isVerifiedR158 = this.CanReadOwnedItem && this.CanReadNotOwnedItem;

            Site.Assert.IsTrue(
                isVerifiedR158,
                "Can read owned item expected to be \"true\" and actual is {0};\n" +
                "Can read not owned item expected to be \"true\" and actual is {1};\n ",
                this.CanReadOwnedItem,
                this.CanReadNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR158,
                158,
                @"[In t:PermissionReadAccessType Simple Type]The value FullDetails means the user has permission to read all items in the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R51002");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R51002
            // If only items that user created can be edited, this requirement can be captured.
            bool isVerifiedR51002 = this.CanEditOwnedItem && !this.CanEditNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR51002,
               "Can edit owned item expected to be \"true\" and actual is {0};\n" +
               "Can edit not owned item expected to be \"false\" and actual is {1};\n ",
               this.CanEditOwnedItem,
               this.CanEditNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR51002,
                51002,
                @"[In t:BasePermissionType Complex Type]The type of element EditItems is ""Owned"", which indicates that the user has permission to perform the action on the items in the folder that the user owns.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R52002");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R52002
            // If only items that user created can be deleted, this requirement can be captured.
            bool isVerifiedR52002 = this.CanDeleteOwnedItem && !this.CanDeleteNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR52002,
               "Can delete owned item expected to be \"true\" and actual is {0};\n" +
               "Can delete not owned item expected to be \"false\" and actual is {1};\n ",
               this.CanDeleteOwnedItem,
               this.CanDeleteNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR52002,
                52002,
                @"[In t:BasePermissionType Complex Type]The type of element DeleteItems is ""Owned"", which indicates that the user has permission to perform the action on the items in the folder that the user owns.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R122");

            // Read items is set in permission and schema is validated in adapter, so this requirement can be captured.
            this.Site.CaptureRequirement(
                122,
                @"[In t:PermissionType Complex Type]The type of element ReadItems is t:PermissionReadAccessType (section 2.2.5.4).");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission read item set to None.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC14_FolderPermissionReadItemNone()
        {
            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a folder in the User1's inbox folder, and enable ReadItem with none permission for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].ReadItems = new PermissionReadAccessType();
            permissionSet.Permissions[0].ReadItems = PermissionReadAccessType.None;
            permissionSet.Permissions[0].ReadItemsSpecified = true;
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Read items that User2 doesn't own with User2's credential

            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", Site) + "@" + Common.GetConfigurationPropertyValue("Domain", Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;

            if (this.CanCreateItem)
            {
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
            }
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R157");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R157
            // If all items can't be read, this requirement can be captured.
            bool isVerifiedR157 = !this.CanReadOwnedItem && !this.CanReadNotOwnedItem;

            Site.Assert.IsTrue(
               isVerifiedR157,
               "Can read owned item expected to be \"false\" and actual is {0};\n" +
               "Can read not owned item expected to be \"false\" and actual is {1};\n ",
               this.CanReadOwnedItem,
               this.CanReadNotOwnedItem);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR157,
                157,
                @"[In t:PermissionReadAccessType Simple Type]The value None means the user does not have permission to read items in the folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission anyone field other than UserId field is set, and the PermissionLevel field is not set to "Custom".
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC15_FolderPermissionLevelNotCustom()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(54911, this.Site), @"Exchange 2007 and Exchange 2010 will return an ErrorInvalidPermissionSettings ([MS-OXWSCDATA] section 2.2.5.24) response code if any field of BasePermissionType other than UserId field is set, and the PermissionLevel field is not set to ""Custom"".");

            #region Switch to User1
            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));
            #endregion

            #region Create a folder in the User1's inbox folder, and enable Contributor permission for User2
            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Set the field CanCreateSubFolders to 'true', and the PermissionLevel field is not set to 'Custom'
            permissionSet.Permissions[0].CanCreateSubFolders = true;
            permissionSet.Permissions[0].CanCreateSubFoldersSpecified = true;
            permissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Contributor;

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R54911");

            // CanCreateSubFolders is set and the PermissionLevel field is not set to 'Custom', the expected ResponseCode is ErrorInvalidPermissionSettings
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPermissionSettings,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                54911,
                @"[In Appendix C: Product Behavior] Implementation does return an ErrorInvalidPermissionSettings response code if any field of BasePermissionType other than UserId field is set, and the PermissionLevel field is not set to ""Custom"". (<7> Section 2.2.4.15:  Exchange 2007 and Exchange 2010 will return an ErrorInvalidPermissionSettings ([MS-OXWSCDATA] section 2.2.5.24) response code if any field of BasePermissionType other than UserId field is set, and the PermissionLevel field is not set to ""Custom"".)");
        }

        /// <summary>
        /// This test case verifies requirements related to folder permission, the PermissionLevel field is set to "Contributor".
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S07_TC16_FolderPermissionLevelContributor()
        {
            #region Switch to User1
            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));
            #endregion

            #region Create a folder in the User1's inbox folder, and enable Contributor permission for User2
            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            permissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            permissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Contributor;

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);
            #endregion

            #region Switch to User2
            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));
            #endregion

            #region Create an item in the folder created in step 1 with User2's credential
            string itemName = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemInFolder = this.CreateItem(Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName);
            Site.Assert.IsNotNull(itemInFolder, "Item should be created successfully!");
            #endregion

            #region Read the new item
            bool canReadItem = this.GetItem(itemInFolder);
            #endregion

            if (Common.IsRequirementEnabled(1531, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1531");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1531, if User2 can read the item, this capture can be verified.
                this.Site.CaptureRequirementIfIsTrue(
                    canReadItem,
                    1531,
                    @"[In Appendix C: Product Behavior] The implementation does support Contributor in PermissionLevelType specifies that the user can create items in the folder and read those items. (<9> Section 2.2.5.3:  In Microsoft Exchange Server 2013 Service Pack 1 (SP1) and Exchange 2016 the user can create items in the folder and read those items.)");
            }

            if (Common.IsRequirementEnabled(1532, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1532");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1531, if User2 can not read the item, this capture can be verified.
                this.Site.CaptureRequirementIfIsFalse(
                    canReadItem,
                    1532,
                    @"[In Appendix C: Product Behavior] The implementation does support Contributor in PermissionLevelType specifies that the user can create items in the folder but cannot read any items in the folder. (Exchange 2007 and Exchange 2010 follow this behavior.)");
            }
        }
        #endregion
    }
}