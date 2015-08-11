namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify all operations for all optional elements.
    /// </summary>
    [TestClass]
    public class S08_OptionalElements : TestSuiteBase
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
        /// This test case verifies requirements related to all operations with all optional elements.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S08_TC01_AllOperationsWithAllOptionalElements()
        {
            #region Configure SOAP header

            this.ConfigureSOAPHeader();

            #endregion

            #region Create new folders in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(
                DistinguishedFolderIdNameType.inbox.ToString(),
                new string[] { "Custom Folder1", "Custom Folder2", "Custom Folder3", "Custom Folder4" },
                new string[] { "IPF.MyCustomFolderClass", "IPF.Appointment", "IPF.Contact", "IPF.Task" },
                null);

            // Set ExtendedProperty defined in BaseFolderType.
            PathToExtendedFieldType publishInAddressBook = new PathToExtendedFieldType();

            // A hexadecimal tag of the extended property.
            publishInAddressBook.PropertyTag = "0x671E";
            publishInAddressBook.PropertyType = MapiPropertyTypeType.Boolean;
            ExtendedPropertyType pubAddressbook = new ExtendedPropertyType();
            pubAddressbook.ExtendedFieldURI = publishInAddressBook;
            pubAddressbook.Item = "1";
            ExtendedPropertyType[] extendedProperties = new ExtendedPropertyType[1];
            extendedProperties[0] = pubAddressbook;

            createFolderRequest.Folders[0].ExtendedProperty = extendedProperties;
            createFolderRequest.Folders[1].ExtendedProperty = extendedProperties;
            createFolderRequest.Folders[2].ExtendedProperty = extendedProperties;
            createFolderRequest.Folders[3].ExtendedProperty = extendedProperties;

            // Define a permissionSet with all optional elements
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

            // Set PermissionSet for FolderType folder.
            ((FolderType)createFolderRequest.Folders[0]).PermissionSet = permissionSet;

            // Set PermissionSet for ContactsType folder.
            ((ContactsFolderType)createFolderRequest.Folders[2]).PermissionSet = permissionSet;

            // Set PermissionSet for TasksFolderType folder.
            ((TasksFolderType)createFolderRequest.Folders[3]).PermissionSet = permissionSet;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 4, this.Site);

            // Folder ids.
            FolderIdType[] folderIds = new FolderIdType[createFolderResponse.ResponseMessages.Items.Length];

            for (int index = 0; index < createFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, createFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be created successfully!");

                // Save folder ids.
                folderIds[index] = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[index]).Folders[0].FolderId;

                // Save the new created folder's folder id.
                this.NewCreatedFolderIds.Add(folderIds[index]);
            }

            #endregion

            #region Create a managedfolder

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site));

            // Add an email address into request.
            EmailAddressType mailBox = new EmailAddressType()
            {
                EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            createManagedFolderRequest.Mailbox = mailBox;

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createManagedFolderResponse, 1, this.Site);

            // Save the new created managed folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created folders

            // GetFolder request.
            GetFolderType getCreatedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderIds);

            // Get the new created folder.
            GetFolderResponseType getCreatedFolderResponse = this.FOLDAdapter.GetFolder(getCreatedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getCreatedFolderResponse, 4, this.Site);

            for (int index = 0; index < getCreatedFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, getCreatedFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder information should be returned!");
            }

            #endregion

            #region Update the new created folders

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest(
                new string[] { "Folder", "CalendarFolder", "ContactsFolder", "TasksFolder" },
                new string[] { "SetFolderField", "SetFolderField", "SetFolderField", "SetFolderField" },
                folderIds);

            // Update the folders' properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 4, this.Site);

            for (int index = 0; index < updateFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, updateFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }

            #endregion

            #region Copy the updated folders to "drafts" folder

            // Copy the folders into "drafts" folder
            CopyFolderType copyFolderRequest = this.GetCopyFolderRequest(DistinguishedFolderIdNameType.drafts.ToString(), folderIds);

            // Copy the folders.
            CopyFolderResponseType copyFolderResponse = this.FOLDAdapter.CopyFolder(copyFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyFolderResponse, 4, this.Site);

            // Copied Folders' id.
            FolderIdType[] copiedFolderIds = new FolderIdType[copyFolderResponse.ResponseMessages.Items.Length];

            for (int index = 0; index < copyFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, copyFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");

                // Variable to save the folders.
                copiedFolderIds[index] = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[index]).Folders[0].FolderId;

                // Save the copied folders' folder id.
                this.NewCreatedFolderIds.Add(copiedFolderIds[index]);
            }

            #endregion

            #region Move the updated folders to "deleteditems" folder

            // MoveFolder request.
            MoveFolderType moveFolderRequest = new MoveFolderType();

            // Set the request's folderId field.
            moveFolderRequest.FolderIds = folderIds;

            // Set the request's destFolderId field.
            DistinguishedFolderIdType toFolderId = new DistinguishedFolderIdType();
            toFolderId.Id = DistinguishedFolderIdNameType.deleteditems;
            moveFolderRequest.ToFolderId = new TargetFolderIdType();
            moveFolderRequest.ToFolderId.Item = toFolderId;

            // Move the specified folders.
            MoveFolderResponseType moveFolderResponse = this.FOLDAdapter.MoveFolder(moveFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveFolderResponse, 4, this.Site);

            for (int index = 0; index < moveFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, moveFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }
            #endregion

            #region Delete all folders

            // All folder ids.
            FolderIdType[] allFolderIds = new FolderIdType[folderIds.Length + copiedFolderIds.Length];

            for (int index = 0; index < allFolderIds.Length / 2; index++)
            {
                allFolderIds[index] = folderIds[index];
                allFolderIds[index + folderIds.Length] = copiedFolderIds[index];
            }

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, allFolderIds);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 8, this.Site);

            for (int index = 0; index < deleteFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, deleteFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to all operations without all optional elements.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S08_TC02_AllOperationsWithoutAllOptionalElements()
        {
            #region Create new folders in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(
                DistinguishedFolderIdNameType.inbox.ToString(),
                new string[] { "Custom Folder1", "Custom Folder2", "Custom Folder3", "Custom Folder4" },
                new string[] { "IPF.MyCustomFolderClass", "IPF.Appointment", "IPF.Contact", "IPF.Task" },
                null);

            // Remove FolderClass for FolderType folder.
            createFolderRequest.Folders[0].FolderClass = null;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 4, this.Site);

            // Folder ids.
            FolderIdType[] folderIds = new FolderIdType[createFolderResponse.ResponseMessages.Items.Length];

            for (int index = 0; index < createFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, createFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be created successfully!");

                // Save folder ids.
                folderIds[index] = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[index]).Folders[0].FolderId;

                // Save the new created folder's folder id.
                this.NewCreatedFolderIds.Add(folderIds[index]);
            }

            #endregion

            #region Create a managedfolder

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site));

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createManagedFolderResponse, 1, this.Site);

            // Save the new created managed folder's folder id.
            FolderIdType newManagedFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newManagedFolderId);

            #endregion

            #region Get the new created folders

            // GetFolder request.
            GetFolderType getCreatedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderIds);

            // Get the new created folder.
            GetFolderResponseType getCreatedFolderResponse = this.FOLDAdapter.GetFolder(getCreatedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getCreatedFolderResponse, 4, this.Site);

            for (int index = 0; index < getCreatedFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, getCreatedFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder information should be returned!");
            }

            #endregion

            #region Update the new created folders

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest(
                new string[] { "Folder", "CalendarFolder", "ContactsFolder", "TasksFolder" },
                new string[] { "SetFolderField", "SetFolderField", "SetFolderField", "SetFolderField" },
                folderIds);

            // Update the folders' properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 4, this.Site);

            for (int index = 0; index < updateFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, updateFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }

            #endregion

            #region Copy the updated folders to "drafts" folder

            // Copy the folders into "drafts" folder
            CopyFolderType copyFolderRequest = this.GetCopyFolderRequest(DistinguishedFolderIdNameType.drafts.ToString(), folderIds);

            // Copy the folders.
            CopyFolderResponseType copyFolderResponse = this.FOLDAdapter.CopyFolder(copyFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyFolderResponse, 4, this.Site);

            // Copied folders' id.
            FolderIdType[] copiedFolderIds = new FolderIdType[copyFolderResponse.ResponseMessages.Items.Length];

            for (int index = 0; index < copyFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, copyFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");

                // Variable to save the folders.
                copiedFolderIds[index] = ((FolderInfoResponseMessageType)copyFolderResponse.ResponseMessages.Items[index]).Folders[0].FolderId;

                // Save the copied folders' folder id.
                this.NewCreatedFolderIds.Add(copiedFolderIds[index]);
            }

            #endregion

            #region Move the updated folders to "deleteditems" folder

            // MoveFolder request.
            MoveFolderType moveFolderRequest = new MoveFolderType();

            // Set the request's folderId field.
            moveFolderRequest.FolderIds = folderIds;

            // Set the request's destFolderId field.
            DistinguishedFolderIdType toFolderId = new DistinguishedFolderIdType();
            toFolderId.Id = DistinguishedFolderIdNameType.deleteditems;
            moveFolderRequest.ToFolderId = new TargetFolderIdType();
            moveFolderRequest.ToFolderId.Item = toFolderId;

            // Move the specified folders.
            MoveFolderResponseType moveFolderResponse = this.FOLDAdapter.MoveFolder(moveFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveFolderResponse, 4, this.Site);

            for (int index = 0; index < moveFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, moveFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }

            #endregion

            #region Delete all folders

            // All folder ids.
            FolderIdType[] allFolderIds = new FolderIdType[folderIds.Length + copiedFolderIds.Length];

            for (int index = 0; index < allFolderIds.Length / 2; index++)
            {
                allFolderIds[index] = folderIds[index];
                allFolderIds[index + folderIds.Length] = copiedFolderIds[index];
            }

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, allFolderIds);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 8, this.Site);

            for (int index = 0; index < deleteFolderResponse.ResponseMessages.Items.Length; index++)
            {
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, deleteFolderResponse.ResponseMessages.Items[index].ResponseClass, "Folder should be updated successfully!");
            }

            #endregion
        }
        #endregion
    }
}