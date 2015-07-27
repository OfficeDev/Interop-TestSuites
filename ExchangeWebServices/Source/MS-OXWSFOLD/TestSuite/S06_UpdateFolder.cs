//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify UpdateFolder operation.
    /// </summary>
    [TestClass]
    public class S06_UpdateFolder : TestSuiteBase
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
        /// This test case verifies requirements related to UpdateFolder operation via creating a folder and updating it with SetFolderFieldType property.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC01_UpdateFolder()
        {
            #region Create a new folder in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            FolderIdType folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(folderId);

            #endregion

            #region Update Folder Operation.

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", folderId);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 1, this.Site);

            string updateNameInRequest = ((SetFolderFieldType)updateFolderRequest.FolderChanges[0].Updates[0]).Item1.DisplayName;
            #endregion

            #region Get the updated folder.

            // GetFolder request.
            GetFolderType getUpdatedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderId);

            // Get the updated folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getUpdatedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0];
            FolderType gotFolderInfo = (FolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R46444");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R46444
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                46444,
                @"[In UpdateFolder Operation]A successful UpdateFolder operation request returns an UpdateFolderResponse element with the ResponseClass attribute of the UpdateFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4644");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4644
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                updateFolderResponse.ResponseMessages.Items[0].ResponseCode,
                4644,
                @"[In UpdateFolder Operation]A successful UpdateFolder operation request returns an UpdateFolderResponse element with the ResponseCode element of the UpdateFolderResponse element set to ""NoError"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R8902");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R8902
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                8902,
                @"[In t:FolderChangeType Complex Type]FolderId specifies the folder identifier and change key.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R582");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R582
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                582,
                @"[In m:UpdateFolderType Complex Type]The UpdateFolderType complex type specifies a request message to update folders in a mailbox. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R534");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R534
            // Set a property on a FolderType folder successfully, indicates that Folder represents a regular folder.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                534,
                @"[In t:SetFolderFieldType Complex Type]Folder represents a regular folder in the server database.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R9301");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R9301
            this.Site.CaptureRequirementIfAreEqual<string>(
                updateNameInRequest,
                gotFolderInfo.DisplayName,
                9301,
                @"[In t:FolderChangeType Complex Type][Updates] Specifies a collection of changes to a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R546");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R546
            this.Site.CaptureRequirementIfAreEqual<string>(
                updateNameInRequest,
                gotFolderInfo.DisplayName,
                546,
                @"[In t:FolderChangeType Complex Type]The FolderChangeType complex type specifies a collection of changes to be performed on a single folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5051");

            // All folders updated successfully!
            this.Site.CaptureRequirement(
                5051,
                @"[In m:UpdateFolderType Complex Type]FolderChanges represents an array of folders to be updated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R531");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R531
            this.Site.CaptureRequirementIfAreEqual<string>(
                updateNameInRequest,
                gotFolderInfo.DisplayName,
                531,
                @"[In t:NonEmptyArrayOfFolderChangesType Complex Type]FolderChange represents a collection of changes to be performed on a single folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5251");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5251
            this.Site.CaptureRequirementIfAreEqual<string>(
                updateNameInRequest,
                gotFolderInfo.DisplayName,
                5251,
                @"[In t:NonEmptyArrayOfFolderChangeDescriptionsType Complex Type]SetFolderField represents an UpdateFolder operation to set a property on an existing folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to updating a folder with AppendToFolderFieldType set.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC02_UpdateFolderWithAppendToFolderFieldType()
        {
            #region Create a new folder in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            FolderIdType folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(folderId);

            #endregion

            #region Update Folder Operation with AppendToFolderFieldType Complex Type set.

            // Specified folder to be updated.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "AppendToFolderField", folderId);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                updateFolderResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                updateFolderResponse.ResponseMessages.Items.GetLength(0));

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R507");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R507.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                507,
                @"[In t:AppendToFolderFieldType Complex Type]Any request that uses this complex type will always return an error response.");

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to updating multiple folders in UpdateFolder operation.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC03_UpdateMultipleFolders()
        {
            #region Create a new folder in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(
                DistinguishedFolderIdNameType.inbox.ToString(),
                new string[] { "Custom Folder1", "Custom Folder2", "Custom Folder3", "Custom Folder4" },
                new string[] { "IPF.Appointment", "IPF.Contact", "IPF.Task", "IPF.Search" },
                null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 4, this.Site);

            FolderIdType folderId1 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            FolderIdType folderId2 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[1]).Folders[0].FolderId;
            FolderIdType folderId3 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[2]).Folders[0].FolderId;
            FolderIdType folderId4 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[3]).Folders[0].FolderId;

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(folderId1);
            this.NewCreatedFolderIds.Add(folderId2);
            this.NewCreatedFolderIds.Add(folderId3);
            this.NewCreatedFolderIds.Add(folderId4);

            #endregion

            #region Update Folder Operation.

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest(
                new string[] { "CalendarFolder", "ContactsFolder", "TasksFolder", "SearchFolder" },
                new string[] { "SetFolderField", "SetFolderField", "SetFolderField", "SetFolderField" },
                new FolderIdType[] { folderId1, folderId2, folderId3, folderId4 });

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 4, this.Site);

            #endregion

            #region Get the updated folder.

            // GetFolder request.
            GetFolderType getUpdatedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderId1, folderId2, folderId3, folderId4);

            // Get the updated folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getUpdatedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 4, this.Site);

            FolderInfoResponseMessageType allPropertyOfSearchFolder = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[3];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R33");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R33
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allPropertyOfSearchFolder.Folders[0],
                typeof(SearchFolderType),
                33,
                @"[In t:ArrayOfFoldersType Complex Type]The type of element SearchFolder is t:SearchFolderType ([MS-OXWSSRCH] section 2.2.4.26).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3302
            this.Site.CaptureRequirementIfIsInstanceOfType(
                allPropertyOfSearchFolder.Folders[0],
                typeof(SearchFolderType),
                3302,
                @"[In t:ArrayOfFoldersType Complex Type]SearchFolder represents a search folder that is contained in a mailbox.");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R536");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R536
            // Set a property on a CalendarFolderType folder successfully, indicates that Folder represents a regular folder.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                536,
                @"[In t:SetFolderFieldType Complex Type]CalendarFolder represents a folder that primarily contains calendar items. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R538");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R538
            // Set a property on a ContactFolderType folder successfully, indicates that Folder represents a regular folder.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[1].ResponseClass,
                538,
                @"[In t:SetFolderFieldType Complex Type]ContactsFolder represents a Contacts folder in a mailbox. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R542");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R542
            // Set a property on a TaskFolderType folder successfully, indicates that Folder represents a regular folder.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[2].ResponseClass,
                542,
                @"[In t:SetFolderFieldType Complex Type]TasksFolder represents a Tasks folder that is contained in a mailbox. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R540");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R540
            // Set a property on a SearchFolderType folder successfully, indicates that Folder represents a regular folder.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                updateFolderResponse.ResponseMessages.Items[3].ResponseClass,
                540,
                @"[In t:SetFolderFieldType Complex Type]SearchFolder represents a search folder that is contained in a mailbox. ");
        }

        /// <summary>
        /// This test case verifies requirements related to UpdateFolder operation via creating a folder and updating it with DeleteFolderFieldType property.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC04_UpdateFolderWithDeleteFolderFieldType()
        {
            #region Create a new folder in the inbox folder

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].CanCreateSubFolders = true;
            permissionSet.Permissions[0].CanCreateSubFoldersSpecified = true;
            permissionSet.Permissions[0].IsFolderOwner = true;
            permissionSet.Permissions[0].IsFolderOwnerSpecified = true;
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

            #region Update Folder Operation.

            // UpdateFolder request to delete folder permission value.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "DeleteFolderField", newFolderId);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 1, this.Site);

            #endregion

            #region Switch user

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

            // Permission have been deleted so create operation should be failed.
            bool isPermissionDeleted = createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass.Equals(ResponseClassType.Error);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R583");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R583
            // One permission set which set in CreateFolder is deleted when calling UpdateFolder.
            this.Site.CaptureRequirementIfIsTrue(
                isPermissionDeleted,
                583,
                @"[In t:DeleteFolderFieldType Complex Type]The DeleteFolderFieldType complex type represents an UpdateFolder operation to delete a property from a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5261");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5261
            this.Site.CaptureRequirementIfIsTrue(
                isPermissionDeleted,
                5261,
                @"[In t:NonEmptyArrayOfFolderChangeDescriptionsType Complex Type]DeleteFolderField represents an UpdateFolder operation to delete a property from a folder.");
        }

        /// <summary>
        /// This test case verifies requirement related to ErrorCode of UpdateFolder operation via creating a folder and updating it with multiple properties.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC05_UpdateFolderWithMultipleProperties()
        {
            #region Create a new folder in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            FolderIdType folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(folderId);

            #endregion

            #region Update Folder Operation.

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", folderId);

            // In order to verify MS-OXWSCDATA_R335, add another property(PermissionSet) into UpdateFolder request.
            FolderType updatingFolder = (FolderType)((SetFolderFieldType)updateFolderRequest.FolderChanges[0].Updates[0]).Item1;
            updatingFolder.PermissionSet = new PermissionSetType();
            updatingFolder.PermissionSet.Permissions = new PermissionType[1];
            updatingFolder.PermissionSet.Permissions[0] = new PermissionType();
            updatingFolder.PermissionSet.Permissions[0].CanCreateSubFolders = true;
            updatingFolder.PermissionSet.Permissions[0].CanCreateSubFoldersSpecified = true;
            updatingFolder.PermissionSet.Permissions[0].IsFolderOwner = true;
            updatingFolder.PermissionSet.Permissions[0].IsFolderOwnerSpecified = true;
            updatingFolder.PermissionSet.Permissions[0].PermissionLevel = new PermissionLevelType();
            updatingFolder.PermissionSet.Permissions[0].PermissionLevel = PermissionLevelType.Custom;
            updatingFolder.PermissionSet.Permissions[0].UserId = new UserIdType();

            updatingFolder.PermissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                1,
                updateFolderResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                updateFolderResponse.ResponseMessages.Items.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R335");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R335
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorIncorrectUpdatePropertyCount,
                updateFolderResponse.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                335,
                @"[In m:ResponseCodeType Simple Type] The value ""ErrorIncorrectUpdatePropertyCount"" specifies that each change description in an UpdateItem or UpdateFolder method call MUST list only one property to be updated.");

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to UpdateFolder operation via updating a distinguished folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S06_TC06_UpdateDistinguishedFolder()
        {
            #region Get the sent items folder.

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.sentitems;

            // GetFolder request.
            GetFolderType getSentItemsFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderId);

            GetFolderResponseType getSentItemsFolderResponse = this.FOLDAdapter.GetFolder(getSentItemsFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getSentItemsFolderResponse, 1, this.Site);

            // Variable to save the folder.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getSentItemsFolderResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            #region Update Folder Operation.

            // UpdateFolder request to delete folder permission value.
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "DeleteFolderField", folderInfo.FolderId);

            // Set change key value.
            folderId.ChangeKey = folderInfo.FolderId.ChangeKey;
            updateFolderRequest.FolderChanges[0].Item = folderId;

            // Update the specific folder's properties.
            UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R9101");

            // Distinguished folder id set and update folder return a successfully, this requirement can be captured.
            this.Site.CaptureRequirement(
                9101,
                @"[In t:FolderChangeType Complex Type]DistinguishedFolderId specifies an identifier for a distinguished folder.");
        }
        #endregion
    }
}