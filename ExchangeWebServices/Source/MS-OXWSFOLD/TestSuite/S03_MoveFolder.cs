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
    /// This scenario is designed to verify MoveFolder operation.
    /// </summary>
    [TestClass]
    public class S03_MoveFolder : TestSuiteBase
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
        /// This test case verifies requirements of all properties via creating two nested folders and moving the high-level folder from Drafts folder to Inbox folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S03_TC01_FolderPropertiesAfterMoved()
        {
            #region Create two nested folders with an item in the high-level one into the "drafts" folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.drafts.ToString(), new string[] { "ForMoveFolder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Create a message into this folder
            string itemName = Common.GenerateResourceName(this.Site, "Test Mail");
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName);
            this.NewCreatedItemIds.Add(itemId);

            // Create sub folder request.
            CreateFolderType createSubFolderRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "SubFolder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createNewFolderResponse = this.FOLDAdapter.CreateFolder(createSubFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createNewFolderResponse, 1, this.Site);

            FolderIdType subFolderId = ((FolderInfoResponseMessageType)createNewFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(subFolderId);

            #endregion

            #region Move the new created folder to the inbox folder

            // MoveFolder request.
            MoveFolderType moveFolderRequest = new MoveFolderType();

            // Set the request's folderId field.
            moveFolderRequest.FolderIds = new BaseFolderIdType[1];
            moveFolderRequest.FolderIds[0] = newFolderId;

            // Set the request's destFolderId field.
            DistinguishedFolderIdType toFolderId = new DistinguishedFolderIdType();
            toFolderId.Id = DistinguishedFolderIdNameType.inbox;
            moveFolderRequest.ToFolderId = new TargetFolderIdType();
            moveFolderRequest.ToFolderId.Item = toFolderId;

            // Move the specified folder.
            MoveFolderResponseType moveFolderResponse = this.FOLDAdapter.MoveFolder(moveFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveFolderResponse, 1, this.Site);

            FolderIdType movedFolderId = ((FolderInfoResponseMessageType)moveFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(movedFolderId);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4314");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4314
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                moveFolderResponse.ResponseMessages.Items[0].ResponseCode,
                4314,
                @"[In MoveFolder Operation]A successful MoveFolder operation request returns a MoveFolderResponse element with the ResponseCode element of the MoveFolderResponse element set to ""NoError"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R43144");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R43144
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                moveFolderResponse.ResponseMessages.Items[0].ResponseClass,
                43144,
                @"[In MoveFolder Operation]A successful MoveFolder operation request returns a MoveFolderResponse element with the ResponseClass attribute of the MoveFolderResponseMessage element set to ""Success"".");

            #region Get the inbox folder's folder id

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.inbox;

            // GetFolder request.
            GetFolderType getInboxFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folder);

            // Get the Inbox folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getInboxFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Variable to save folder.
            FolderInfoResponseMessageType inboxFolder = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            BaseFolderType inboxFolderInfo = (BaseFolderType)inboxFolder.Folders[0];

            // Save the inbox's folder id.
            FolderIdType inboxFolderId = inboxFolderInfo.FolderId;

            #endregion

            #region Get the new created folder after moved to inbox folder

            // GetFolder request.
            GetFolderType getSubFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, movedFolderId);

            // Get the specified folder.
            GetFolderResponseType getSubFolderResponse = this.FOLDAdapter.GetFolder(getSubFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getSubFolderResponse, 1, this.Site);

            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getSubFolderResponse.ResponseMessages.Items[0];
            BaseFolderType folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R595");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R595
            // Since one message in the folder before move, if TotalCount for the folder is 1 after move, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                folderInfo.TotalCount,
                595,
                @"[In MoveFolder Operation]The contents of the folder move with the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R461");

            // The moved folder can be gotten successfully, so the specified folder was moved.
            this.Site.CaptureRequirement(
                461,
                @"[In m:MoveFolderType Complex Type]The MoveFolderType complex type specifies a request message to move folders in a mailbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R429");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R429
            this.Site.CaptureRequirementIfAreEqual<string>(
                inboxFolderId.Id,
                folderInfo.ParentFolderId.Id,
                429,
                @"[In MoveFolder Operation]The MoveFolder operation moves folders from a specified parent folder and puts them in another parent folder.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R594");

            // Verify MS-OXWSFOLD_R594.
            // Verify if the values of FolderClass and DisplayName property are changed after the folder being moved.
            bool isVerifyR594 = folderInfo.FolderClass == "IPF.MyCustomFolderClass" &&
                folderInfo.DisplayName == createFolderRequest.Folders[0].DisplayName;

            Site.Log.Add(
                LogEntryKind.Debug,
               "FolderClass expected to be \"IPF.MyCustomFolderClass\" and actual is {0};\n" +
               "DisplayName expected to be {1} and actual is {2};\n ",
               folderInfo.FolderClass,
               createFolderRequest.Folders[0].DisplayName,
               folderInfo.DisplayName);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR594,
                594,
                @"[In MoveFolder Operation]The properties FolderClass and DisplayName of the folder move with the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7501");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7501
            // Only one child folder was created in the folder.
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                folderInfo.ChildFolderCount,
                7501,
                @"[In t:BaseFolderType Complex Type]ChildFolderCount specifies the total number of child folders in a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5912");

            // Child folder count is returned from server so this requirement can be captured.
            this.Site.CaptureRequirement(
                5912,
                @"[In t:BaseFolderType Complex Type]This property[ChildFolderCount] is returned in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R75");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R75
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.ChildFolderCount,
                75,
                @"[In t:BaseFolderType Complex Type]The type of element ChildFolderCount is xs:int.");

            #region Get subfolder in the moved folder

            // GetFolder request.
            GetFolderType getSubFolderAfterMovedRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, subFolderId);

            // Get the specified folder.
            GetFolderResponseType getFolderAfterMovedResponse = this.FOLDAdapter.GetFolder(getSubFolderAfterMovedRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderAfterMovedResponse, 1, this.Site);

            allFolders = (FolderInfoResponseMessageType)getFolderAfterMovedResponse.ResponseMessages.Items[0];
            folderInfo = (BaseFolderType)allFolders.Folders[0];

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R596");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R596
            this.Site.CaptureRequirementIfAreEqual<string>(
                movedFolderId.ToString(),
                folderInfo.ParentFolderId.ToString(),
                596,
                @"[In MoveFolder Operation]The subfolders of the folder move with the folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R68");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R68
            // Parent folder id is returned from server, and schema is verified in adapter so this requirement can be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                folderInfo.ParentFolderId,
                68,
                @"[In t:BaseFolderType Complex Type]The type of element ParentFolderId is t:FolderIdType.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6801");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6801
            this.Site.CaptureRequirementIfAreEqual<string>(
                NewCreatedFolderIds[0].ToString(),
                folderInfo.ParentFolderId.ToString(),
                6801,
                @"[In t:BaseFolderType Complex Type]ParentFolderId specifies the folder identifier and change key for the parent folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to moving multiple folders in MoveFolder operation.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S03_TC02_MoveMultipleFolders()
        {
            #region Create multiple folders in the "drafts" folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.drafts.ToString(), new string[] { "ForMoveFolder1", "ForMoveFolder2" }, new string[] { "IPF.MyCustomFolderClass", "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            FolderIdType newFolderId1 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            FolderIdType newFolderId2 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[1]).Folders[0].FolderId;

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 2, this.Site);

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId1);
            this.NewCreatedFolderIds.Add(newFolderId2);

            #endregion

            #region Move the new created folder to the inbox folder

            // MoveFolder request.
            MoveFolderType moveFolderRequest = new MoveFolderType();

            // Set the request's folderId field.
            moveFolderRequest.FolderIds = new BaseFolderIdType[2];
            moveFolderRequest.FolderIds[0] = newFolderId1;
            moveFolderRequest.FolderIds[1] = newFolderId2;

            // Set the request's destFolderId field.
            DistinguishedFolderIdType toFolderId = new DistinguishedFolderIdType();
            toFolderId.Id = DistinguishedFolderIdNameType.inbox;
            moveFolderRequest.ToFolderId = new TargetFolderIdType();
            moveFolderRequest.ToFolderId.Item = toFolderId;

            // Move the specified folder.
            MoveFolderResponseType moveFolderResponse = this.FOLDAdapter.MoveFolder(moveFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveFolderResponse, 2, this.Site);

            #endregion
        }
        #endregion
    }
}