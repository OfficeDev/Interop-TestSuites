namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to verify CreateFolder and DeleteFolder operation.
    /// </summary>
    [TestClass]
    public class S01_CreateDeleteFolder : TestSuiteBase
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
        /// This test case verifies requirements related to creating and deleting a folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC01_CreateDeleteFolder()
        {
            #region Create a new folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            //Get DisplayName in the request.
            string displayName = createFolderRequest.Folders[0].DisplayName;

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R71021");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R71021
            //displayNames is not null, so this requirement can be verified directly.
            this.Site.CaptureRequirementIfIsNotNull(
                displayName,
                71021,
                @"[In t:BaseFolderType Complex Type]This element[DisplayName] is required in a CreateFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7102");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7102
            //MS-OXWSFOLD_71021 is verified, so this requirement can be verified directly.
            this.Site.CaptureRequirement(
                7102,
                @"[In t:BaseFolderType Complex Type]This element[DisplayName] can be present.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R216");

            // Parent folder already exists and folder can be created this requirement can be captured.
            this.Site.CaptureRequirement(
                216,
                @"[In CreateFolder Operation]Before a folder can be created, the parent folder MUST already exist.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R67");

            // FolderId is not set during folder creation and if it exists in the response, this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                newFolderId,
                67,
                @"[In t:BaseFolderType Complex Type]This element[FolderId] can be present and cannot be set during folder creation.");

            #region Create an item

            string itemName = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemId = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site), newFolderId.Id, itemName);
            Site.Assert.IsNotNull(itemId, "Item should be created successfully!");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5890");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5890
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createFolderResponse.ResponseMessages.Items[0].ResponseClass,
                5890,
                @"[In t:BaseFolderType Complex Type]The folders class can be a custom class.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2202
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createFolderResponse.ResponseMessages.Items[0].ResponseClass,
                2202,
                @"[In CreateFolder Operation]A successful CreateFolder operation request returns a CreateFolderResponse element with the ResponseClass attribute of the CreateFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R560");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R560
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createFolderResponse.ResponseMessages.Items[0].ResponseClass,
                560,
                @"[In m:CreateFolderType Complex Type]The CreateFolderType complex type specifies a request message to create a folder in the server database.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R589201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R589201
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                589201,
                @"[In t:BaseFolderType Complex Type]This element [FolderClass] is present, server responses NO_ERROR.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R22202");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R22202
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createFolderResponse.ResponseMessages.Items[0].ResponseCode,
                22202,
                @"[In CreateFolder Operation]A successful CreateFolder operation request returns a CreateFolderResponse element with the ResponseCode element of the CreateFolderResponse element set to ""NoError"".");

            #region Get the inbox folder

            DistinguishedFolderIdType folder = new DistinguishedFolderIdType();
            folder.Id = (DistinguishedFolderIdNameType)DistinguishedFolderIdNameType.inbox;

            // GetFolder request.
            GetFolderType getInboxRequest = this.GetGetFolderRequest(DefaultShapeNamesType.IdOnly, folder);

            // Get the inbox folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getInboxRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R28");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R28
            // If gotten folder is FolderType, then this requirement will be captured.
            this.Site.CaptureRequirementIfIsInstanceOfType(
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0],
                typeof(FolderType),
                28,
                @"[In t:ArrayOfFoldersType Complex Type]The type of element Folder is t:FolderType (section 2.2.4.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2802");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2802
            // If gotten folder is FolderType, then this requirement will be captured.
            this.Site.CaptureRequirementIfIsInstanceOfType(
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0],
                typeof(FolderType),
                2802,
                @"[In t:ArrayOfFoldersType Complex Type]Folder represents a regular folder in the server database.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2531");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2531
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getFolderResponse.ResponseMessages.Items[0].ResponseClass,
                2531,
                @"[In m:CreateFolderType Complex Type]Folders Represents an array of folders to be created.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R252");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R252
            this.Site.CaptureRequirementIfAreEqual<string>(
                ((FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId.Id,
                ((FolderType)((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0]).ParentFolderId.Id,
                252,
                @"[In m:CreateFolderType Complex Type]ParentFolderId is the identifier of the folder that will contain the newly created folder. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6802");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6802
            this.Site.CaptureRequirementIfIsNotNull(
                ((FolderType)((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0]).ParentFolderId.Id,
                6802,
                @"[In t:BaseFolderType Complex Type]This element[ParentFolderId] can be present and cannot be set during folder creation.");


            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7201
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].TotalCount,
                7201,
                @"[In t:BaseFolderType Complex Type]TotalCount specifies the total number of items in a folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R72");

            // Since R7201 can be verified, this requirement will be captured.
            this.Site.CaptureRequirement(
                72,
                @"[In t:BaseFolderType Complex Type]The type of element TotalCount is xs:int ([XMLSCHEMA2]).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5900");

            // Since R7201 can be verified, this requirement will be captured.
            this.Site.CaptureRequirement(
                5900,
                @"[In t:BaseFolderType Complex Type]This property[TotalCount] is returned in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R42101");

            // "AllProperties" is set in request, and the operation executes successfully, this requirement can be covered.
            this.Site.CaptureRequirement(
                42101,
                @"[In t:DefaultShapeNamesType Simple Type] A value of ""AllProperties"" [in DefaultShapeNamesType] specifies all the properties that are defined for the AllProperties shape to include in the response.");

            #region Delete the created folder

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, newFolderId);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3104");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3104
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                3104,
                @"[In DeleteFolder Operation]A successful DeleteFolder operation request returns a DeleteFolderResponse element with the ResponseClass attribute of the DeleteFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R31044");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R31044
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                deleteFolderResponse.ResponseMessages.Items[0].ResponseCode,
                31044,
                @"[In DeleteFolder Operation]A successful DeleteFolder operation request returns a DeleteFolderResponse element with the ResponseCode element of the DeleteFolderResponse element set to ""NoError"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R568");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R568
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                568,
                @"[In m:DeleteFolderType Complex Type]The DeleteFolderType complex type specifies a request message to delete folders from a mailbox.");

            // The folder has been deleted, so its folder id has disappeared.
            this.NewCreatedFolderIds.Remove(newFolderId);

            #endregion

            #region Get the folder again to see whether it's been deleted.

            // Get the new created folder.
            GetFolderResponseType getDeletedFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 getDeletedFolderResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getDeletedFolderResponse.ResponseMessages.Items.GetLength(0));
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R306");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R306
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Success,
                getDeletedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                306,
                @"[In DeleteFolder Operation]The DeleteFolder operation deletes folders from a mailbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R341");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R341
            // If call GetFolder with the FolderIds unsuccessful, indicates that the folders are deleted from the mailbox.
            this.Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Success,
                getDeletedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                341,
                @"[In m:DeleteFolderType Complex Type]FolderIds is an array of folders to be deleted from a mailbox. ");
        }

        /// <summary>
        /// This test case verifies that error occurred when creating a folder that has existed.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC02_CreateExistFolder()
        {
            #region Create a folder with the name that exists in the inbox folder

            // CreatFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Variable to save the new folder's folder id
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);

            // Create the same folder again that already exists.
            CreateFolderResponseType createNewFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 createNewFolderResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createNewFolderResponse.ResponseMessages.Items.GetLength(0));
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R217.");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R217
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createNewFolderResponse.ResponseMessages.Items[0].ResponseClass,
                217,
                @"[In CreateFolder Operation]Trying to create a folder that already exists results in an error.");
        }

        /// <summary>
        /// This test case verifies that default folder cannot be deleted. 
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC03_DeleteDefaultFolder()
        {
            #region Delete the default folder

            // Default folder Id.
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;

            // Delete folder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.SoftDelete, folderId);

            // Delete the specific folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 deleteFolderResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 deleteFolderResponse.ResponseMessages.Items.GetLength(0));
            #endregion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R308.");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R308
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                308,
                @"[In DeleteFolder Operation]This operation cannot delete default folders, such as the Inbox folder or the Deleted Items folder.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R31045.");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R31045
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                31045,
                @"[In DeleteFolder Operation]An unsuccessful DeleteFolder operation request returns a DeleteFolderResponse element with the ResponseClass attribute of the DeleteFolderResponseMessage element set to ""Error"".");
        }

        /// <summary>
        /// This test case verifies requirements related to creating a managed folder.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC04_CreateManagedFolder()
        {
            #region Create a new managed folder

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

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2692");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2692
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                2692,
                @"[In CreateManagedFolder Operation]A successful CreateManagedFolder operation request returns a CreateManagedFolderResponse element with the ResponseClass attribute of the CreateManagedFolderResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R26992");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R26992
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseCode,
                26992,
                @"[In CreateManagedFolder Operation]A successful CreateManagedFolder operation request returns a CreateManagedFolderResponse element with the ResponseCode element of the CreateManagedFolderResponse element set to ""NoError"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R565");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R565
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                565,
                @"[In m:CreateManagedFolderRequestType Complex Type]The CreateManagedFolderRequestType complex type specifies a request message to create a managed folder in a server database.");

            #region Get the new created managed folder to verify ManagedFolderInformationType

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2961");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2961
            this.Site.CaptureRequirementIfAreEqual<string>(
                createManagedFolderRequest.FolderNames[0],
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].DisplayName,
                2961,
                @"[In m:CreateManagedFolderRequestType Complex Type]FolderNames specifies an array of managed folders to add to a mailbox. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R298");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R298
            // Because the email address is current user, so if this folder can be gotten by current user, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                getFolderResponse.ResponseMessages.Items[0].ResponseClass,
                298,
                @"[In m:CreateManagedFolderRequestType Complex Type]Mailbox specifies the e-mail address of the mailbox in which the managed folders are added. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R5922");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R5922
            // If ManagedFolderInformation is not null, this requirement will be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation,
                5922,
                @"[In t:BaseFolderType Complex Type]This property[ManagedFolderInformation] is returned in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1111");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1111
            // Comment returned from server this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation.Comment,
                1111,
                @"[In t:ManagedFolderInformationType Complex Type]Comment is a comment that is associated with a managed folder.");

            #region Get the new created managed folder's parent folder

            FolderIdType parentFolderId = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId;

            // GetFolder request.
            GetFolderType getParentFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, parentFolderId);

            // Get the specified managed folder's parent folder.
            GetFolderResponseType getParentFolderResponse = this.FOLDAdapter.GetFolder(getParentFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getParentFolderResponse, 1, this.Site);
            #endregion

            Site.Assert.AreEqual<bool>(false, ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation.IsManagedFoldersRoot, "Folder should not be a root managed folder!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R10811");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R10811
            // Managed folder's parent folder is still a managed folder and its IsManagedFolderRoot is false, this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                ((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation,
                10811,
                @"[In t:ManagedFolderInformationType Complex Type][IsManagedFoldersRoot]A value of ""false"" indicates that the managed is not the root managed folder. ");

            bool isRootManagedFolderFound = true;

            // Loop to find the root managed folder.
            while (((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation.IsManagedFoldersRoot != true)
            {
                parentFolderId = ((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId;
                if (parentFolderId == null)
                {
                    isRootManagedFolderFound = false;
                    break;
                }

                getParentFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, parentFolderId);
                getParentFolderResponse = this.FOLDAdapter.GetFolder(getParentFolderRequest);

                // Check the response.
                Common.CheckOperationSuccess(getParentFolderResponse, 1, this.Site);

                // If the folder is not a managed folder, break the loop
                if (((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation == null)
                {
                    isRootManagedFolderFound = false;
                    break;
                }
            }

            Site.Assert.AreEqual<bool>(true, isRootManagedFolderFound, "The root managed folder should exist!");

            #region Get the root managed folder's parent folder

            FolderIdType ancestorFolderId = ((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId;

            // GetFolder request.
            GetFolderType getAncestorFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, ancestorFolderId);
            GetFolderResponseType getAncestorFolderResponse = this.FOLDAdapter.GetFolder(getAncestorFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getAncestorFolderResponse, 1, this.Site);

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1082");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1082
            // Root managed folder's parent folder is not a managed folder and Managed folder's parent folder's IsManagedFolderRoot is true, this requirement can be captured.
            this.Site.CaptureRequirementIfIsNull(
                ((FolderInfoResponseMessageType)getAncestorFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation,
                1082,
                @"[In t:ManagedFolderInformationType Complex Type][IsManagedFoldersRoot]A value of ""true"" indicates that the managed is the root managed folder. ");
        }

        /// <summary>
        /// This test case verifies that DeleteFolder with delete type set to HardDelete. 
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC05_DeleteFolderHardDelete()
        {
            #region Create a new folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            #endregion

            #region Hard delete the created folder in inbox

            // Delete folder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, newFolderId);

            // Delete the specific folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, getInboxFolderResponse.ResponseMessages.Items[0].ResponseClass, "Folder should not be found!");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R34301");

            // Specific folder was not found in mailbox. The folder was deleted from the store.
            this.Site.CaptureRequirement(
                34301,
                @"[In m:DeleteFolderType Complex Type]DeleteType which value is HardDelete specifies that a folder is permanently removed from the store.");
        }

        /// <summary>
        /// This test case verifies that DeleteFolder with delete type set to MoveToDeleteItems. 
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC06_DeleteFolderMoveToDeleteItems()
        {
            #region Create a new folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Ensure this folder will be deleted in clean up.
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Delete the created folder in inbox

            // Delete folder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.MoveToDeletedItems, newFolderId);

            // Delete the specific folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);

            #endregion

            #region Get the new created folder.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getNewCreatedFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getNewCreatedFolderResponse, 1, this.Site);

            #endregion

            #region Get new created folder's parent folder

            GetFolderType getParentFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, ((FolderInfoResponseMessageType)getNewCreatedFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId);
            GetFolderResponseType getParentFolderResponse = this.FOLDAdapter.GetFolder(getParentFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getParentFolderResponse, 1, this.Site);

            string folderDisplayName = ((FolderInfoResponseMessageType)getParentFolderResponse.ResponseMessages.Items[0]).Folders[0].DisplayName;

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R34302");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R34302
            this.Site.CaptureRequirementIfAreEqual<string>(
                "Deleted Items",
                folderDisplayName,
                34302,
                @"[In m:DeleteFolderType Complex Type ]DeleteType which value is MoveToDeletedItems specifies that a folder is moved to the Deleted Items folder.");
        }

        /// <summary>
        /// This test case verifies the error code ErrorNonExistentMailbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC07_CreateManagedFolderInNonexistingMailbox()
        {
            #region Create a managed folder

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site));

            // Set mailbox value.
            createManagedFolderRequest.Mailbox = new EmailAddressType();

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, createManagedFolderResponse.ResponseMessages.Items[0].ResponseClass, "Managed folder should not be created if the e-mail address is empty in the CreateManagedFolder method!");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1453");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1453
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorNonExistentMailbox,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                1453,
                @"[In m:ResponseCodeType Simple Type] [The value ""ErrorNonExistentMailbox"" ] Specifies one of the following: 1) The e-mail address is empty in the CreateManagedFolder method.");
        }

        /// <summary>
        /// This test case verifies the error code ErrorDuplicateInputFolderNames.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC08_CreateDuplicateManagedFolder()
        {
            #region Create two new managed folders

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(new string[] { Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site), Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site) });

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, createManagedFolderResponse.ResponseMessages.Items[0].ResponseClass, "Managed folder should be created failed when duplicate folder names are passed to the CreateManagedFolder method!");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R301");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R301
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorDuplicateInputFolderNames,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                301,
                @"[In m:ResponseCodeType Simple Type] [The value ""ErrorDuplicateInputFolderNames"" represent ] Occurs when there are duplicate folder names in the array that was passed to the CreateManagedFolder method.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R26993");

            // Verify MS-OXWSCDATA requirement: MS-OXWSFOLD_R26993
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createManagedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                26993,
                @"[In CreateManagedFolder Operation]An unsuccessful CreateManagedFolder operation request returns a CreateManagedFolderResponse element with the ResponseClass attribute of the CreateManagedFolderResponseMessage element set to ""Error"".");
        }

        /// <summary>
        /// This test case verifies requirements related to create and delete multiple folders.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC09_CreateDeleteMultipleFolders()
        {
            #region Create multiple new folders in the inbox folder.

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder1", "Custom Folder2" }, new string[] { "IPF.MyCustomFolderClass", "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 2, this.Site);

            // Create child folders and save their ids.
            FolderIdType newFolderId1 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            FolderIdType newFolderId2 = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[1]).Folders[0].FolderId;

            #endregion

            #region Get the new created folders.

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.Default, new FolderIdType[] { newFolderId1, newFolderId2 });

            // Get the new created child folders.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 2, this.Site);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R42102");

            // "Default" is set in request, and the operation executes successfully, this requirement can be covered.
            this.Site.CaptureRequirement(
                42102,
                @"[In t:DefaultShapeNamesType Simple Type] A value of ""Default"" [in DefaultShapeNamesType] specifies a set of properties that are defined as the default for the item or folder to include in the response.");
            #endregion

            #region Delete the created folders.

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.HardDelete, new FolderIdType[] { newFolderId1, newFolderId2 });

            // Delete the specified child folders.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 2, this.Site);

            #endregion
        }

        /// <summary>
        /// This test case verifies creating multiple managed folders at same time.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC10_CreateMultipleManagedFolder()
        {
            #region Create two new managed folders

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(new string[] { Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site), Common.GetConfigurationPropertyValue("ManagedFolderName2", this.Site) });

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createManagedFolderResponse, 2, this.Site);

            FolderIdType newFolderId1 = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            FolderIdType newFolderId2 = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[1]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId1);
            this.NewCreatedFolderIds.Add(newFolderId2);

            #endregion

            #region Get the new created managed folder

            // Create a FolderId array to save newFolderId1 and 2.
            FolderIdType[] folderIds = new FolderIdType[2] { newFolderId1, newFolderId2 };

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, folderIds);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 2, this.Site);

            #endregion

            // Managed folder ids.
            string firstManagedFolderId = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation.ManagedFolderId;
            string secondManagedFolderId = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[1]).Folders[0].ManagedFolderInformation.ManagedFolderId;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1091 \nThe ManageFolderId of the first managed folder is: {0}, and the ManageFolderId of the second managed folder is: {1}. If they are equal to each other, this requirement will be captured", firstManagedFolderId, secondManagedFolderId);

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1091
            bool isVerifiedR1091 = firstManagedFolderId != secondManagedFolderId;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1091,
                1091,
                @"[In t:ManagedFolderInformationType Complex Type]ManageFolderId is the unique identifier of a managed folder.");
        }

        /// <summary>
        /// This test case verifies requirements related to creating and deleting a calendar folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC11_CreateDeleteCalendarFolder()
        {
            #region Create a calendar folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Calendar Folder" }, new string[] { "IPF.Appointment" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderId.Id, allFolders.Folders[0].FolderId.Id, "The calendar folder should be created successfully in inbox.");

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
        }

        /// <summary>
        /// This test case verifies requirements related to creating and deleting a tasks-folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC12_CreateDeleteTasksFolder()
        {
            #region Create a tasks folder in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Tasks Folder" }, new string[] { "IPF.Task" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderId.Id, allFolders.Folders[0].FolderId.Id, "The tasks folder should be created successfully in inbox.", null);

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
        }

        /// <summary>
        /// This test case verifies requirements related to creating and deleting a contacts-folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC13_CreateDeleteContactsFolder()
        {
            #region Create a contacts folder in the inbox folder

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
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Contacts Folder" }, new string[] { "IPF.Contact" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createFolderResponse, 1, this.Site);

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created folder

            // GetFolder request.
            GetFolderType getNewFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the new created folder.
            GetFolderResponseType getInboxFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getInboxFolderResponse, 1, this.Site);

            // Verify the folder is created successfully.
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getInboxFolderResponse.ResponseMessages.Items[0];
            Site.Assert.AreEqual<string>(newFolderId.Id, allFolders.Folders[0].FolderId.Id, "The contacts folder should be created successfully in inbox.");

            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R217");

            // PermissionSet value is set in create folder operation and schema is verified in adapter, so this requirement can be captured.
            this.Site.CaptureRequirement(
                "MS-OXWSCONT",
                217,
                @"[In t:ContactsFolderType Complex Type] The type of the element of PermissionSet is t:PermissionSetType ([MS-OXWSFOLD] section 2.2.4.12)");

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
        }

        /// <summary>
        /// This test case verifies requirements related to read-only elements.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC14_CreateFolderWithReadOnlyElements()
        {
            #region Create a new folder with TotalCount in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequestWithTotalCount = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Add read-only elements.
            createFolderRequestWithTotalCount.Folders[0].TotalCountSpecified = true;
            createFolderRequestWithTotalCount.Folders[0].TotalCount = 3;

            // Create a new folder.
            CreateFolderResponseType createFolderResponseWithTotalCount = this.FOLDAdapter.CreateFolder(createFolderRequestWithTotalCount);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 createFolderResponseWithTotalCount.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createFolderResponseWithTotalCount.ResponseMessages.Items.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R73");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R73
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseWithTotalCount.ResponseMessages.Items[0].ResponseClass,
                73,
                @"[In t:BaseFolderType Complex Type]This property[TotalCount] MUST be read-only for a client.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R22203");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R22203
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseWithTotalCount.ResponseMessages.Items[0].ResponseClass,
                22203,
                @"[In CreateFolder Operation]An unsuccessful CreateFolder operation request returns a CreateFolderResponse element with the ResponseClass attribute of the CreateFolderResponseMessage element set to ""Error"".");

            #endregion

            #region Create a new folder with ChildFolderCount in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequestWithChildFolderCount = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Add read-only elements.
            createFolderRequestWithChildFolderCount.Folders[0].ChildFolderCountSpecified = true;
            createFolderRequestWithChildFolderCount.Folders[0].ChildFolderCount = 3;

            // Create a new folder.
            CreateFolderResponseType createFolderResponseWithChildFolderCount = this.FOLDAdapter.CreateFolder(createFolderRequestWithChildFolderCount);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 createFolderResponseWithChildFolderCount.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createFolderResponseWithChildFolderCount.ResponseMessages.Items.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R76");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R76
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseWithChildFolderCount.ResponseMessages.Items[0].ResponseClass,
                76,
                @"[In t:BaseFolderType Complex Type]This property[ChildFolderCount] MUST be read-only for a client .");

            #endregion

            #region Create a new folder with ManagedFolderInformation in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequestWithManagedFolderInformation = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Add read-only elements.
            createFolderRequestWithManagedFolderInformation.Folders[0].ManagedFolderInformation = new ManagedFolderInformationType()
            {
                CanDeleteSpecified = true,
                CanDelete = true
            };

            // Create a new folder.
            CreateFolderResponseType createFolderResponseWithManagedFolderInformation = this.FOLDAdapter.CreateFolder(createFolderRequestWithManagedFolderInformation);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 createFolderResponseWithManagedFolderInformation.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createFolderResponseWithManagedFolderInformation.ResponseMessages.Items.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R80");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R80
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseWithManagedFolderInformation.ResponseMessages.Items[0].ResponseClass,
                80,
                @"[In t:BaseFolderType Complex Type]This property[ManagedFolderInformation] MUST be read-only for a client.");

            #endregion

            #region Create a new folder with EffectiveRights in the inbox folder

            // CreateFolder request.
            CreateFolderType createFolderRequestWithEffectiveRights = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Add read-only elements.
            createFolderRequestWithEffectiveRights.Folders[0].EffectiveRights = new EffectiveRightsType()
            {
                Delete = true
            };

            // Create a new folder.
            CreateFolderResponseType createFolderResponseWithEffectiveRights = this.FOLDAdapter.CreateFolder(createFolderRequestWithEffectiveRights);

            // Check the length.
            Site.Assert.AreEqual<int>(
                 1,
                 createFolderResponseWithEffectiveRights.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createFolderResponseWithEffectiveRights.ResponseMessages.Items.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R83");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R83
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createFolderResponseWithEffectiveRights.ResponseMessages.Items[0].ResponseClass,
                83,
                @"[In t:BaseFolderType Complex Type]This property[EffectiveRights] MUST be read-only for a client.");

            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to get folder failed elements.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC15_GetFolderFailed()
        {
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

            #region Get the deleted folder
            GetFolderType getFolderRequest = this.GetGetFolderRequest( DefaultShapeNamesType.AllProperties, newFolderId);

            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R38645");

            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                getFolderResponse.ResponseMessages.Items[0].ResponseClass,
                38645,
                @"[In GetFolder Operation]An unsuccessful GetFolder operation request returns a GetFolderResponse element with the ResponseClass attribute of the GetFolderResponseMessage element set to ""Error"".");
            #endregion
        }

        /// <summary>
        /// This test case verifies requirements related to soft delete a folder in Inbox.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC16_SoftDeleteFolder()
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

            #region Delete the created folder

            // DeleteFolder request.
            DeleteFolderType deleteFolderRequest = this.GetDeleteFolderRequest(DisposalType.SoftDelete, newFolderId);

            // Delete the specified folder.
            DeleteFolderResponseType deleteFolderResponse = this.FOLDAdapter.DeleteFolder(deleteFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteFolderResponse, 1, this.Site);

            // The folder has been deleted, so its folder id has disappeared.
            this.NewCreatedFolderIds.Remove(newFolderId);

            #endregion

            #region Get the folder again to see whether it's been deleted.

            // Get the new created folder.
            GetFolderResponseType getDeletedFolderResponse = this.FOLDAdapter.GetFolder(getNewFolderRequest);

            // Check the length.
            Site.Assert.AreEqual<ResponseClassType>(
                 ResponseClassType.Error,
                 getDeletedFolderResponse.ResponseMessages.Items[0].ResponseClass,
                 "The folder should be deleted successfully.");
            #endregion

            #region Find in recoverableitemsdeletions
            ItemIdType findItemID = this.FindItem(DistinguishedFolderIdNameType.recoverableitemsdeletions.ToString(), itemName);
            Site.Assert.IsNotNull(findItemID, "The item in folder which is soft deleted should be found.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R34303");

            // The item exists in the delete folder is found, this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                findItemID,
                34303,
                @"[In m:DeleteFolderType Complex Type ]DeleteType which value is SoftDelete specifies that a folder is moved to the dumpster if the dumpster is enabled.");
            #endregion
        }

        /// <summary>
        /// This test case verifies RenameOrMoveIsTrue.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC17_VerifyRenameOrMoveTrue()
        {
            #region Create a new managed folder
           
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

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);
            #endregion

            #region Get the new created managed folder to verify ManagedFolderInformationType

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);
            #endregion

            #region Rename the new created folder
            if (Common.IsRequirementEnabled(105211, this.Site))
            {
                #region Update Folder Operation.

                // UpdateFolder request.
                UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", newFolderId);

                // Update the specific folder's properties.
                UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

                // Check the response.
                Common.CheckOperationSuccess(updateFolderResponse, 1, this.Site);

                string updateNameInRequest = ((SetFolderFieldType)updateFolderRequest.FolderChanges[0].Updates[0]).Item1.DisplayName;
                #endregion

                #region Get the updated folder.

                // GetFolder request.
                GetFolderType getUpdatedFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

                // Get the updated folder.
                GetFolderResponseType getUpdateFolderResponse = this.FOLDAdapter.GetFolder(getUpdatedFolderRequest);

                // Check the response.
                Common.CheckOperationSuccess(getUpdateFolderResponse, 1, this.Site);

                ManagedFolderInformationType managedFolderInformation = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation;

                Site.Assert.IsTrue(managedFolderInformation.CanRenameOrMove, "The CanRenameOrMove element should be present!");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R105211");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R105211
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    getUpdateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                    105211,
                    @"[In Appendix C: Product Behavior] Implementation does support value of ""true"" for CanRenameOrMove to indicate that the managed folder can be renamed. (Exchange 2013 and above follow this behavior.)");
                #endregion
            }
            #endregion

            #region Move the new created folder to the inbox folder
            if (Common.IsRequirementEnabled(105212, this.Site))
            {
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

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R105212");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R105212
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    moveFolderResponse.ResponseMessages.Items[0].ResponseClass,
                    105212,
                    @"[In Appendix C: Product Behavior] Implementation does support value of ""true"" for CanRenameOrMove to indicate that the managed folder can be moved. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case verifies RenameOrMoveIsFalse.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC18_VerifyRenameOrMoveFalse()
        {
            #region Create a new managed folder

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

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created managed folder to verify ManagedFolderInformationType

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            #endregion

            #region Get the new created managed folder's parent folder

            FolderIdType parentFolderId = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ParentFolderId;

            // GetFolder request.
            GetFolderType getParentFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, parentFolderId);

            // Get the specified managed folder's parent folder.
            GetFolderResponseType getParentFolderResponse = this.FOLDAdapter.GetFolder(getParentFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getParentFolderResponse, 1, this.Site);
            #endregion

            #region Move the new created folder to the inbox folder
            ManagedFolderInformationType managedFolderInformation = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation;

            if (Common.IsRequirementEnabled(1051112, this.Site))
            {
                // MoveFolder request.
                MoveFolderType moveFolderRequest = new MoveFolderType();

                // Set the request's folderId field.
                moveFolderRequest.FolderIds = new BaseFolderIdType[1];
                moveFolderRequest.FolderIds[0] = parentFolderId;

                // Set the request's destFolderId field.
                DistinguishedFolderIdType toFolderId = new DistinguishedFolderIdType();
                toFolderId.Id = DistinguishedFolderIdNameType.inbox;
                moveFolderRequest.ToFolderId = new TargetFolderIdType();
                moveFolderRequest.ToFolderId.Item = toFolderId;

                // Move the specified folder.
                MoveFolderResponseType moveFolderResponse = this.FOLDAdapter.MoveFolder(moveFolderRequest);

                // Check the response.
                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, moveFolderResponse.ResponseMessages.Items[0].ResponseClass, "Managed folder should not be moved");

                Site.Assert.IsTrue(managedFolderInformation.CanRenameOrMoveSpecified && !managedFolderInformation.CanRenameOrMove, "The CanRenameOrMove element should be present!");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1051112");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1051112
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Error,
                    moveFolderResponse.ResponseMessages.Items[0].ResponseClass,
                    1051112,
                    @"[In Appendix C: Product Behavior] Implementation does support value of ""false"" for CanRenameOrMove to indicate that the managed folder can not be moved. (Exchange 2013 and above follow this behavior.)");
                }
            #endregion

            #region Update Folder Operation.
            if (Common.IsRequirementEnabled(1051111, this.Site))
            {
                // UpdateFolder request.
                UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest("Folder", "SetFolderField", parentFolderId);

                // Update the specific folder's properties.
                UpdateFolderResponseType updateFolderResponse = this.FOLDAdapter.UpdateFolder(updateFolderRequest);

                Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Error, updateFolderResponse.ResponseMessages.Items[0].ResponseClass, "Managed folder should not be updated");
                managedFolderInformation = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1051111");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1051111
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Error,
                    updateFolderResponse.ResponseMessages.Items[0].ResponseClass,
                    1051111,
                    @"[In t:ManagedFolderInformationType Complex Type][CanRenameOrMove]A value of ""false"" indicates that the managed folder cannot be renamed [or moved].");
            }
            #endregion
        }

        /// <summary>
        /// This test case verifies StorageQuota not set.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC19_VerifyStorageQuotaNotSet()
        {
            #region Create a new managed folder

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

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);
            #endregion

            #region Get the new created managed folder to verify ManagedFolderInformationType

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            ManagedFolderInformationType managedFolderInformation = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation;

            Site.Assert.IsTrue(managedFolderInformation.HasQuotaSpecified && !managedFolderInformation.HasQuota, "The HasQuota element should be present!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R10711");

            //Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R10711
            this.Site.CaptureRequirementIfAreEqual(
                0,
                managedFolderInformation.StorageQuota,
                10711,
                @"[In t:ManagedFolderInformationType Complex Type][HasQuota]A value of ""false"" indicates that the StorageQuota property was not serialized into the SOAP response.");
            #endregion
        }

        /// <summary>
        /// This test case verifies StorageQuota set.
        /// </summary>
        [TestCategory("MSOXWSFOLD"), TestMethod()]
        public void MSOXWSFOLD_S01_TC20_VerifyStorageQuotaSet()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1121111, this.Site), "Exchange 2010 uses the StorageQuota which is the storage quota for a managed folder.");
 
            #region Create a new managed folder

            CreateManagedFolderRequestType createManagedFolderRequest = this.GetCreateManagedFolderRequest(Common.GetConfigurationPropertyValue("ManagedFolderName1", this.Site));

            // Add an email address into request.
            EmailAddressType mailBox = new EmailAddressType()
            {
                EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            createManagedFolderRequest.Mailbox = mailBox;

            this.FOLDSUTControlAdapter.SetManagedFolderStoreQuota(createManagedFolderRequest.FolderNames[0]);

            // Create the specified managed folder.
            CreateManagedFolderResponseType createManagedFolderResponse = this.FOLDAdapter.CreateManagedFolder(createManagedFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(createManagedFolderResponse, 1, this.Site);

            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createManagedFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;

            // Save the new created managed folder's folder id.
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Get the new created managed folder to verify ManagedFolderInformationType

            // GetFolder request.
            GetFolderType getFolderRequest = this.GetGetFolderRequest(DefaultShapeNamesType.AllProperties, newFolderId);

            // Get the specified managed folder.
            GetFolderResponseType getFolderResponse = this.FOLDAdapter.GetFolder(getFolderRequest);

            // Check the response.
            Common.CheckOperationSuccess(getFolderResponse, 1, this.Site);

            ManagedFolderInformationType managedFolderInformation = ((FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0]).Folders[0].ManagedFolderInformation;

            Site.Assert.IsTrue(managedFolderInformation.HasQuota, "The HasQuota element should be present!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1072");

            //Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1072
            this.Site.CaptureRequirementIfIsTrue(
                managedFolderInformation.HasQuota,
                1072,
                @"[In t:ManagedFolderInformationType Complex Type][HasQuota]A value of ""true"" indicates that the StorageQuota property was serialized into the SOAP response.");

            Site.Assert.IsTrue(managedFolderInformation.StorageQuotaSpecified && managedFolderInformation.StorageQuota != 0, "The StorageQuota element should be present!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1121111");

            //Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1121
            this.Site.CaptureRequirementIfAreEqual<int>(
                100,
                managedFolderInformation.StorageQuota,
                1121111,
                @"[In Appendix C: Product Behavior] Implementation does use StorageQuota which is the storage quota for a managed folder. (Exchange Server 2010 follow this behavior.)");
            #endregion

            #region Set the StoreQuota element to default value.
            this.FOLDSUTControlAdapter.DoNotSetManagedFolderStoreQuota(createManagedFolderRequest.FolderNames[0]);
            #endregion
        }
        #endregion
    }
}