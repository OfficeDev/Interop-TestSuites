namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Folder ROPs. 
    /// </summary>
    [TestClass]
    public class S02_FolderROPs : TestSuiteBase
    {
        #region Class Initialization and Cleanup

        /// <summary>
        /// Class initialize.
        /// </summary>
        /// <param name="testContext">The session context handle</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases
        /// <summary>
        /// This method tests ROP buffers of RopOpenFolder, RopCreateFolder, RopMoveFolder and RopDeleteFolder.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S02_TC01_TestOpen_Create_MoveAndDeleteFolder()
        {
            this.CheckTransportIsSupported();

            // Step 1: Send a RopOpenFolder request to the server and verify the success response.
            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            #region RopOpenFolder Success Response

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to 6th of logonResponse. This folder will be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[5];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the folder handle. This handle will be used as input handle in RopCreateFolder
            // also as a source folder handle in RopMoveFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion
            
            if (openFolderResponse.IsGhosted == 0x00)
            {
                #region Verify R561, R565 and R569

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R561");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R561
                // If ServerCount is null mean not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    openFolderResponse.ServerCount,
                    561,
                    @"[In RopOpenFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is not present if IsGhosted is zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R565");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R565
                // If CheapServerCount is null mean not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    openFolderResponse.CheapServerCount,
                    565,
                    @"[In RopOpenFolder ROP Success Response Buffer] CheapServerCount (2 bytes): This field is not present if IsGhosted is zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R569");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R569
                // If the Servers field in openFolderResponse is null, it means this field is not present.
                Site.CaptureRequirementIfIsNull(
                    openFolderResponse.Servers,
                    569,
                    @"[In RopOpenFolder ROP Success Response Buffer] Servers (variable): This field is not present if IsGhosted is zero.");

                #endregion
            }
            else
            {
                #region Verify R560, R564 and R568

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R560");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R560
                // If ServerCount is not null mean present.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.ServerCount,
                    560,
                    @"[In RopOpenFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is present if IsGhosted is nonzero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R564");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R564
                // If CheapServerCount is not null mean present.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.CheapServerCount,
                    564,
                    @"[In RopOpenFolder ROP Success Response Buffer] CheapServerCount (2 bytes): This field is present if IsGhosted is nonzero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R568");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R568
                // If the Servers field in openFolderResponse is not null, it means the field is present.
                Site.CaptureRequirementIfIsNotNull(
                    openFolderResponse.Servers,
                    568,
                    @"[In RopOpenFolder ROP Success Response Buffer] Servers (variable): This field is present if IsGhosted is nonzero.");

                #endregion
            }

            // Step 2: Create a subfolder under the opened folder and verify the success response.
            #region RopCreateFolder Success Response

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0. This field is reserved and MUST be set to 0.
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            // Send the RopCreateFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the folder id of created folder, which will be used in the following RopMoveFolder.
            ulong targetFolderId = createFolderResponse.FolderId;

            #endregion

            #region Verify R627, R631, R635, R639, R642, R643, R644, R4613 and R4614

            if (createFolderResponse.IsExistingFolder == Convert.ToByte(TestSuiteBase.Zero))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4613");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4613
                // HasRules is null means not present.
                Site.CaptureRequirementIfAreEqual<byte?>(
                    null,
                    createFolderResponse.HasRules,
                    4613,
                    @"[In RopCreateFolder ROP Success Response Buffer] HasRules (1 byte): otherwise[if the IsExistingFolder field is zero], it[HasRules (1 byte)] is not present.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R627");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R627
                // IsGhosted is null means not present.
                Site.CaptureRequirementIfAreEqual<byte?>(
                    null,
                    createFolderResponse.IsGhosted,
                    627,
                    @"[In RopCreateFolder ROP Success Response Buffer] IsGhosted (1 byte): This field is not present otherwise[if the value of the IsExistingFolder field is zero].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R631");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R631
                // ServerCount is null means not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    createFolderResponse.ServerCount,
                    631,
                    @"[In RopCreateFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is not present otherwise[if IsExistingFolder is zero].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R635");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R635
                // CheapServerCount is null means not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    createFolderResponse.CheapServerCount,
                    635,
                    @"[In RopCreateFolder ROP Success Response Buffer] CheapServerCount (2 bytes): This field is not present otherwise[if IsExistingFolder is zero].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R639");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R639
                // Servers is null means not present.
                Site.CaptureRequirementIfIsNull(
                    createFolderResponse.Servers,
                    639,
                    @"[In RopCreateFolder ROP Success Response Buffer] Servers (variable): This field is not present otherwise[if IsExistingFolder is zero].");
            }

            // This bit is set for logon to a private mailbox and is not set for logon to public folders. 

            if (createFolderResponse.IsGhosted == null || createFolderResponse.IsGhosted == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R642");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R642
                // ServerCount is null means not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    createFolderResponse.ServerCount,
                    642,
                    @"[In RopCreateFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is not present otherwise[if IsGhosted is zero].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R643");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R643
                // CheapServerCount is null means not present.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    createFolderResponse.CheapServerCount,
                    643,
                    @"[In RopCreateFolder ROP Success Response Buffer] CheapServerCount (2 bytes): This field is not present otherwise[if IsGhosted is zero].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R644");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R644
                // Servers is null means not present.
                Site.CaptureRequirementIfIsNull(
                    createFolderResponse.Servers,
                    644,
                    @"[In RopCreateFolder ROP Success Response Buffer] Servers (variable): This field is not present otherwise[if IsGhosted is zero].");
            }

            #endregion

            // Step 3: Open the second folder and set it as Destination folder (parent folder).
            #region Open Folder

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to 4th of logonResponse. This folder will be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            uint destinationParentFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];
            List<uint> handleList = new List<uint>
            {
                // The openedFolderHandle will be used as source handle in the following RopMoveFolder.
                // The destinationParentFolderHandle will be used as destination handle in the following RopMoveFolder.
                openedFolderHandle, destinationParentFolderHandle
            };
            #endregion

            // Step 4: Send a RopMoveFolder request to the server and verify the success response.
            #region RopMoveFolder Success Response

            // Move the created folder from Inbox to the second folder.
            RopMoveFolderRequest moveFolderRequest;
            RopMoveFolderResponse moveFolderResponse;

            moveFolderRequest.RopId = (byte)RopId.RopMoveFolder;
            moveFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set SourceHandleIndex to 0x00, which  specifies the location in the Server object handle table
            // where the handle for the source Server object is stored.
            moveFolderRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            // Set DestHandleIndex to 0x01, which index specifies the location in the Server object handle table
            // where the handle for the destination Server object is stored.
            moveFolderRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            moveFolderRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set UseUnicode to 0x00(FALSE), which specifies the NewFolderName field does not contain Unicode characters or multi-byte characters.
            moveFolderRequest.UseUnicode = Convert.ToByte(TestSuiteBase.Zero);

            moveFolderRequest.FolderId = targetFolderId;

            // Set NewFolderName, which specifies the name for the new moved folder.
            moveFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(Common.GenerateResourceName(this.Site, "MovedToHereByTest") + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopMoveFolder request.");

            // Send a RopMoveFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                moveFolderRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            moveFolderResponse = (RopMoveFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                moveFolderResponse.ReturnValue,
               "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 5: Send a RopDeleteFolder request to the server and verify the success response.
            #region RopDeleteFolder Response

            RopDeleteFolderRequest deleteFolderRequest;
            RopDeleteFolderResponse deleteFolderResponse;

            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            deleteFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;

            // Set FolderId to targetFolderId. This folder is to be deleted.
            deleteFolderRequest.FolderId = targetFolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopDeleteFolder request.");

            // Send a RopDeleteFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deleteFolderRequest,
                destinationParentFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            deleteFolderResponse = (RopDeleteFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 6: Send a RopOpenFolder request to the server and verify the failure response.
            #region RopOpenFolder Failure Response

            // Set FolderId to 0x1, which does not exist and will lead to a failure response.
            openFolderRequest.FolderId = TestSuiteBase.WrongFolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopOpenFolder request.");

            // Send a RopOpenFolder request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 7: Send a RopCreateFolder request to the server and verify the failure response.
            #region RopCreateFolder Failure Response

            createFolderRequest.FolderType = (byte)FolderType.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopCreateFolder request.");

            // Send a RopCreateFolder request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 8: Send a RopMoveFolder request to the server and verify the failure response about Null Destination.
            #region RopMoveFolder Null Destination Failure Response

            // Remove the destination handle, then send the RopMoveFolder request and verify the failure response.
            handleList.Remove(destinationParentFolderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopMoveFolder request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                moveFolderRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.NullDestinationFailureResponse);
            moveFolderResponse = (RopMoveFolderResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                moveFolderResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 9: Send a RopRelease request to release all resources associated with the Server object.
            #region RopRelease

            this.response = null;
            RopReleaseRequest releaseRequest;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopRelease request.");

            // Send a RopRelease request to release all resources associated with the Server object.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4581");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4581
            // Response is null means RopRelease does not return response.
            Site.CaptureRequirementIfIsNull(
                this.response,
                4581,
                @"[In Processing the RopRelease ROP Request] The server MUST not return response for a RopRelease ROP request.");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopCopyFolder, RopMoveCopyMessages,
        /// RopEmptyFolder and RopHardDeleteMessagesAndSubfolders.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S02_TC02_TestRopCopy_EmptyAndHardDeleteFolder()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Open a folder.
            #region Open Folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 6th of logonResponse. This folder is to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[5];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send a RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the folder handle. This handle will be used as input handle in the following RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a subfolder under the opened folder.
            #region Create Folder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0, which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder to be created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0. This field is reserved and MUST be set to 0.
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            // Send a RopCreateFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the folder handle, which will be used as source handle in the following RopMOveCopyMessages.
            uint createFolderHandle = responseSOHs[0][createFolderResponse.OutputHandleIndex];
            ulong targetFolderId = createFolderResponse.FolderId;

            #endregion

            // Step 3: Copy the created folder from Inbox to the second folder.
            #region RopCopyFolder Success Response

            List<uint> handleList = new List<uint>
            {
                // Add createFolderHandle into handleList, createFolderHandle will be used as source handle.
                // Add openedFolderHandle into handleList, openedFolderHandle will be used as destination handle, source and destination in the same folder.
                createFolderHandle, openedFolderHandle
            };
            RopCopyFolderRequest copyFolderRequest;
            RopCopyFolderResponse copyFolderResponse;

            // Make folder name distinct for every test run, in this way, one successful test run won't FAIL next test run.
            Random random = new Random();
            int folderNo = random.Next(MaxValueForRandom);
            string destinationTargetFolderName = "CopiedToHereByTest-" + folderNo.ToString() + "\0";

            copyFolderRequest.RopId = (byte)RopId.RopCopyFolder;
            copyFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set SourceHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the source Server object is stored.
            copyFolderRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            // Set DestHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the destination Server object is stored.
            copyFolderRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            copyFolderRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            copyFolderRequest.UseUnicode = Convert.ToByte(TestSuiteBase.Zero);

            // Set WantRecursive to 0xFF(TRUE), which specifies that the copy is recursive.
            copyFolderRequest.WantRecursive = TestSuiteBase.NonZero;

            copyFolderRequest.FolderId = targetFolderId;
            copyFolderRequest.NewFolderName = Encoding.ASCII.GetBytes(destinationTargetFolderName);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopCopyFolder request.");

            // Send the RopCopyFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyFolderRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            copyFolderResponse = (RopCopyFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 4: Create a message in the created folder.
            #region Create Message

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specifies the code page for the message.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Set FolderId to targetFolderId, which identifies the parent folder.
            createMessageRequest.FolderId = targetFolderId;

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 5: Save the created message.
            #region Save Message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            ulong targetMessageId = saveChangesMessageResponse.MessageId;

            #endregion

            // Step 6: Open the second folder, and set it as Destination folder of RopMoveCopyMessages.
            #region Open folder

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 7th folder. This folder is to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[6];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint destinationFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            handleList.Clear();

            // Add createFolderHandle into handleList. This handle will be used as source handle in the following RopMoveCopyMessages.
            handleList.Add(createFolderHandle);

            // Add createdestinationFolderHandleFolderHandle into handleList. This handle will be used as destination handle
            // in the following RopMoveCopyMessages.
            handleList.Add(destinationFolderHandle);
            ulong[] messageIds = new ulong[1];
            messageIds[0] = targetMessageId;

            #endregion

            // Step 7: Copy messages from created folder.
            #region RopMoveCopyMessages Success Response

            RopMoveCopyMessagesRequest moveCopyMessagesRequest;
            RopMoveCopyMessagesResponse moveCopyMessagesResponse;

            moveCopyMessagesRequest.RopId = (byte)RopId.RopMoveCopyMessages;
            moveCopyMessagesRequest.LogonId = TestSuiteBase.LogonId;

            // Set SourceHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the source Server object is stored.
            moveCopyMessagesRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            // Set DestHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the destination Server object is stored.
            moveCopyMessagesRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;

            // Set MessageIdCount to the length of messageIds, which specifies the size of the MessageIds field.
            moveCopyMessagesRequest.MessageIdCount = (ushort)messageIds.Length;

            // Set MessageIds to messageIds, which specify which messages to move or copy.
            moveCopyMessagesRequest.MessageIds = messageIds;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            moveCopyMessagesRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set WantCopy to 0xFF(TRUE), which specifies the operation is a copy.
            moveCopyMessagesRequest.WantCopy = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopMoveCopyMessages request.");

            // Send the RopMoveCopyMessages request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                moveCopyMessagesRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            moveCopyMessagesResponse = (RopMoveCopyMessagesResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                moveCopyMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 8: Delete all messages and subfolders from a folder through RopEmptyFolder.
            #region RopEmptyFolder Response

            RopEmptyFolderRequest emptyFolderRequest;
            RopEmptyFolderResponse emptyFolderResponse;

            emptyFolderRequest.RopId = (byte)RopId.RopEmptyFolder;
            emptyFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            emptyFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            emptyFolderRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set WantDeleteAssociated to 0xFF(TRUE), which specifies the operation also deletes folder associated information (FAI) messages.
            emptyFolderRequest.WantDeleteAssociated = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopEmptyFolder request.");

            // Send the RopEmptyFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                emptyFolderRequest,
                destinationFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            emptyFolderResponse = (RopEmptyFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                emptyFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 9: Send RopCopyFolder request to the server and verify Null Destination Failure Response.
            #region RopCopyFolder Null Destination Failure Response

            // Use createFolderHandle as source handle, but no destination handle,
            // this will lead to a failure response.
            handleList.Clear();
            handleList.Add(createFolderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopCopyFolder request.");

            // Send RopCopyFolder request to the server and verify Null Destination Failure Response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyFolderRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.NullDestinationFailureResponse);
            copyFolderResponse = (RopCopyFolderResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyFolderResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 10: Send RopMoveCopyMessages request to the server and verify Null Destination Failure Response.
            #region RopMoveCopyMessages Null Destination Failure Response

            // Use createFolderHandle as source handle, but no destination handle,
            // this will lead to a failure response.
            handleList.Clear();
            handleList.Add(createFolderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 10: Begin to send the RopMoveCopyMessages request.");

            // Send RopMoveCopyMessages request to the server and verify Null Destination Failure Response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                moveCopyMessagesRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.NullDestinationFailureResponse);
            moveCopyMessagesResponse = (RopMoveCopyMessagesResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                moveCopyMessagesResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 11: Hard delete messages and sub folder in inbox folder through RopHardDeleteMessagesAndSubfolders.
            #region RopHardDeleteMessagesAndSubfolders Response

            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest;
            RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse;

            hardDeleteMessagesAndSubfoldersRequest.RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders;
            hardDeleteMessagesAndSubfoldersRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            hardDeleteMessagesAndSubfoldersRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            hardDeleteMessagesAndSubfoldersRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set WantDeleteAssociated to 0xFF(TRUE), which specifies the operation also deletes folder associated information (FAI) messages.
            hardDeleteMessagesAndSubfoldersRequest.WantDeleteAssociated = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 11: Begin to send the RopHardDeleteMessagesAndSubfolders request.");

            // Send RopHardDeleteMessagesAndSubfolders request to the server and verify the Response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                hardDeleteMessagesAndSubfoldersRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            hardDeleteMessagesAndSubfoldersResponse = (RopHardDeleteMessagesAndSubfoldersResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 12: Release all resources associated with the Server object.
            #region RopRelease

            RopReleaseRequest releaseRequest;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 12: Begin to send the RopRelease request.");

            // Send the RopRelease request to the server.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSetSearchCriteria and RopGetSearchCriteria.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S02_TC03_TestRopGetSearchCriteria()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 5th of logonResponse. This folder is to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            uint folderHandler = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a search folder.
            #region Create folder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Searchfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName field and the Comment field
            // does not contain Unicode characters or multi-byte characters.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder to be created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0. This field is reserved and MUST be set to 0x00.
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForSearchFolder + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request.");

            // Send RopCreateFolder request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createFolderRequest,
                folderHandler,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint searchFolderHandler = responseSOHs[0][createFolderResponse.OutputHandleIndex];

            #endregion

            // Step 3: Send a RopSetSearchCriteria request to the server and verify the response.
            #region RopSetSearchCriteria Response

            RopSetSearchCriteriaRequest setSearchCriteriaRequest;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse;

            setSearchCriteriaRequest.RopId = (byte)RopId.RopSetSearchCriteria;
            setSearchCriteriaRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            setSearchCriteriaRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set RestrictionDataSize to 0x0005, which specifies the length of the RestrictionData field.
            setSearchCriteriaRequest.RestrictionDataSize = TestSuiteBase.RestrictionDataSize1;

            byte[] restrictionData = new byte[5];
            ushort pidTagMessageClassID = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyId;
            ushort typeOfPidTagMessageClass = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyType;
            restrictionData[0] = (byte)Restrictions.ExistRestriction;
            Array.Copy(BitConverter.GetBytes(typeOfPidTagMessageClass), 0, restrictionData, 1, sizeof(ushort));
            Array.Copy(BitConverter.GetBytes(pidTagMessageClassID), 0, restrictionData, 3, sizeof(ushort));
            setSearchCriteriaRequest.RestrictionData = restrictionData;

            ulong[] tempFolderIds = new ulong[1];
            tempFolderIds[0] = logonResponse.FolderIds[4];

            // Set FolderIdCount to the length of FolderIds, which specifies the number of IDs in the FolderIds field.
            setSearchCriteriaRequest.FolderIdCount = (ushort)tempFolderIds.Length;

            // Set FolderIds to that of logonResponse, which contains identifiers that specify which folders are searched. 
            setSearchCriteriaRequest.FolderIds = tempFolderIds;

            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.RestartSearch;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSetSearchCriteria request.");

            // Send RopSetSearchCriteria request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setSearchCriteriaRequest,
                searchFolderHandler,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setSearchCriteriaResponse = (RopSetSearchCriteriaResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setSearchCriteriaResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send a RopGetSearchCriteria request to the server and verify the response.
            #region RopGetSearchCriteria Success Response

            RopGetSearchCriteriaRequest getSearchCriteriaRequest;
            RopGetSearchCriteriaResponse getSearchCriteriaResponse;

            getSearchCriteriaRequest.RopId = (byte)RopId.RopGetSearchCriteria;
            getSearchCriteriaRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getSearchCriteriaRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            getSearchCriteriaRequest.UseUnicode = Convert.ToByte(TestSuiteBase.Zero);

            // Set IncludeRestriction to 0x00(FALSE), which specifies restriction data is NOT required in the response.
            getSearchCriteriaRequest.IncludeRestriction = Convert.ToByte(TestSuiteBase.Zero);

            // Set IncludeFolders to 0xFF(TRUE), which specifies the folders list is required in the response.
            getSearchCriteriaRequest.IncludeFolders = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetSearchCriteria request.");

            // Send a RopGetSearchCriteria request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getSearchCriteriaRequest,
                searchFolderHandler,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getSearchCriteriaResponse = (RopGetSearchCriteriaResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getSearchCriteriaResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 5: Send a RopGetSearchCriteria request to the server and verify the failure response.
            #region RopGetSearchCriteria Failure Response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response,
            getSearchCriteriaRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetSearchCriteria request.");

            // Send a RopGetSearchCriteria request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getSearchCriteriaRequest,
                searchFolderHandler,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getSearchCriteriaResponse = (RopGetSearchCriteriaResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getSearchCriteriaResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 6: Release all resources associated with the Server object.
            #region RopRelease

            RopReleaseRequest releaseRequest;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopRelease request.");

            // Send a RopRelease request to release all resources associated with the Server object.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopDeleteMessage and RopHardDeleteMessage.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S02_TC04_TestRopDeleteMessages()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Create a message in inbox folder.
            #region Create message

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;

            createMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specifies the code page for the message is the code page of Logon object.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Set FolderId to the 5th of logonResponse, which identifies the parent folder.
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00(FALSE), which specifies the message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Save the created message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            ulong targetMessageId = saveChangesMessageResponse.MessageId;

            #endregion

            // Step 3: Open inbox folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 5th of logonResponse. This folder is to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint inboxFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 4: Delete message in inbox folder.
            #region RopDeleteMessages Response

            ulong[] messageIds = new ulong[1];
            messageIds[0] = targetMessageId;

            RopDeleteMessagesRequest deleteMessagesRequest;
            RopDeleteMessagesResponse deleteMessagesResponse;

            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            deleteMessagesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            deleteMessagesRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set NotifyNonRead to 0x00, which specifies the server does not generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = Convert.ToByte(TestSuiteBase.Zero);

            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopDeleteMessages request.");

            // Send the RopDeleteMessages request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deleteMessagesRequest,
                inboxFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            deleteMessagesResponse = (RopDeleteMessagesResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                deleteMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 5: Create, save message and then open inbox folder.
            #region Create, save message and open inbox folder

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopCreateMessages request.");

            // Create a message in Inbox.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            targetMessageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSaveMessages request.");

            // Save the created message.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            targetMessageId = saveChangesMessageResponse.MessageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopOpenFolder request.");

            // Open Inbox.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            inboxFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 6: Send RopHardDeleteMessages request to the server and verify the success response.
            #region RopHardDeleteMessages Response

            messageIds[0] = targetMessageId;
            RopHardDeleteMessagesRequest hardDeleteMessagesRequest;

            hardDeleteMessagesRequest.RopId = (byte)RopId.RopHardDeleteMessages;
            hardDeleteMessagesRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            hardDeleteMessagesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            hardDeleteMessagesRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            // Set NotifyNonRead to 0x00(FALSE), which specifies the server does not generate a non-read receipt for the deleted messages.
            hardDeleteMessagesRequest.NotifyNonRead = Convert.ToByte(TestSuiteBase.Zero);

            hardDeleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            hardDeleteMessagesRequest.MessageIds = messageIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopHardDeleteMessages request.");

            // Send RopHardDeleteMessages request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                hardDeleteMessagesRequest,
                inboxFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 7: Release all resources associated with the Server object.
            #region RopRelease

            RopReleaseRequest releaseRequest;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopRelease request.");

            // Send RopRelease request to the server to release all resources associated with the Server object.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetHierarchyTable and RopGetContentsTable.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S02_TC05_TestGetHierarchyTableAndGetContentsTable()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 4th of logonResponse. This folder is to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Get the Hierarchy table of the opened folder through RopGetHierarchyTable.
            #region RopGetHierarchyTable Success Response

            RopGetHierarchyTableRequest getHierarchyTableRequest;
            RopGetHierarchyTableResponse getHierarchyTableResponse;

            getHierarchyTableRequest.RopId = (byte)RopId.RopGetHierarchyTable;
            getHierarchyTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getHierarchyTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            getHierarchyTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.Depth;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetHierarchyTable request.");

            // Send a RopGetHierarchyTable request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getHierarchyTableRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getHierarchyTableResponse = (RopGetHierarchyTableResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getHierarchyTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Get the Contents table of the opened folder.
            #region RopGetContentsTable Success Response

            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getContentsTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            getContentsTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetContentsTable request.");

            // Send a RopGetContentsTable request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getContentsTableRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getContentsTableResponse = (RopGetContentsTableResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getContentsTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send the RopGetHierarchyTable request to the server and verify the failure response.
            #region RopGetHierarchyTable Failure Response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            getHierarchyTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetHierarchyTable request.");

            // Send the RopGetHierarchyTable request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getHierarchyTableRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getHierarchyTableResponse = (RopGetHierarchyTableResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getHierarchyTableResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 5: Send the RopGetContentsTable request to the server and verify the failure response.
            #region RopGetContentsTable Failure Response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            getContentsTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetContentsTable request.");

            // Send the RopGetContentsTable request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getContentsTableRequest,
                targetFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getContentsTableResponse = (RopGetContentsTableResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getContentsTableResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 6: Send a RopRelease request to release all resources associated with the Server object.
            #region RopRelease

            RopReleaseRequest releaseRequest;

            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopRelease request.");

            // Send a RopRelease request to release all resources associated with the Server object.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        #endregion
    }
}