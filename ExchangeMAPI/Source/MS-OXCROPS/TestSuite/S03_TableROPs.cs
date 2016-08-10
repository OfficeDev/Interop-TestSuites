namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Table ROPs. 
    /// </summary>
    [TestClass]
    public class S03_TableROPs : TestSuiteBase
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

        //-------------------------------------------------------------
        // Table Type        | ROP to get a table handle | Specified in
        //-------------------------------------------------------------
        // Contents Table    | RopGetContentsTable       | MS-OXCFOLD
        //-------------------------------------------------------------
        // Hierarchy Table   | RopGetHierarchyTable      | MS-OXCFOLD
        //-------------------------------------------------------------
        // Attachments Table | RopGetAttachmentTable     | MS-OXCMSG
        //-------------------------------------------------------------
        // Permissions Table | RopGetPermissionsTable    | MS-OXCPERM
        //-------------------------------------------------------------
        // Rules Table       | RopGetRulesTable          | MS-OXORULE
        //-------------------------------------------------------------
        #region Test Cases

        /// <summary>
        /// This method tests the ROP buffers of RopSetColumns, RopQueryRows, RopSortTable, RopRestrict, RopGetStatusare, RopQueryPosition, RopQueryColumnsAll and RopResetTable.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC01_TestTableROPsNoDependency()
        {
            this.CheckTransportIsSupported();

            // Step 1: Open a folder and create a subfolder.
            #region Preparing the table: CreateFolder

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Open a folder first.
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 4th folder of the logonResponse, which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the handle of the opened folder, which will be used in the following RopCreateFolder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            // Create a subfolder of the opened folder, which will be used as target folder in the following ROPs.
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder being created will be opened when it is already existed.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint targetFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];
            uint tableHandle = GetContentsTableHandle(targetFolderHandle);
            ulong folderId = createFolderResponse.FolderId;

            #endregion

            // Step 2: Create and save a message.
            #region Preparing the table: RopCreateAndSaveMessages

            // Create a message.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specifies the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Set FolderId to that of the opened folder, which identifies the parent folder.
            createMessageRequest.FolderId = folderId;

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateMessage request.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            // Get the handle of created message, which will be used in the following RopSaveChangesMessage.
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            // Save message.
            RopSaveChangesMessageRequest saveChangesMessageRequest;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
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
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Send a RopSetColumns request and verify the success response.
            #region RopSetColumns success response

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            // Set propertyTags to a Sample ContentsTable PropertyTags created by CreateSampleContentsTablePropertyTags method.
            PropertyTag[] propertyTags = CreateSampleContentsTablePropertyTags();

            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSetColumns request.");

            // Send the RopSetColumns request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setColumnsRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setColumnsResponse = (RopSetColumnsResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send a RopSetColumns request and verify the failure response.
            #region RopSetColumns failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetColumns request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setColumnsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            setColumnsResponse = (RopSetColumnsResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 5: Send a RopQueryRows request and verify the success response.
            #region RopQueryRows success response

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            queryRowsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;

            // Set ForwardRead to 0x01(TRUE), which specifies the direction to read rows (forwards).
            queryRowsRequest.ForwardRead = TestSuiteBase.NonZero;

            // Set RowCount to 0x0032, which specifies the number of requested rows.
            queryRowsRequest.RowCount = TestSuiteBase.RowCount;

            List<ISerializable> ropRequests = new List<ISerializable>
            {
                setColumnsRequest, queryRowsRequest
            };

            List<uint> inputObjects = new List<uint>
            {
                tableHandle
            };

            List<IDeserializable> ropResponses = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopQueryRows request.");

            // Send a RopQueryRows request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                ropRequests,
                inputObjects,
                ref ropResponses,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            queryRowsResponse = (RopQueryRowsResponse)ropResponses[1];

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryRowsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 6: Send a RopQueryRows request and verify the failure response.
            #region RopQueryRows failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopQueryRows request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryRowsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            queryRowsResponse = (RopQueryRowsResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryRowsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 7: Send a RopSortTable request and verify the success response.
            #region RopSortTable success response

            RopSortTableRequest sortTableRequest;
            RopSortTableResponse sortTableResponse;

            // Create a Sample SortOrders by calling the CreateSampleSortOrders method.
            SortOrder[] sortOrders = this.CreateSampleSortOrders();

            sortTableRequest.RopId = (byte)RopId.RopSortTable;
            sortTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            sortTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            sortTableRequest.SortTableFlags = (byte)AsynchronousFlags.None;

            // Set SortOrderCount, which specifies how many SortOrder structures are present in the SortOrders field.
            sortTableRequest.SortOrderCount = (ushort)sortOrders.Length;

            // Set CategoryCount, which specifies the number of category SortOrder structures in the SortOrders field.
            sortTableRequest.CategoryCount = (ushort)(sortOrders.Length - 1);

            // Set ExpandedCount, which specifies the number of expanded categories in the SortOrders field.
            sortTableRequest.ExpandedCount = sortTableRequest.CategoryCount;

            sortTableRequest.SortOrders = sortOrders;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSortTable request.");

            // Send a RopSortTable request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                sortTableRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            sortTableResponse = (RopSortTableResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                sortTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 8: Send a RopSortTable request and verify the failure response.
            #region RopSortTable failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopSortTable request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                sortTableRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            sortTableResponse = (RopSortTableResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                sortTableResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 9: Send a RopRestrict request and verify the success response.
            #region RopRestrict success response

            RopRestrictRequest restrictRequest;
            RopRestrictResponse restrictResponse;

            restrictRequest.RopId = (byte)RopId.RopRestrict;
            restrictRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            restrictRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            restrictRequest.RestrictFlags = (byte)AsynchronousFlags.None;

            // Set RestrictionDataSize, which specifies the length of the RestrictionData field.
            restrictRequest.RestrictionDataSize = TestSuiteBase.RestrictionDataSize2;

            // Set RestrictionData to null, which specifies there is no filter for limiting the view of a table to particular set of rows,
            // as specified in [MS-OXCDATA].
            restrictRequest.RestrictionData = null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopRestrict request.");

            // Send a RopRestrict request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                restrictRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            restrictResponse = (RopRestrictResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                restrictResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 10: Send a RopRestrict request and verify the failure response.
            #region RopRestrict failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 10: Begin to send the RopRestrict request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                restrictRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            restrictResponse = (RopRestrictResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                restrictResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 11: Send a RopQueryPosition request and verify the success response.
            #region RopQueryPosition success response

            RopQueryPositionRequest queryPositionRequest;
            RopQueryPositionResponse queryPositionResponse;

            queryPositionRequest.RopId = (byte)RopId.RopQueryPosition;
            queryPositionRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            queryPositionRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 11: Begin to send the RopQueryPosition request.");

            // Send a RopQueryPosition request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryPositionRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            queryPositionResponse = (RopQueryPositionResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryPositionResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 12: Send a RopQueryPosition request and verify the failure response.
            #region RopQueryPosition failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 12: Begin to send the RopQueryPosition request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryPositionRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            queryPositionResponse = (RopQueryPositionResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryPositionResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 13: Send a RopQueryColumnsAll request and verify the success response.
            #region RopQueryColumnsAll success response

            RopQueryColumnsAllRequest queryColumnsAllRequest;
            RopQueryColumnsAllResponse queryColumnsAllResponse;

            queryColumnsAllRequest.RopId = (byte)RopId.RopQueryColumnsAll;
            queryColumnsAllRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            queryColumnsAllRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 13: Begin to send the RopQueryColumnsAll request.");

            // Send a RopQueryColumnsAll request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryColumnsAllRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            queryColumnsAllResponse = (RopQueryColumnsAllResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryColumnsAllResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 14: Send a RopQueryColumnsAll request and verify the failure response.
            #region RopQueryColumnsAll failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            queryColumnsAllRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 14: Begin to send the RopQueryColumnsAll request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryColumnsAllRequest,
                tableHandle,
                ref this.response,
                ref this.rawData, 
                RopResponseType.FailureResponse);
            queryColumnsAllResponse = (RopQueryColumnsAllResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue, 
                queryColumnsAllResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");
            #endregion

            // Step 15: Send a RopResetTable request and verify the success response.
            #region RopResetTable success response

            RopResetTableRequest resetTableRequest;
            RopResetTableResponse resetTableResponse;

            resetTableRequest.RopId = (byte)RopId.RopResetTable;
            resetTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            resetTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 15: Begin to send the RopResetTable request.");

            // Send a RopResetTable request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                resetTableRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            resetTableResponse = (RopResetTableResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                resetTableResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 16: Send a RopResetTable request and verify the failure response.
            #region RopResetTable failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 16: Begin to send the RopResetTable request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                resetTableRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            resetTableResponse = (RopResetTableResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                resetTableResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 17: Send a RopResetTable request and verify the success and failure response.
            #region RopGetStatus

            // Refer to MS-OXCTABL endnote<11>: Exchange 2010 and Exchange 2013 do not support asynchronous operations(RopGetStatus) 
            // on tables and ignore the TABL_ASYNC flags.
            if (Common.IsRequirementEnabled(60401, this.Site))
            {
                RopGetStatusRequest getStatusRequest;
                RopGetStatusResponse getStatusResponse;

                getStatusRequest.RopId = (byte)RopId.RopGetStatus;
                getStatusRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                getStatusRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 17: Begin to send the RopResetTable request to invoke success response.");

                // Send a RopResetTable request and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getStatusRequest,
                    tableHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getStatusResponse = (RopGetStatusResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R60401");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R60401
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getStatusResponse.ReturnValue,
                    60401,
                    @"[In Appendix A: Product Behavior] Implementation does support asynchronous operations(RopGetStatus) on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4. (Exchange 2007 follow this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getStatusResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 17: Begin to send the RopResetTable request to invoke failure response.");

                // Send a RopResetTable request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getStatusRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getStatusResponse = (RopGetStatusResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getStatusResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");
            }
            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopAbort.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC02_TestRopAbort()
        {
            this.CheckTransportIsSupported();

            // Step 1: Open a folder.
            #region Prepare table: Open a Folder

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Open a folder first
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 5th folder of the logonResponse, which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];
            #endregion

            // Step 2: Create a subfolder of the opened folder.
            #region Prepare table: Create a subfolder of the opened folder
            // Create a subfolder of the opened folder, which will be used as target folder in the following ROP.
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder being created will be opened when it is already existed,
            // as specified in [MS-OXCFOLD].
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint targetFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 3: Send a RopGetContentsTable request and verify the success response.
            #region RopGetContentsTable

            RopGetContentsTableRequest getContentsTableRequest;
            RopGetContentsTableResponse getContentsTableResponse;

            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            getContentsTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            getContentsTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetContentsTable request.");

            // Send a RopGetContentsTable request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint contentsTableHandle = responseSOHs[0][getContentsTableResponse.OutputHandleIndex];

            #endregion

            // Step 4: Send RopSetColumns and RopAbort request and verify the success responses.
            #region RopSetColumns and RopAbort

            RopSetColumnsRequest setColumnsRequest;

            // Create Sample ContentsTable PropertyTags by calling CreateSampleContentsTablePropertyTags2 method.
            PropertyTag[] propertyTags = this.CreateSampleContentsTableWith8PropertyTags();

            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.TblAsync;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            RopAbortRequest abortRequest;
            RopAbortResponse abortResponse;

            abortRequest.RopId = (byte)RopId.RopAbort;
            abortRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            abortRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopAbort request.");

            // Send a RopAbort request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                abortRequest,
                contentsTableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            abortResponse = (RopAbortResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                abortResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            // Refer to MS-OXCTABL endnote<9>: Exchange 2010 and Exchange 2013 do not support asynchronous operations(RopAbort) 
            // on tables and ignore the TABL_ASYNC flags.
            if (Common.IsRequirementEnabled(60201, this.Site))
            {
                List<ISerializable> ropRequests = new List<ISerializable>
                {
                    // Add RopSetColumns and RopAbort requests into the request buffer.
                    setColumnsRequest, abortRequest
                };
                List<uint> inputObjects = new List<uint>
                {
                    contentsTableHandle
                };
                List<IDeserializable> ropResponses = new List<IDeserializable>();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the requests, including RopSetColumns and RopAbort requests.");

                // Send the requests, including RopSetColumns and RopAbort requests.
                this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                    ropRequests,
                    inputObjects,
                    ref ropResponses,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                abortResponse = (RopAbortResponse)ropResponses[1];

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R60201");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R60201
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    abortResponse.ReturnValue,
                    60201,
                    @"[In Appendix A: Product Behavior] Implementation does support asynchronous operations(RopAbort) on tables and ignore the TABL_ASYNC flags, as described in section 2.2.2.1.4. (Exchange 2007 follow this behavior.)");
            }
            #endregion

            // Step 5: Send a RopDeleteFolder request and verify the success response.
            #region Delete the folder

            RopDeleteFolderRequest deleteFolderRequest;
            RopDeleteFolderResponse deleteFolderResponse;

            deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
            deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            deleteFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;

            // Set FolderId to that of the created folder, which identifies the folder to be deleted.
            deleteFolderRequest.FolderId = createFolderResponse.FolderId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopDeleteFolder request.");

            // Send a RopDeleteFolder request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deleteFolderRequest,
                openedFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            deleteFolderResponse = (RopDeleteFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                deleteFolderResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopCreateBookmark and RopSeekRowBookmark.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC03_TestRopCreateAndSeekRow()
        {
            this.CheckTransportIsSupported();

            // Step 1: Prepare Table object, which will be used in the following ROPs.
            #region PrepareTable

            ulong folderID;

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call CrateFolder method to open a folder and create a subfolder.");

            // Call CrateFolder method to open a folder and create a subfolder.
            uint folderHandle = this.CreateFolder(out folderID);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call GetHierarchyTableHandle method to Get HierarchyTable Handle.");

            // Call GetHierarchyTableHandle method to get hierarchy table handle.
            uint tableHandle = this.GetHierarchyTableHandle(folderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call CreateAndSaveMessage method to create and save a message in the created folder.");

            // Call CreateAndSaveMessage method to create and save a message in the created folder.
            this.CreateAndSaveMessage(folderID);

            // Send the RopSetColumns request and verify the success response.
            #region RopSetColumns

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            PropertyTag[] propertyTags = CreateSampleContentsTablePropertyTags();
            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSetColumns request.");

            // Send the RopSetColumns request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setColumnsRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setColumnsResponse = (RopSetColumnsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Step 2: Send the RopCreateBookmark request to the server and verify the success response.
            #region RopCreateBookmark success response

            // Create user-defined bookmark.
            RopCreateBookmarkRequest createBookmarkRequest;
            RopCreateBookmarkResponse createBookmarkResponse;

            createBookmarkRequest.RopId = (byte)RopId.RopCreateBookmark;
            createBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateBookmark request.");

            // Send the RopCreateBookmark request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createBookmarkResponse = (RopCreateBookmarkResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createBookmarkResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            byte[] userDefinedBookmark = createBookmarkResponse.Bookmark;

            #endregion

            // Step 3: Send the RopSeekRowBookmark request to the server and verify the success response.
            #region RopSeekRowBookmark success response

            RopSeekRowBookmarkRequest seekRowBookmarkRequest;

            seekRowBookmarkRequest.RopId = (byte)RopId.RopSeekRowBookmark;
            seekRowBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            seekRowBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set BookmarkSize, which specifies the size of the Bookmark field.
            seekRowBookmarkRequest.BookmarkSize = (ushort)userDefinedBookmark.Length;

            // Set Bookmark, which specifies the origin for the seek operation.
            seekRowBookmarkRequest.Bookmark = userDefinedBookmark;

            // Set RowCount, which specifies the direction and the number of rows to seek.
            seekRowBookmarkRequest.RowCount = TestSuiteBase.RowCount;

            // Set WantRowMovedCount to 0xff(TRUE), which specifies the server returns the actual number of rows sought
            // in the response.
            seekRowBookmarkRequest.WantRowMovedCount = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSeekRowBookmark request.");

            // Send the RopSeekRowBookmark request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekRowBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createBookmarkResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send the RopCreateBookmark request to the server and verify the failure response.
            #region RopCreateBookmark failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            createBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopCreateBookmark request.");

            // Send the RopCreateBookmark request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            createBookmarkResponse = (RopCreateBookmarkResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createBookmarkResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 5: Send the RopSeekRowBookmark request to the server and verify the failure response.
            #region RopSeekRowBookmark failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            seekRowBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSeekRowBookmark request.");

            // Send the RopSeekRowBookmark request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekRowBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createBookmarkResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSeekRowFractional and RopSeekRow.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC04_TestRopSeekRowAndFractional()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopSeekRowFractional request to the server and verify the success response.
            #region RopSeekRowFractional success response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call CreateFolder method to open a folder and create a subfolder.");

            // Call CreateFolder method to open a folder and create a subfolder.
            uint folderHandle = this.CreateFolder();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call GetContentsTableHandle method to Get ContentsTable Handle.");

            // Call GetContentsTableHandle method to get contents table handle.
            uint tableHandle = GetContentsTableHandle(folderHandle);

            RopSeekRowFractionalRequest seekRowFractionalRequest;
            RopSeekRowFractionalResponse seekRowFractionalResponse;

            seekRowFractionalRequest.RopId = (byte)RopId.RopSeekRowFractional;
            
            seekRowFractionalRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            seekRowFractionalRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set Numerator, which represents the numerator of the fraction identifying the table position to seek to.
            seekRowFractionalRequest.Numerator = TestSuiteBase.Numerator;

            // Set Denominator, which represents the denominator of the fraction identifying the table position to seek to.
            seekRowFractionalRequest.Denominator = TestSuiteBase.Denominator;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSeekRowFractional request.");

            // Send the RopSeekRowFractional request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekRowFractionalRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            seekRowFractionalResponse = (RopSeekRowFractionalResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                seekRowFractionalResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 2: Send the RopSeekRow request to the server and verify the success response.
            #region RopSeekRow success response

            RopSeekRowRequest seekRowRequest;
            RopSeekRowResponse seekRowResponse;

            seekRowRequest.RopId = (byte)RopId.RopSeekRow;
            seekRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            seekRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            seekRowRequest.Origin = (byte)Origin.Beginning;

            // Set RowCount, which specifies the direction and the number of rows to seek.
            seekRowRequest.RowCount = TestSuiteBase.RowCount;

            // Set WantRowMovedCount to 0xff(TRUE), which specifies the server returns the actual number of rows moved
            // in the response.
            seekRowRequest.WantRowMovedCount = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSeekRow request.");

            // Send the RopSeekRow request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            seekRowResponse = (RopSeekRowResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                seekRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Send the RopSeekRow request to the server and verify the failure response.
            #region RopSeekRow failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            seekRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSeekRow request.");

            // Send the RopSeekRow request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            seekRowResponse = (RopSeekRowResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                seekRowResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopCreateBookmark and RopFreeBookmark.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC05_TestRopFreeBookmark()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Prepare Table object, which will be used in the following ROPs.
            #region PrepareTable

            ulong folderID;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call CreateFolder method to open a folder and create a subfolder.");

            // Call CreateFolder method to open a folder and create a subfolder.
            uint folderHandle = this.CreateFolder(out folderID);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to call GetHierarchyTableHandle method to Get HierarchyTable Handle.");

            // Call GetHierarchyTableHandle method to get hierarchy table handle.
            uint tableHandle = this.GetHierarchyTableHandle(folderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to call CreateAndSaveMessage method to create and save a message in the created folder.");

            // Call CreateAndSaveMessage method to create and save a message in the created folder.
            this.CreateAndSaveMessage(folderID);

            #region RopSetColumns

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            PropertyTag[] propertyTags = CreateSampleContentsTablePropertyTags();
            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetColumns request.");

            // Send the RopSetColumns request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setColumnsRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setColumnsResponse = (RopSetColumnsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Step 2: Send the RopCreateBookmark request to the server and verify the success response.
            #region RopCreateBookmark success response

            // Create user-defined bookmark. 
            RopCreateBookmarkRequest createBookmarkRequest;
            RopCreateBookmarkResponse createBookmarkResponse;

            // Set RopId to 0x1B, which specifies the type of ROP is RopCreateBookmark.
            createBookmarkRequest.RopId = (byte)RopId.RopCreateBookmark;

            createBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopCreateBookmark request.");

            // Send the RopCreateBookmark request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createBookmarkResponse = (RopCreateBookmarkResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createBookmarkResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            byte[] userDefinedBookmark = createBookmarkResponse.Bookmark;

            #endregion

            // Step 3: Send the RopFreeBookmark request to the server and verify the success response.
            #region RopFreeBookmark

            RopFreeBookmarkRequest freeBookmarkRequest;
            RopFreeBookmarkResponse freeBookmarkResponse;

            freeBookmarkRequest.RopId = (byte)RopId.RopFreeBookmark;
            freeBookmarkRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            freeBookmarkRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            freeBookmarkRequest.BookmarkSize = (ushort)userDefinedBookmark.Length;
            freeBookmarkRequest.Bookmark = userDefinedBookmark;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopFreeBookmark request.");

            // Send the RopFreeBookmark request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                freeBookmarkRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            freeBookmarkResponse = (RopFreeBookmarkResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                freeBookmarkResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSetColumns, RopFindRows, RopSortTable, RopQueryRows,
        /// RopCollapseRow, RopExpandRows,RopGetCollapseState and RopSetCollapseState.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S03_TC06_TestCollapseAndExpandRow()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Prepare Table object, which will be used in the following ROPs.
            #region PrepareTable

            ulong folderID;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call CreateFolder method to open a folder and create a subfolder.");

            // Call CreateFolder method to open a folder and create a subfolder.
            uint folderHandle = this.CreateFolder(out folderID);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to call GetContentsTableHandle method to Get ContentsTable Handle.");

            // Call GetContentsTableHandle method to get contents table handle.
            uint tableHandle = GetContentsTableHandle(folderHandle);

            // Preparing the table: RopCreateAndSaveMessages
            for (int i = 0; i < 5; i++)
            {
                this.CreateAndSaveMessage(folderID);
            }

            #region RopSetColumns

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            PropertyTag[] propertyTags = CreateSampleContentsTablePropertyTags();
            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSetColumns request.");

            // Send the RopSetColumns request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setColumnsRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setColumnsResponse = (RopSetColumnsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setColumnsResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Step 2: Send the RopFindRows request to the server and verify the success response.
            #region RopFindRows success response with the field Origin value is Beginning

            RopFindRowRequest findRowRequest;
            RopFindRowResponse findRowResponse;

            findRowRequest.RopId = (byte)RopId.RopFindRow;
            findRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            findRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            findRowRequest.FindRowFlags = (byte)FindRowFlags.Forwards;

            // Set RestrictionDataSize, which specifies the length of the RestrictionData field.
            findRowRequest.RestrictionDataSize = TestSuiteBase.RestrictionDataSize1;

            byte[] restrictionData = new byte[5];
            ushort pidTagMessageClassID = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyId;
            ushort typeOfPidTagMessageClass = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyType;
            restrictionData[0] = (byte)Restrictions.ExistRestriction;
            Array.Copy(BitConverter.GetBytes(typeOfPidTagMessageClass), 0, restrictionData, 1, sizeof(ushort));
            Array.Copy(BitConverter.GetBytes(pidTagMessageClassID), 0, restrictionData, 3, sizeof(ushort));
            findRowRequest.RestrictionData = restrictionData;
            findRowRequest.Origin = (byte)Origin.Beginning;

            // Set BookmarkSize, which specifies the size of the Bookmark field.
            findRowRequest.BookmarkSize = TestSuiteBase.BookmarkSize;

            // Set Bookmark, which specifies the bookmark to use as the origin.
            findRowRequest.Bookmark = null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopFindRows request.");

            // Send the RopFindRows request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                findRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            findRowResponse = (RopFindRowResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                findRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #region RopFindRows success response with the field Origin value is Current

            findRowRequest.Origin = (byte)Origin.Current;

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                findRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            findRowResponse = (RopFindRowResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                findRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #region RopFindRows success response with the field Origin value is End

            findRowRequest.FindRowFlags = (byte)FindRowFlags.Backwards;
            findRowRequest.Origin = (byte)Origin.End;

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                findRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            findRowResponse = (RopFindRowResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                findRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            if (findRowResponse.HasRowData != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1484");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1484
                // RowData is not null means present.
                Site.CaptureRequirementIfIsNotNull(
                    findRowResponse.RowData,
                    1484,
                    @"[In RopFindRow ROP Success Response Buffer] RowData (variable): This field is present only when the HasRowData field is set to a nonzero value.");
            }

            #endregion

            // Step 3: Send the RopFindRows request to the server and verify the failure response.
            #region RopFindRows failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            findRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopFindRows request.");

            // Send the RopFindRows request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                findRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            findRowResponse = (RopFindRowResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                findRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send the RopSortTable request to the server and verify the success response.
            #region RopSortTable

            RopSortTableRequest sortTableRequest;

            // Call CreateSampleSortOrders method to Create Sample SortOrders.
            SortOrder[] sortOrders = this.CreateSampleSortOrders();

            sortTableRequest.RopId = (byte)RopId.RopSortTable;
            sortTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            sortTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            sortTableRequest.SortTableFlags = (byte)AsynchronousFlags.None;
            sortTableRequest.SortOrderCount = (ushort)sortOrders.Length;
            sortTableRequest.CategoryCount = (ushort)(sortOrders.Length - 1);
            sortTableRequest.ExpandedCount = sortTableRequest.CategoryCount;
            sortTableRequest.SortOrders = sortOrders;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSortTable request.");

            // Send the RopSortTable request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                sortTableRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopSortTableResponse sortTableResponse = (RopSortTableResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
            sortTableResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0x00000000(success)");

            #endregion

            // Step 5: Send the RopQueryRows request to the server and verify the success response.
            #region RopQueryRows: send RopSetColumns and RopQueryRows in a request buffer

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            queryRowsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;

            // Set ForwardRead to 0xff(TRUE), which specifies the direction to read rows (forwards).
            queryRowsRequest.ForwardRead = TestSuiteBase.NonZero;

            // Set RowCount to 0x0032, which the number of requested rows.
            queryRowsRequest.RowCount = TestSuiteBase.RowCount;

            List<ISerializable> ropRequests = new List<ISerializable>
            {
                setColumnsRequest, queryRowsRequest
            };
            List<uint> inputObjects = new List<uint>
            {
                tableHandle
            };
            List<IDeserializable> ropResponses = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the requests, including RopQueryRows and RopSetColumns requests.");

            // Send RopQueryRows and RopSetColumns requests and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                ropRequests,
                inputObjects,
                ref ropResponses,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            queryRowsResponse = (RopQueryRowsResponse)ropResponses[1];

            #endregion

            #region GetCategoryId

            byte[] pidTagInstValue = queryRowsResponse.RowData.PropertyRows[0].PropertyValues[2].Value;
            ulong index = 1;
            ulong categoryId = 0;

            foreach (byte a in pidTagInstValue)
            {
                categoryId += (ulong)a * index;
                index *= 0x100;
            }

            #endregion

            // Step 6: Send the RopCollapseRow request to the server and verify the success response.
            #region RopCollapseRow success response

            RopCollapseRowRequest collapseRowRequest;
            RopCollapseRowResponse collapseRowResponse;

            collapseRowRequest.RopId = (byte)RopId.RopCollapseRow;
            collapseRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            collapseRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set CategoryId, which specifies the category to be collapsed.
            collapseRowRequest.CategoryId = categoryId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopCollapseRow request.");

            // Send the RopCollapseRow request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                collapseRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            collapseRowResponse = (RopCollapseRowResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
            collapseRowResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0x00000000(success)");

            #endregion

            // Step 7: Send the RopCollapseRow request to the server and verify the failure response.
            #region RopCollapseRow failure response

            // Set CategoryId to 0x00, which cannot be found and will lead to a failure response.
            collapseRowRequest.CategoryId = TestSuiteBase.WrongCategoryId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopCollapseRow request.");

            // Send the RopCollapseRow request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                collapseRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            collapseRowResponse = (RopCollapseRowResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                collapseRowResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 8: Send the RopExpandRows request to the server and verify the success response.
            #region RopExpandRows success response: send RopSetColumns and RopExpandRows in a request buffer

            RopExpandRowRequest expandRowRequest;
            RopExpandRowResponse expandRowResponse;

            expandRowRequest.RopId = (byte)RopId.RopExpandRow;
            expandRowRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            expandRowRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MaxRowCount, which specifies the maximum number of expanded rows to return data for.
            expandRowRequest.MaxRowCount = TestSuiteBase.MaxRowCount;

            expandRowRequest.CategoryId = categoryId;
            ropRequests = new List<ISerializable>
            {
                setColumnsRequest, expandRowRequest
            };

            inputObjects = new List<uint>
            {
                tableHandle
            };
            ropResponses = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the requests, including RopSetColumns and RopExpandRows requests.");

            // Refer to MS-OXCTABL endnote<15>: Exchange 2013 does not support a value greater than 0 for the MaxRowCount field.
            if (Common.IsRequirementEnabled(74801, this.Site))
            {
                // Send the RopSetColumns and RopExpandRows request and verify the success response of RopExpandRows.
                this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                    ropRequests,
                    inputObjects,
                    ref ropResponses,
                    ref this.rawData,
                RopResponseType.SuccessResponse);
                expandRowResponse = (RopExpandRowResponse)ropResponses[1];

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R74801");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R74801
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    expandRowResponse.ReturnValue,
                    74801,
                    @"[In Appendix A: Product Behavior] For RopExpandRow, implementation does support a value greater than 0 for the MaxRowCount field. (Exchange 2007 and Exchange 2010 follow this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    expandRowResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");
            }

            #endregion

            // Step 9: Send the RopExpandRows request to the server and verify the failure response.
            #region RopExpandRows failure response

            // Set CategoryId to 0x00, which cannot be found and will lead to a failure response.
            expandRowRequest.CategoryId = TestSuiteBase.WrongCategoryId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopExpandRows request.");

            // Send the RopExpandRows request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                expandRowRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            expandRowResponse = (RopExpandRowResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                expandRowResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 10: Send the RopGetCollapseState request to the server and verify the failure response.
            #region RopGetCollapseState failure response

            RopGetCollapseStateRequest getCollapseStateRequest;
            RopGetCollapseStateResponse getCollapseStateResponse;

            getCollapseStateRequest.RopId = (byte)RopId.RopGetCollapseState;
            getCollapseStateRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            getCollapseStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set RowId, which specifies the row to be preserved as the cursor.
            getCollapseStateRequest.RowId = categoryId;

            // Set RowInstanceNumber, which specifies the instance number of the row that is to be preserved as the cursor.
            getCollapseStateRequest.RowInstanceNumber = TestSuiteBase.RowInstanceNumber;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 10: Begin to send the RopGetCollapseState request.");

            // Send the RopGetCollapseState request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getCollapseStateRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getCollapseStateResponse = (RopGetCollapseStateResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getCollapseStateResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 11: Send the RopGetCollapseState request to the server and verify the success response.
            #region RopGetCollapseState success response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 11: Begin to send the RopGetCollapseState request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getCollapseStateRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getCollapseStateResponse = (RopGetCollapseStateResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getCollapseStateResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0x00000000(success)");

            #endregion

            // Step 12: Send the RopSetCollapseState request to the server and verify the success response.
            #region RopSetCollapseState success response

            RopSetCollapseStateRequest setCollapseStateRequest;
            RopSetCollapseStateResponse setCollapseStateResponse;

            setCollapseStateRequest.RopId = (byte)RopId.RopSetCollapseState;
            setCollapseStateRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            setCollapseStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set CollapseStateSize, which specifies the size of the CollapseState field.
            setCollapseStateRequest.CollapseStateSize = getCollapseStateResponse.CollapseStateSize;

            // Set CollapseState, which specifies a collapse state for a categorized table.
            setCollapseStateRequest.CollapseState = (byte[])getCollapseStateResponse.CollapseState;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 12: Begin to send the RopSetCollapseState request.");

            // Send the RopSetCollapseState request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setCollapseStateRequest,
                tableHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setCollapseStateResponse = (RopSetCollapseStateResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setCollapseStateResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 13: Send the RopSetCollapseState request to the server and verify the failure response.
            #region RopSetCollapseState failure response

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 13: Begin to send the RopSetCollapseState request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setCollapseStateRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            setCollapseStateResponse = (RopSetCollapseStateResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setCollapseStateResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        #endregion

        #region Common methods

        /// <summary>
        /// Create Sample SortOrders
        /// </summary>
        /// <returns>Return SortOrder array</returns>
        private SortOrder[] CreateSampleSortOrders()
        {
            SortOrder[] sortOrders = new SortOrder[2];
            SortOrder sortOrder;

            // The tags are from CreateSampleContentsTablePropertyTags(),
            // which references MS-OXCTABL 4.2

            // PidTagMessageDeliveryTime
            sortOrder.PropertyId = this.propertyDictionary[PropertyNames.PidTagMessageDeliveryTime].PropertyId;
            sortOrder.PropertyType = this.propertyDictionary[PropertyNames.PidTagMessageDeliveryTime].PropertyType;
            sortOrder.Order = (byte)Order.Ascending;
            sortOrders[0] = sortOrder;

            // PidTagInstID
            sortOrder.PropertyId = this.propertyDictionary[PropertyNames.PidTagInstID].PropertyId;
            sortOrder.PropertyType = this.propertyDictionary[PropertyNames.PidTagInstID].PropertyType;
            sortOrder.Order = (byte)Order.Descending;
            sortOrders[1] = sortOrder;

            return sortOrders;
        }

        /// <summary>
        /// Create Folder
        /// </summary>
        /// <returns>Return the created folder object handle</returns>
        private uint CreateFolder()
        {
            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Open a folder.
            #region Open a folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 4th folder of the logonResponse, which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request in CreateFolder method.");

            // Send the RopOpenFolder request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a subfolder of the opened folder.
            #region Create a subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0. This field is reserved and MUST be set to 0.
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request in CreateFolder method.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            uint targetFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];
            return targetFolderHandle;
        }

        /// <summary>
        /// Create a folder.
        /// </summary>
        /// <param name="folderID">The folder id of the created folder.</param>
        /// <returns>Return the created folder object handle.</returns>
        private uint CreateFolder(out ulong folderID)
        {
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Step 1: Open a folder
            #region Open a folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            
            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 4th folder of the logonResponse, which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request in CreateFolder method.");

            // Send the RopOpenFolder request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Create a subfolder of the opened folder.
            #region Create a subfolder

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

            // Set UseUnicodeStrings to 0x0(FALSE), which specifies the DisplayName and Comment are not specified in Unicode.
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

            // Set OpenExisting to 0xFF(TRUE), which means the folder being created will be opened when it is already existed,
            // as specified in [MS-OXCFOLD].
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0. This field is reserved and MUST be set to 0.
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateFolder request in CreateFolder method.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            folderID = createFolderResponse.FolderId;
            return targetFolderHandle;
        }

        /// <summary>
        /// Create and save a message.
        /// </summary>
        /// <param name="folderID">The folder id of parent folder.</param>
        private void CreateAndSaveMessage(ulong folderID)
        {
            // Step 1: Create a message.
            #region Create a message

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;

            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            
            createMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            createMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Set FolderId to that of the opened folder, which identifies the parent folder.
            createMessageRequest.FolderId = folderID;

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request in CreateAndSaveMessage method.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Save the created message.
            #region Save the created message

            RopSaveChangesMessageRequest saveChangesMessageRequest;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSaveChangesMessage request in CreateAndSaveMessage method.");

            // Send the RopSaveChangesMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion

            // Step 3: Release the server object.
            #region Release the server object

            RopReleaseRequest releaseRequest;
            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopRelease request in CreateAndSaveMessage method.");

            // Send the RopRelease request to release the server object.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                releaseRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion
        }

        /// <summary>
        /// Get HierarchyTable Handle
        /// </summary>
        /// <param name="targetFolderHandle">The target folder object handle</param>
        /// <returns>Return HierarchyTable handle</returns>
        private uint GetHierarchyTableHandle(uint targetFolderHandle)
        {
            RopGetHierarchyTableRequest getHierarchyTableRequest;
            RopGetHierarchyTableResponse getHierarchyTableResponse;

            getHierarchyTableRequest.RopId = (byte)RopId.RopGetHierarchyTable;

            getHierarchyTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            getHierarchyTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            getHierarchyTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getHierarchyTableRequest.TableFlags = (byte)FolderTableFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetHierarchyTable request in GetHierarchyTableHandle method.");

            // Send the RopGetHierarchyTable request and verify the success response.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            uint hierarchyTableHandle = responseSOHs[0][getHierarchyTableResponse.OutputHandleIndex];
            return hierarchyTableHandle;
        }

        /// <summary>
        /// Create Sample ContentsTable PropertyTags
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        private PropertyTag[] CreateSampleContentsTableWith8PropertyTags()
        {
            // The following sample tags is from [MS-OXCTABL].
            PropertyTag[] propertyTags = new PropertyTag[8];

            // PidTagFolderId
            propertyTags[0] = this.propertyDictionary[PropertyNames.PidTagFolderId];

            // PidTagMid
            propertyTags[1] = this.propertyDictionary[PropertyNames.PidTagMid];

            // PidTagInstID
            propertyTags[2] = this.propertyDictionary[PropertyNames.PidTagInstID];

            // PidTagInstanceNum
            propertyTags[3] = this.propertyDictionary[PropertyNames.PidTagInstanceNum];

            // PidTagSubject
            propertyTags[4] = this.propertyDictionary[PropertyNames.PidTagSubject];

            // PidTagMessageDeliveryTime
            propertyTags[5] = this.propertyDictionary[PropertyNames.PidTagMessageDeliveryTime];

            // PidTagRowType
            propertyTags[6] = this.propertyDictionary[PropertyNames.PidTagRowType];

            // PidTagContentCount
            propertyTags[7] = this.propertyDictionary[PropertyNames.PidTagContentCount];

            return propertyTags;
        }

        #endregion
    }
}