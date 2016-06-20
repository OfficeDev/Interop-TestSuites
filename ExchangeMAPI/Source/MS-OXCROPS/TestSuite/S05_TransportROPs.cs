namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Transport ROPs. 
    /// </summary>
    [TestClass]
    public class S05_TransportROPs : TestSuiteBase
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
        /// This method tests ROP buffers of RopSubmitMessage.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC01_TestRopSubmitMessage()
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

            // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
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
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
            uint messageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Step 2: Configure recipients.
            #region Configure recipients

            RopModifyRecipientsRequest modifyRecipientsRequest;
            RopModifyRecipientsResponse modifyRecipientsResponse;

            modifyRecipientsRequest.RopId = (byte)RopId.RopModifyRecipients;
            modifyRecipientsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            modifyRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set recipientColumns to null, which means no property values are included for each recipient.
            PropertyTag[] recipientColumns = null;

            // Call CreateSampleRecipientColumnsAndRecipientRows method to create Recipient Rows and Recipient Columns.
            ModifyRecipientRow[] recipientRows = null;
            this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);

            // Set ColumnCount, which specifies the number of rows in the RecipientRows field.
            modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;

            // Set RecipientColumns to that created above, which specifies the property values that can be included
            // for each recipient.
            modifyRecipientsRequest.RecipientColumns = recipientColumns;

            // Set RowCount, which specifies the number of rows in the RecipientRows field.
            modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;

            // Set RecipientRows to that created above, which is a list of ModifyRecipientRow structures.
            modifyRecipientsRequest.RecipientRows = recipientRows;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopModifyRecipients request.");

            // Send the RopModifyRecipients request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                modifyRecipientsRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                modifyRecipientsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3: Send the RopSubmitMessage request and verify the success response.
            #region RopSubmitMessage response with the field SubmitFlags value is None

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

            // Set FolderId to the 5th of logonResponse, this folder is to be opened.
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

            #region Send RopGetContentsTable request

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

            // The mail count before send RopSubmitMessage request.
            uint mailCount = getContentsTableResponse.RowCount;

            #endregion

            #region Send RopSubmitMessage request
            RopSubmitMessageRequest submitMessageRequest;

            submitMessageRequest.RopId = (byte)RopId.RopSubmitMessage;
            submitMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            submitMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            submitMessageRequest.SubmitFlags = (byte)SubmitFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSubmitMessage request with the field SubmitFlags value is None.");

            // Send the RopSubmitMessage request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                submitMessageRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.Response);
            #endregion

            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            while (retryCount >= 0)
            {
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

                retryCount--;
                System.Threading.Thread.Sleep(waitTime);

                if (getContentsTableResponse.RowCount == mailCount + 1)
                {
                    break;
                }

                if (retryCount < 0)
                {
                    Site.Assert.Fail("The message fails to be submitted to server by sending RopSubmitMessage request.");
                }
            }

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopAbortSubmit.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC02_TestRopAbortSubmit()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create, modify recipients and submit message.

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);
            uint logonHandle = responseSOHs[0][logonResponse.OutputHandleIndex];
            int runTimes = 0;
            int runTimesLimit = 10;
            bool abortSubmitFailed = false;
            do
            {
                #region Create, modify recipients and submit message

                // Create a message.
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

                // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
                createMessageRequest.FolderId = logonResponse.FolderIds[4];

                // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
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
                    "If ROP succeeds, the ReturnValue of its response is 0(success).");
                uint messageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

                #endregion

                // Configure recipients.
                #region Configure recipients

                RopModifyRecipientsRequest modifyRecipientsRequest;
                RopModifyRecipientsResponse modifyRecipientsResponse;

                modifyRecipientsRequest.RopId = (byte)RopId.RopModifyRecipients;
                modifyRecipientsRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                modifyRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Call CreateSampleRecipientColumnsAndRecipientRows method to create Recipient Rows and Recipient Columns.
                PropertyTag[] recipientColumns = null;
                ModifyRecipientRow[] recipientRows = null;
                this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);

                // Set ColumnCount, which specifies the number of rows in the RecipientRows field.
                modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;

                // Set RecipientColumns to that created above, which specifies the property values that can be included
                // for each recipient.
                modifyRecipientsRequest.RecipientColumns = recipientColumns;

                // Set RowCount, which specifies the number of rows in the RecipientRows field.
                modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;

                // Set RecipientRows to that created above, which is a list of ModifyRecipientRow structures.
                modifyRecipientsRequest.RecipientRows = recipientRows;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopModifyRecipients request.");

                // Send the RopModifyRecipients request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    modifyRecipientsRequest,
                    messageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                modifyRecipientsResponse = (RopModifyRecipientsResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    modifyRecipientsResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Save the created message.
                #region Save message

                RopSaveChangesMessageRequest saveChangesMessageRequest;
                RopSaveChangesMessageResponse saveChangesMessageResponse;

                saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
                saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
                // in the response.
                saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

                saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSaveChangesMessage request.");

                // Send the RopSaveChangesMessage request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    saveChangesMessageRequest,
                    messageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    saveChangesMessageResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");
                ulong messageId = saveChangesMessageResponse.MessageId;

                #endregion

                // Submit message.
                #region Submit message

                RopSubmitMessageRequest submitMessageRequest;
                RopSubmitMessageResponse submitMessageResponse;

                submitMessageRequest.RopId = (byte)RopId.RopSubmitMessage;
                submitMessageRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                submitMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                submitMessageRequest.SubmitFlags = (byte)SubmitFlags.None;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSubmitMessage request.");

                // Send the RopSubmitMessage request to the server and verify the success response.
                cropsAdapter.ProcessSingleRop(
                    submitMessageRequest,
                    messageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                submitMessageResponse = (RopSubmitMessageResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    submitMessageResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0(success).");

                #endregion
                
                #endregion

                // Step 2: Send the RopAbortSubmit.
                #region RopAbortSubmit response

                RopAbortSubmitRequest abortSubmitRequest;
                RopAbortSubmitResponse abortSubmitResponse;

                abortSubmitRequest.RopId = (byte)RopId.RopAbortSubmit;
                abortSubmitRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                abortSubmitRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
                abortSubmitRequest.FolderId = logonResponse.FolderIds[4];

                // Set MessageId to that of created message, which identifies the submitted message.
                abortSubmitRequest.MessageId = messageId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopAbortSubmit request.");

                // Send the RopAbortSubmit request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    abortSubmitRequest,
                    logonHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                abortSubmitResponse = (RopAbortSubmitResponse)response;
                Site.Assert.IsTrue(abortSubmitResponse.ReturnValue == 0 || abortSubmitResponse.ReturnValue == 0x80040601, "abort submit response fail, the error code is {0}", abortSubmitResponse.ReturnValue);
                
                // 0x80040601: The message is no longer in the spooler queue of the Message store. Specified in MS-OXOMSG 3.3.5.2
                abortSubmitFailed = abortSubmitResponse.ReturnValue == 0x80040601;

                #endregion
                runTimes++;
            }
            while (runTimes < runTimesLimit && abortSubmitFailed);

            // If abort submit failed, it means there are messages left in inbox which need to be deleted
            if (runTimes > 1)
            {
                this.DeleteSubFolderAndMessage(logonResponse.FolderIds[4]);

                if (runTimes == runTimesLimit)
                {
                    Site.Assert.Fail("Retry to send RopAbortSubmit {0} times, but fail to get successful RopAbortSubmit response since it always returns an error code 0x80040601 (The message is no longer in the spooler queue of the Message store. Specified in MS-OXOMSG 3.3.5.2).", runTimesLimit);
                }
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetAddressTypes.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC03_TestRopGetAddressTypes()
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
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            // Step 1: Send the RopGetAddressTypes request and verify the success response.
            #region RopGetAddressTypes success response

            RopGetAddressTypesRequest getAddressTypesRequest;
            RopGetAddressTypesResponse getAddressTypesResponse;

            getAddressTypesRequest.RopId = (byte)RopId.RopGetAddressTypes;
            
            getAddressTypesRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            getAddressTypesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetAddressTypes request.");

            // Send the RopGetAddressTypes request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getAddressTypesRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getAddressTypesResponse = (RopGetAddressTypesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getAddressTypesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 2: Send the RopGetAddressTypes request and verify the failure response.
            #region RopGetAddressTypes failure response

            // Refer to MS-OXCROPS endnote<14>: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the Server object 
            // and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                getAddressTypesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetAddressTypes request to invoke failure response.");

                // Send the RopGetAddressTypes request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getAddressTypesRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getAddressTypesResponse = (RopGetAddressTypesResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getAddressTypesResponse.ReturnValue,
                    "Exchange 2007 does not fail the ROP if the index is invalid. The ReturnValue of ROP RopGetAddressTypes is {0}.",
                    getAddressTypesResponse.ReturnValue);
            }
            else
            {
                // For the server other than Exchange 2007, the ReturnValue is not equal to 0.
                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                getAddressTypesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetAddressTypes request to invoke failure response.");

                // Send the RopGetAddressTypes request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getAddressTypesRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getAddressTypesResponse = (RopGetAddressTypesResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getAddressTypesResponse.ReturnValue,
                    "For this response, this field is set to a value other than 0x00000000.");
            }

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopOptionsData.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC04_TestRopOptionsData()
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
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            // Step 1: Send a RopOptionsData request and verify the success response.
            #region RopOptionsData success response

            RopOptionsDataRequest optionsDataRequest;
            RopOptionsDataResponse optionsDataResponse;

            optionsDataRequest.RopId = (byte)RopId.RopOptionsData;
            
            optionsDataRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            optionsDataRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set AddressType, which specifies the address type for which options are to be returned.
            optionsDataRequest.AddressType = Encoding.ASCII.GetBytes(TestSuiteBase.AddressType);

            // Set WantWin32 to 0xff(TRUE), which specifies the help file data is to be returned in a format
            // that is suited for 32-bit machines.
            optionsDataRequest.WantWin32 = TestSuiteBase.NonZero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOptionsData request.");

            // Send the RopOptionsData request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                optionsDataRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            optionsDataResponse = (RopOptionsDataResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                optionsDataResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            if (optionsDataResponse.HelpFileSize == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2589");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2589
                // HelpFileName is null means not present.
                Site.CaptureRequirementIfIsNull(
                    optionsDataResponse.HelpFileName,
                    2589,
                    @"[In RopOptionsData ROP Success Response Buffer,HelpFileName (variable)]is not present otherwise[HelpFileSize is zero].");
            }

            #endregion

            // Step 2: Send a RopOptionsData request and verify the failure response.
            #region RopOptionsData failure response

            // Refer to MS-OXCROPS: Exchange 2007 sets the ReturnValue field for the RopOptionsData 
            // ROP response to 0x00000000 regardless of the failure of the ROP.
            if (Common.IsRequirementEnabled(4690, this.Site))
            {
                // Exchange 2007 sets the ReturnValue field to 0x00000000 regardless of the failure of the ROP.
                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                optionsDataRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOptionsData request.");

                // Send the RopOptionsData request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    optionsDataRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                optionsDataResponse = (RopOptionsDataResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    optionsDataResponse.ReturnValue,
                    "<5> Section 2.2.7.9.3: Exchange 2007 sets the ReturnValue field to 0x00000000 regardless of the failure of the ROP.");
            }
            else
            {
                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                optionsDataRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOptionsData request.");

                // Send the RopOptionsData request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    optionsDataRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                RopResponseType.FailureResponse);
                optionsDataResponse = (RopOptionsDataResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    optionsDataResponse.ReturnValue,
                    "For this response, this field SHOULD be set to a value other than 0x00000000.");
            }

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSetSpooler.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC05_TestRopSetSpooler()
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
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            #region RopSetSpooler response

            RopSetSpoolerRequest setSpoolerRequest;
            RopSetSpoolerResponse setSpoolerResponse;

            setSpoolerRequest.RopId = (byte)RopId.RopSetSpooler;
            
            setSpoolerRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            setSpoolerRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopSetSpooler request.");

            // Send the RopSetSpooler request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setSpoolerRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setSpoolerResponse = (RopSetSpoolerResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setSpoolerResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSpoolerLockMessage.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC06_TestRopSpoolerLockMessage()
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

            // Step 1: Preparations-Create and save message.
            #region Create and save message

            // Create a message
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

            // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Save the created message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSaveChangesMessage request.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            ulong messageId = saveChangesMessageResponse.MessageId;

            #endregion

            #endregion

            // Step 2: Send the RopSetSpooler request and verify the success response.
            #region RopSetSpooler response

            RopSetSpoolerRequest setSpoolerRequest;
            RopSetSpoolerResponse setSpoolerResponse;

            setSpoolerRequest.RopId = (byte)RopId.RopSetSpooler;
            setSpoolerRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            setSpoolerRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSetSpooler request.");

            // Send the RopSetSpooler request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setSpoolerRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setSpoolerResponse = (RopSetSpoolerResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setSpoolerResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3: Send the RopSpoolerLockMessage request and verify the success response.
            #region RopSpoolerLockMessage response with the field LockState value is Lock

            RopSpoolerLockMessageRequest spoolerLockMessageRequest;

            spoolerLockMessageRequest.RopId = (byte)RopId.RopSpoolerLockMessage;
            spoolerLockMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            spoolerLockMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MessageId to that of created message, which identifies the message for which the status will be changed.
            spoolerLockMessageRequest.MessageId = messageId;

            spoolerLockMessageRequest.LockState = (byte)LockState.Lock;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSpoolerLockMessage request with the field LockState value is Lock.");

            // Send the RopSpoolerLockMessage request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                spoolerLockMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.Response);

            #endregion

            #region RopSpoolerLockMessage response with the field LockState value is Unlock

            spoolerLockMessageRequest.LockState = (byte)LockState.Unlock;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSpoolerLockMessage request with the field LockState value is Unlock.");

            // Send the RopSpoolerLockMessage request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                spoolerLockMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.Response);

            #endregion

            #region RopSpoolerLockMessage response with the field LockState value is Finished

            spoolerLockMessageRequest.LockState = (byte)LockState.Finished;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSpoolerLockMessage request with the field LockState value is Finished.");

            // Send the RopSpoolerLockMessage request to the server and verify the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                spoolerLockMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.Response);

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopTransportSend.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC07_TestRopTransportSend()
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
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            // Step 1: Preparations-Get transport folder, create, modify recipients and save message.
            #region Preparations

            // Get Transport folder.
            #region Get Transport folder

            RopGetTransportFolderRequest getTransportFolderRequest;
            RopGetTransportFolderResponse getTransportFolderResponse;

            getTransportFolderRequest.RopId = (byte)RopId.RopGetTransportFolder;
            
            getTransportFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            getTransportFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetTransportFolder request.");

            // Send the RopGetTransportFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getTransportFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getTransportFolderResponse = (RopGetTransportFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getTransportFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Create a message
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

            // Set FolderId to that of Transport folder, which identifies the parent folder.
            createMessageRequest.FolderId = getTransportFolderResponse.FolderId;

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

            #endregion

            // Configure recipients
            #region Configure recipients

            RopModifyRecipientsRequest modifyRecipientsRequest;
            RopModifyRecipientsResponse modifyRecipientsResponse;

            modifyRecipientsRequest.RopId = (byte)RopId.RopModifyRecipients;
            modifyRecipientsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            modifyRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Call CreateSampleRecipientColumnsAndRecipientRows method to create Recipient Rows and Recipient Columns.
            PropertyTag[] recipientColumns = null;
            ModifyRecipientRow[] recipientRows = null;
            this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);

            // Set ColumnCount, which specifies the number of rows in the RecipientRows field.
            modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;

            // Set RecipientColumns to that created above, which specifies the property values that can be included
            // for each recipient.
            modifyRecipientsRequest.RecipientColumns = recipientColumns;

            // Set RowCount, which specifies the number of rows in the RecipientRows field.
            modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;

            // Set RecipientRows to that created above, which is a list of ModifyRecipientRow structures.
            modifyRecipientsRequest.RecipientRows = recipientRows;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopModifyRecipients request.");

            // Send the RopModifyRecipients request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                modifyRecipientsRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                modifyRecipientsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Save message 
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSaveChangesMessage request.");

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
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Step 2: Send the RopTransportSend request and verify the success response.
            #region RopTransportSend success response

            RopTransportSendRequest transportSendRequest;
            RopTransportSendResponse transportSendResponse;

            transportSendRequest.RopId = (byte)RopId.RopTransportSend;
            transportSendRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            transportSendRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopTransportSend request.");

            // Send the RopTransportSend request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                transportSendRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            transportSendResponse = (RopTransportSendResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                transportSendResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Send the RopTransportSend request and verify the failure response.
            #region RopTransportSend failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            transportSendRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopTransportSend request.");

            // Send the RopTransportSend request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                transportSendRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            transportSendResponse = (RopTransportSendResponse)response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R457302");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R457302
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.ReturnValueForecNullObject,
                transportSendResponse.ReturnValue,
                457302,
                @"[in Error Codes Returned When an Object Is Invalid]ecNullObject [value] 0x000004B9 [Meaning] Returned when the client attempts to use a Server object handle value that was never assigned to an open object.");

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                transportSendResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send the RopTransportSend request using the Server object handle from a different logon and verify the failure response.
            #region RopTransportSend failure response of using the Server object handle from a different logon

            #region Logon to another logonId

            RopLogonRequest logonRequest;

            logonRequest.RopId = (byte)RopId.RopLogon;
            logonRequest.LogonId = TestSuiteBase.LogonId1;

            // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            logonRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex0;

            // Get user DN from configure file.
            string userDN = Common.GetConfigurationPropertyValue("UserEssdn", this.Site) + "\0";

            logonRequest.StoreState = (uint)StoreState.None;

            // Set other parameters for logon type of Mailbox (private mailbox).
            logonRequest.LogonFlags = (byte)LogonFlags.Private;
            logonRequest.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;

            // Set EssdnSize to the byte count of user DN, which specifies the size of the Essdn field.
            logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);

            // Set Essdn to the content of user DN, which specifies it will log on to the mail box of user represented by the user DN.
            logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);

            // Send the RopLogon request and get the response.
            this.responseSOHs = this.cropsAdapter.ProcessSingleRop(
                logonRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            this.inputObjHandle = this.responseSOHs[0][logonResponse.OutputHandleIndex];

            #endregion

            #region Send RopTransportSend request and get response

            // Set InputHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            transportSendRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopTransportSend request.");

            // Send the RopTransportSend request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                transportSendRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            transportSendResponse = (RopTransportSendResponse)response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R457303");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R457303
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.ReturnValueForecAccessDenied,
                transportSendResponse.ReturnValue,
                457303,
                @"[in Error Codes Returned When an Object Is Invalid]ecAccessDenied [value] 0x80070005 [Meaning] Returned when the client attempts to use the Server object handle from a different logon.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopTransportNewMail.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC08_TestRopTransportNewMail()
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

            // Step 1: Send the RopSetSpooler request and verify the success response.
            #region RopSetSpooler response

            RopSetSpoolerRequest setSpoolerRequest;
            RopSetSpoolerResponse setSpoolerResponse;

            setSpoolerRequest.RopId = (byte)RopId.RopSetSpooler;
            
            setSpoolerRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            setSpoolerRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSetSpooler request.");

            // Send the RopSetSpooler request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setSpoolerRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setSpoolerResponse = (RopSetSpoolerResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setSpoolerResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            if (Common.IsRequirementEnabled(1772, this.Site))
            {
                // Step 2: Create, modify recipients and save message.
                #region Create, modify recipients and save message

                // Create a message
                #region Create message

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

                // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
                createMessageRequest.FolderId = logonResponse.FolderIds[4];

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
                uint targetMessageHandle = responseSOHs[0][createMessageResponse.OutputHandleIndex];

                #endregion

                // Configure recipients
                #region Configure recipients

                RopModifyRecipientsRequest modifyRecipientsRequest;
                RopModifyRecipientsResponse modifyRecipientsResponse;

                modifyRecipientsRequest.RopId = (byte)RopId.RopModifyRecipients;
                modifyRecipientsRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                modifyRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Call CreateSampleRecipientColumnsAndRecipientRows method to create Recipient Rows and Recipient Columns.
                PropertyTag[] recipientColumns = null;
                ModifyRecipientRow[] recipientRows = null;
                this.CreateSampleRecipientColumnsAndRecipientRows(out recipientColumns, out recipientRows);

                // Set recipientColumns to null, which means no property values are included for each recipient.
                modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;

                // Set RecipientColumns to that created above, which specifies the property values that can be included
                // for each recipient.
                modifyRecipientsRequest.RecipientColumns = recipientColumns;

                // Set RowCount, which specifies the number of rows in the RecipientRows field.
                modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;

                // Set RecipientRows to that created above, which is a list of ModifyRecipientRow structures.
                modifyRecipientsRequest.RecipientRows = recipientRows;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopModifyRecipients request.");

                // Send the RopModifyRecipients request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    modifyRecipientsRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                modifyRecipientsResponse = (RopModifyRecipientsResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    modifyRecipientsResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Save the created message.
                #region Save message

                RopSaveChangesMessageRequest saveChangesMessageRequest;
                RopSaveChangesMessageResponse saveChangesMessageResponse;

                saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
                saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
                // in the response.
                saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

                saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSaveChangesMessage request.");

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
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");
                ulong messageId = saveChangesMessageResponse.MessageId;

                #endregion

                #endregion

                // Step 3: Send the RopTransportNewMail request and verify the success response.
                #region RopTransportNewMail response with the field MessageFlags value is MfRead

                RopTransportNewMailRequest transportNewMailRequest;
                RopTransportNewMailResponse transportNewMailResponse;

                transportNewMailRequest.RopId = (byte)RopId.RopTransportNewMail;
                transportNewMailRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                transportNewMailRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set MessageId to that of created message, which identifies the new Message object.
                transportNewMailRequest.MessageId = messageId;

                // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
                transportNewMailRequest.FolderId = logonResponse.FolderIds[4];

                // Set MessageClass, which specifies the message class of the new Message object.
                transportNewMailRequest.MessageClass = Encoding.ASCII.GetBytes(TestSuiteBase.MessageClassForRopTransportNewMail + "\0");

                transportNewMailRequest.MessageFlags = (byte)MessageFlags.MfRead;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopTransportNewMail request the field MessageFlags value is MfRead.");

                // Send the RopTransportNewMail request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    transportNewMailRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                transportNewMailResponse = (RopTransportNewMailResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    transportNewMailResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion

                #region RopTransportNewMail response with the field MessageFlags value is MfUnsent

                transportNewMailRequest.MessageFlags = (byte)MessageFlags.MfUnsent;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopTransportNewMail request the field MessageFlags value is MfUnsent.");

                // Send the RopTransportNewMail request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    transportNewMailRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                transportNewMailResponse = (RopTransportNewMailResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    transportNewMailResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion

                #region RopTransportNewMail response with the field MessageFlags value is MfResend

                transportNewMailRequest.MessageFlags = (byte)MessageFlags.MfResend;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopTransportNewMail request the field MessageFlags value is MfResend.");

                // Send the RopTransportNewMail request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    transportNewMailRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                transportNewMailResponse = (RopTransportNewMailResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    transportNewMailResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetTransportFolder.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S05_TC09_TestRopGetTransportFolder()
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
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

            // Step 1: Send the RopGetTransportFolder request and verify the success response.
            #region RopGetTransportFolder success response

            RopGetTransportFolderRequest getTransportFolderRequest;
            RopGetTransportFolderResponse getTransportFolderResponse;

            getTransportFolderRequest.RopId = (byte)RopId.RopGetTransportFolder;
            
            getTransportFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            getTransportFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetTransportFolder request.");

            // Send the RopGetTransportFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getTransportFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getTransportFolderResponse = (RopGetTransportFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getTransportFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 2: Send the RopGetTransportFolder request and verify the failure response.
            #region RopGetTransportFolder failure response

            // Refer to MS-OXCROPS endnote<14>: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the Server object 
            // and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Set InputHandleIndex to 0x01, which is an invalid index.
                getTransportFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetTransportFolder request to invoke failure response.");

                // Send the RopGetTransportFolder request to the server and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getTransportFolderRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getTransportFolderResponse = (RopGetTransportFolderResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getTransportFolderResponse.ReturnValue,
                    "<14> Section 3.2.5.1: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.");
            }
            else
            {
                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                getTransportFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetTransportFolder request to invoke failure response.");

                // Send the RopGetTransportFolder request to the server and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getTransportFolderRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getTransportFolderResponse = (RopGetTransportFolderResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getTransportFolderResponse.ReturnValue,
                    "For this response, this field is set to a value other than 0x00000000.");
            }

            #endregion
        }

        #endregion

        #region Common method

        /// <summary>
        /// Delete subfolder and message
        /// </summary>
        /// <param name="folderId">The id of folder which its subfolder and message need delete</param>
        private void DeleteSubFolderAndMessage(ulong folderId)
        {
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

            // Set FolderId to the 5th of logonResponse, this folder is to be opened.
            openFolderRequest.FolderId = folderId;

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopOpenFolder request.");

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

            // Get the folder handle, this handle will be used as input handle in the following RopHardDeleteMessageAndSubfolders.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

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
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopHardDeleteMessagesAndSubfolders request.");

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
        }
        
        #endregion
    }
}