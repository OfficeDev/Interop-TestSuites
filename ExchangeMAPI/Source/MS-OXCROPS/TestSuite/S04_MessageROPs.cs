namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Message ROPs. 
    /// </summary>
    [TestClass]
    public class S04_MessageROPs : TestSuiteBase
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
        /// This method tests ROP buffers of RopCreateMessage, RopModifyRecipients, RopSaveChangesMessage, RopReadRecipients,
        /// RopRemoveAllRecipients, RopOpenMessag and RopReloadCachedInformationRequest.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC01_TestRopReloadCachedInformation()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCreateMessage request and verify the success response.
            #region RopCreateMessage success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

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

            // Send the RopCreateMessage request and verify the success response.
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

            if (createMessageResponse.HasMessageId == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1778");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1778
                // MessageId is null means not present
                Site.CaptureRequirementIfAreEqual<ulong?>(
                    null,
                    createMessageResponse.MessageId,
                    1778,
                    @"[In RopCreateMessage ROP Success Response Buffer,MessageId (8 bytes)]is not present if it[HasMessageId] is zero.");
            }

            #endregion

            // Step 2: Send the RopModifyRecipients request and verify the success response.
            #region RopModifyRecipients success response

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
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopModifyRecipients request.");

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

            // Step 3: Send the RopSaveChangesMessage request and verify the success response.
            #region RopSaveChangesMessage success response

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSaveChangesMessage request.");

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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // targetMessageHandle is OutputHandle of createMessageResponse and specified by the output index in the ROP createMessage request
            // program go here, then targetMessageHandle is useful
            Site.CaptureRequirement(
                4558,
                @"[In Processing a ROP Input Buffer] The handle assigned is then set in the Server object handle table at the location specified by the output index in the ROP request.");

            #endregion

            // Step 4: Send the RopReadRecipients request and verify the success response.
            #region RopReadRecipients success response

            RopReadRecipientsRequest readRecipientsRequest;
            RopReadRecipientsResponse readRecipientsResponse;

            readRecipientsRequest.RopId = (byte)RopId.RopReadRecipients;
            readRecipientsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            readRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set RowId, which specifies the recipient to start reading.
            readRecipientsRequest.RowId = TestSuiteBase.RowId;

            // Set Reserved, this field is reserved and MUST be set to 0.
            readRecipientsRequest.Reserved = TestSuiteBase.Reserved;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopReadRecipients request.");

            // Send the RopReadRecipients request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readRecipientsRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            readRecipientsResponse = (RopReadRecipientsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                readRecipientsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 5: Send the RopOpenMessag request and verify the success response.
            #region RopOpenMessag success response

            RopOpenMessageRequest openMessageRequest;
            RopOpenMessageResponse openMessageResponse;

            openMessageRequest.RopId = (byte)RopId.RopOpenMessage;
            openMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            openMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            openMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            openMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
            openMessageRequest.FolderId = logonResponse.FolderIds[4];

            openMessageRequest.OpenModeFlags = (byte)MessageOpenModeFlags.ReadOnly;

            // Set MessageId to that of created message, which identifies the message to be opened.
            openMessageRequest.MessageId = saveChangesMessageResponse.MessageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopOpenMessage request.");

            // Send the RopOpenMessage request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openMessageResponse = (RopOpenMessageResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openMessageResponse.ReturnValue,
            "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 6: Send the RopReloadCachedInformation request and verify the success response.
            #region RopReloadCachedInformation success response

            RopReloadCachedInformationRequest reloadCachedInformationRequest;
            RopReloadCachedInformationResponse reloadCachedInformationResponse;

            reloadCachedInformationRequest.RopId = (byte)RopId.RopReloadCachedInformation;
            reloadCachedInformationRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            reloadCachedInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set Reserved to 0x0000, this field is reserved and MUST be set to 0x0000.
            reloadCachedInformationRequest.Reserved = TestSuiteBase.Reserved;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopReloadCachedInformation request.");

            // Send the RopReloadCachedInformation request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                reloadCachedInformationRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            reloadCachedInformationResponse = (RopReloadCachedInformationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                reloadCachedInformationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 7: Send the RopRemoveAllRecipients request and verify the success response.
            #region RopRemoveAllRecipients response

            RopRemoveAllRecipientsRequest removeAllRecipientsRequest;
            RopRemoveAllRecipientsResponse removeAllRecipientsResponse;

            removeAllRecipientsRequest.RopId = (byte)RopId.RopRemoveAllRecipients;
            removeAllRecipientsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            removeAllRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set Reserved to 0x0, this field is reserved and MUST be set to 0x00000000.
            removeAllRecipientsRequest.Reserved = TestSuiteBase.Reserved;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopRemoveAllRecipients request.");

            // Send the RopRemoveAllRecipients request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                removeAllRecipientsRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                removeAllRecipientsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 8: Verify Reserved field and verify R2005
            #region Verify Reserved field and verify R2005

            // Set Reserved non-0x00000000 to get the reply.
            removeAllRecipientsRequest.Reserved = TestSuiteBase.Reserved;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopRemoveAllRecipients request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                removeAllRecipientsRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopRemoveAllRecipientsResponse removeAllRecipientsResponse1 = (RopRemoveAllRecipientsResponse)response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2005");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2005
            // If two response return value is the same, the reply is the same.
            Site.CaptureRequirementIfAreEqual<uint>(
                removeAllRecipientsResponse.ReturnValue,
                removeAllRecipientsResponse1.ReturnValue,
                2005,
                @"[In RopRemoveAllRecipients ROP Request Buffer,Reserved (4 bytes)]Reply is the same whether 0x00000000 or non-0x00000000 is used for this Field[Reserved].");

            #endregion

            // Step 9: Send the RopCreateMessage request and verify the failure response.
            #region RopCreateMessage failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createMessageRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            createMessageResponse = (RopCreateMessageResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createMessageResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 10: Send the RopSaveChangesMessage request and verify the failure response.
            #region RopSaveChangesMessage failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 10: Begin to send the RopSaveChangesMessage request.");

            // Send the RopSaveChangesMessage request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesMessageResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 11: Send the RopReadRecipients request and verify the failure response.
            #region RopReadRecipients failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            readRecipientsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 11: Begin to send the RopReadRecipients request.");

            // Send the RopReadRecipients request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readRecipientsRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            readRecipientsResponse = (RopReadRecipientsResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                readRecipientsResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 12: Send the RopOpenMessage request and verify the failure response.
            #region RopOpenMessage failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            openMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 12: Begin to send the RopOpenMessage request.");

            // Send the RopOpenMessage request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openMessageRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            openMessageResponse = (RopOpenMessageResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openMessageResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 13: Send the RopReloadCachedInformation request and verify the failure response.
            #region RopReloadCachedInformation failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            reloadCachedInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 13: Begin to send the RopReloadCachedInformation request.");

            // Send the RopReloadCachedInformation request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                reloadCachedInformationRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            reloadCachedInformationResponse = (RopReloadCachedInformationResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                reloadCachedInformationResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSetMessageStatus and RopGetMessageStatus.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC02_TestRopGetMessageStatus()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create and save message, then open the folder containing created message.
            #region Create and save message, open folder containing created message

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create message.
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

            // Save message.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            ulong messageId = saveChangesMessageResponse.MessageId;

            // Open the folder(Inbox) containing the created message.
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

            // Set FolderId to the 5th folder of the logonResponse(INBOX), which specifies the folder to be opened.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint folderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSetMessageStatus request and verify the success response.
            #region RopSetMessageStatus success response

            RopSetMessageStatusRequest setMessageStatusRequest;
            RopSetMessageStatusResponse setMessageStatusResponse;

            setMessageStatusRequest.RopId = (byte)RopId.RopSetMessageStatus;
            setMessageStatusRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            setMessageStatusRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MessageId to that of created message, which value identifies the message for which the status will be changed.
            setMessageStatusRequest.MessageId = messageId;

            setMessageStatusRequest.MessageStatusFlags = (uint)MessageStatusFlags.MsRemoteDownload;

            // Set MessageStatusMask, which specifies which bits in the MessageStatusFlags field are to be changed.
            setMessageStatusRequest.MessageStatusMask = TestSuiteBase.MessageStatusMask;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSetMessageStatus request.");

            // Send the RopSetMessageStatus request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setMessageStatusRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setMessageStatusResponse = (RopSetMessageStatusResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setMessageStatusResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3: Send the RopGetMessageStatus request and verify the success response.
            #region RopGetMessageStatus success response

            RopGetMessageStatusRequest getMessageStatusRequest;

            getMessageStatusRequest.RopId = (byte)RopId.RopGetMessageStatus;
            getMessageStatusRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            getMessageStatusRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MessageId to that of created message, which identifies the message for which the status will be returned.
            getMessageStatusRequest.MessageId = messageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetMessageStatus request.");

            // Send the RopGetMessageStatus request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getMessageStatusRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // The response buffers for RopGetMessageStatus are the same as those for RopSetMessageStatus.
            setMessageStatusResponse = (RopSetMessageStatusResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setMessageStatusResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 4: Send the RopSetMessageStatus request and verify the failure response.
            #region RopSetMessageStatus failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            setMessageStatusRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetMessageStatus request.");

            // Send the RopSetMessageStatus request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setMessageStatusRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            setMessageStatusResponse = (RopSetMessageStatusResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setMessageStatusResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion

            // Step 5: Send the RopGetMessageStatus request and verify the failure response.
            #region RopGetMessageStatus failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            getMessageStatusRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Set MessageId to that of created message, which identifies the message for which the status will be returned.
            getMessageStatusRequest.MessageId = messageId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetMessageStatus request.");

            // Send the RopGetMessageStatus request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getMessageStatusRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            // The response buffers for RopGetMessageStatus are the same as those for RopSetMessageStatus
            // MS-OXCROPS 2.2.5.9.2
            setMessageStatusResponse = (RopSetMessageStatusResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setMessageStatusResponse.ReturnValue,
               "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSetReadFlags.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC03_TestRopSetReadFlags()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create and save message, then open the folder containing created message.            
            #region Create and save message, open folder containing created message

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Create a message.
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

            // Save message.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            ulong messageId = saveChangesMessageResponse.MessageId;

            // Open the folder(Inbox) containing the created message.
            // The folder handle will be used as input handle in next ROP: SetReadFlags
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

            // Set FolderId to the 5th folder of the logonResponse(INBOX), which specifies the folder to be opened.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            uint folderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2: Send the RopSetReadFlags request and verify the success response.
            #region RopSetReadFlags response

            RopSetReadFlagsRequest setReadFlagsRequest;
            RopSetReadFlagsResponse setReadFlagsResponse;

            setReadFlagsRequest.RopId = (byte)RopId.RopSetReadFlags;
            setReadFlagsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            setReadFlagsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set WantAsynchronous to 0x00(FALSE), which specifies the operation is to be executed synchronously.
            setReadFlagsRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);

            setReadFlagsRequest.ReadFlags = (byte)ReadFlags.Default;
            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageId;

            // Set MessageIdCount, which specifies the number of identifiers in the MessageIds field.
            setReadFlagsRequest.MessageIdCount = (ushort)messageIds.Length;

            // Set MessageIds, which specify the messages that are to have their read flags changed.
            setReadFlagsRequest.MessageIds = messageIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSetReadFlags request.");

            // Send the RopSetReadFlags request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setReadFlagsRequest,
                folderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setReadFlagsResponse = (RopSetReadFlagsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setReadFlagsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopSetMessageReadFlag.
        /// </summary>
         [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC04_TestRopSetMessageReadFlags()
        {
            this.CheckTransportIsSupported();

            // Check whether the environment supported public folder.
            if (bool.Parse(Common.GetConfigurationPropertyValue("IsPublicFolderSupported", this.Site)))
            {
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PublicFolderServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));

                // Step 1: Open the second folder and create a subfolder.
                #region Open the second folder to create public folder under root folder

                // Log on to a private mailbox.
                RopLogonResponse logonResponse = Logon(LogonType.PublicFolder, this.userDN, out inputObjHandle);

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

                // Set FolderId to the second folder of the logonResponse, which specifies the folder to be opened.
                openFolderRequest.FolderId = logonResponse.FolderIds[1];

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

                // Step 2: Create a non-ghosted public folder.
                #region Create a non-ghosted public folder

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

                // Set Reserved to 0x0, this field is reserved and MUST be set to 0.
                createFolderRequest.Reserved = TestSuiteBase.Reserved;

                // Set new value for DisplayNameAndCommentForNonSearchFolder in Public Folder
                string displayNameAndCommentForNonSearchFolder = "MS-OXCROPSPublicFolder" + DisplayNameAndCommentForNonSearchFolder;

                // Set DisplayName, which specifies the name of the created folder.
                createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(displayNameAndCommentForNonSearchFolder + "\0");

                // Set Comment, which specifies the folder comment that is associated with the created folder.
                createFolderRequest.Comment = Encoding.ASCII.GetBytes(displayNameAndCommentForNonSearchFolder + "\0");

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
                ulong folderId = createFolderResponse.FolderId;

                #endregion

                // Step 3: Open the folder containing the created message.
                #region Open folder

                openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
                openFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
                // for the input Server object is stored.
                openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
                // for the output Server object will be stored.
                openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                // Set FolderId to that of created folder, which specifies the folder to be opened.
                openFolderRequest.FolderId = folderId;

                openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenFolder request.");

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
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 4: Create and save message.
                #region Create and save message

                #region Create a message.
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

                // Set FolderId to that of created folder, which identifies the parent folder.
                createMessageRequest.FolderId = createFolderResponse.FolderId;

                // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
                createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopCreateMessage request.");

                // Send the RopCreateMessage request and verify the success response.
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

                #region RopModifyRecipients response

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
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopModifyRecipients request.");

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

                #region Save message

                RopSaveChangesMessageRequest saveChangesMessageRequest;
                RopSaveChangesMessageResponse saveChangesMessageResponse;

                saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
                saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
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
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                #endregion

                // Step 5: Send the RopCreateMessage request and verify the success response.
                #region RopLongTermIdFromId success response

                RopLongTermIdFromIdRequest longTermIdFromIdRequest;
                RopLongTermIdFromIdResponse longTermIdFromIdResponse;

                longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
                longTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
                // for the input Server object is stored.
                longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set ObjectId to that got in the foregoing code, this id will be converted to a short-term ID.
                longTermIdFromIdRequest.ObjectId = saveChangesMessageResponse.MessageId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopLongTermIdFromId request.");

                // Send the RopLongTermIdFromId request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    longTermIdFromIdRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    longTermIdFromIdResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 6: Send the RopOpenMessag request and verify the success response.
                #region RopOpenMessag success response

                RopOpenMessageRequest openMessageRequest;
                RopOpenMessageResponse openMessageResponse;

                openMessageRequest.RopId = (byte)RopId.RopOpenMessage;
                openMessageRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                openMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
                openMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
                openMessageRequest.CodePageId = TestSuiteBase.CodePageId;

                // Set FolderId to that of created folder, which identifies the parent folder.
                openMessageRequest.FolderId = createFolderResponse.FolderId;

                openMessageRequest.OpenModeFlags = (byte)MessageOpenModeFlags.ReadWrite;

                // Set MessageId to that of created message, which identifies the message to be opened.
                openMessageRequest.MessageId = saveChangesMessageResponse.MessageId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopCreateMessage request.");

                // Send the RopCreateMessage request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    openMessageRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                openMessageResponse = (RopOpenMessageResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    openMessageResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion

                // Step 7: Send the RopSetMessageReadFlag request and verify the success response.
                #region RopSetMessageReadFlag success response Test not changed

                RopSetMessageReadFlagRequest setMessageReadFlagRequest;
                RopSetMessageReadFlagResponse setMessageReadFlagResponse;

                setMessageReadFlagRequest.RopId = (byte)RopId.RopSetMessageReadFlag;
                setMessageReadFlagRequest.LogonId = TestSuiteBase.LogonId;

                // Set ResponseHandleIndex, which specifies the location in the Server object handle table that is referenced
                // in the response.
                setMessageReadFlagRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                setMessageReadFlagRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.GenerateReceiptOnly;

                // Set ClientData, which specifies the information that is returned to the client in a successful response.
                setMessageReadFlagRequest.ClientData = longTermIdFromIdResponse.LongTermId.Serialize();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopSetMessageReadFlag request.");

                // Send the RopSetMessageReadFlag request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    setMessageReadFlagRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    setMessageReadFlagResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                #region Verify R2110 and R2114

                if (setMessageReadFlagResponse.ReadStatusChanged == Convert.ToByte(TestSuiteBase.Zero))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2110");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R2110
                    // LogonId is null means not present.
                    Site.CaptureRequirementIfAreEqual<byte?>(
                        null,
                        setMessageReadFlagResponse.LogonId,
                        2110,
                        @"[In RopSetMessageReadFlag ROP Success Response Buffer,LogonId (1 byte)]is not present otherwise[when the value in the ReadStatusChanged field is zero].");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2114");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R2114
                    // ClientData is null means not present.
                    Site.CaptureRequirementIfIsNull(
                        setMessageReadFlagResponse.ClientData,
                        2114,
                        @"[In RopSetMessageReadFlag ROP Success Response Buffer,ClientData (24 bytes)]is not present otherwise[when the value in the ReadStatusChanged field is zero].");
                }

                #endregion

                // Step 8: Send the RopSetMessageReadFlag request and verify the success response.
                #region RopSetMessageReadFlag success response Test changed

                setMessageReadFlagRequest.RopId = (byte)RopId.RopSetMessageReadFlag;
                setMessageReadFlagRequest.LogonId = TestSuiteBase.LogonId;

                // Set ResponseHandleIndex, which specifies the location in the Server object handle table that is referenced
                // in the response.
                setMessageReadFlagRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                setMessageReadFlagRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.ClearReadFlag;

                // Set ClientData, which specifies the information that is returned to the client in a successful response.
                setMessageReadFlagRequest.ClientData = longTermIdFromIdResponse.LongTermId.Serialize();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopSetMessageReadFlag request.");

                // Send the RopSetMessageReadFlag request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    setMessageReadFlagRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    setMessageReadFlagResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                #region Verify R2109 and R2113

                if (setMessageReadFlagResponse.ReadStatusChanged != Convert.ToByte(TestSuiteBase.Zero))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2109");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R2109
                    // LogonId is not null means present.
                    Site.CaptureRequirementIfIsNotNull(
                        setMessageReadFlagResponse.LogonId,
                        2109,
                        @"[In RopSetMessageReadFlag ROP Success Response Buffer,LogonId (1 byte)]This field is present when the value in the ReadStatusChanged field is nonzero.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2113");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R2113
                    // ClientData is not null means present.
                    Site.CaptureRequirementIfIsNotNull(
                        setMessageReadFlagResponse.ClientData,
                        2113,
                        @"[In RopSetMessageReadFlag ROP Success Response Buffer,ClientData (24 bytes)]This field is present when the value in the ReadStatusChanged field is nonzero.");
                }

                #endregion

                // Step 9: Send the RopSetMessageReadFlag request and verify the failure response.
                #region RopSetMessageReadFlag failure response

                // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
                setMessageReadFlagRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopSetMessageReadFlag request.");

                // Send the RopSetMessageReadFlag request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    setMessageReadFlagRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    setMessageReadFlagResponse.ReturnValue,
                   "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion

                // Step 10: Send a RopDeleteFolder request to the server.
                #region RopDeleteFolder Response
                RopDeleteFolderRequest deleteFolderRequest;

                deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
                deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                deleteFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelMessages;

                // Set FolderId to targetFolderId, this folder is to be deleted.
                deleteFolderRequest.FolderId = folderId;

                // Send a RopDeleteFolder request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    deleteFolderRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                #endregion
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the first system supports public folder logon.");
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopCreateAttachment, RopSaveChangesAttachment, RopOpenAttachment and RopDeleteAttachment.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC05_TestRopOpenAndDeleteAttachment()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Create and save message.
            #region Create and save message

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

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

            // Send the RopCreateMessage request and verify the success response.
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

            // Save message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            #endregion

            // Step 2: Send the RopCreateAttachment request and verify the success response.
            #region RopCreateAttachment success response

            RopCreateAttachmentRequest createAttachmentRequest;
            RopCreateAttachmentResponse createAttachmentResponse;

            createAttachmentRequest.RopId = (byte)RopId.RopCreateAttachment;
            createAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            createAttachmentRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCreateAttachment request.");

            // Send the RopCreateAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createAttachmentResponse = (RopCreateAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createAttachmentResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint attachmentHandle = responseSOHs[0][createAttachmentResponse.OutputHandleIndex];
            uint attachmentId = createAttachmentResponse.AttachmentID;

            #endregion

            // Step 3: Send the RopSaveChangesAttachment request and verify the success response.
            #region RopSaveChangesAttachment success response

            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest;
            RopSaveChangesAttachmentResponse saveChangesAttachmentReponse;

            saveChangesAttachmentRequest.RopId = (byte)RopId.RopSaveChangesAttachment;
            saveChangesAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesAttachmentRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            saveChangesAttachmentRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSaveChangesAttachment request.");

            // Send the RopSaveChangesAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesAttachmentRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesAttachmentReponse = (RopSaveChangesAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesAttachmentReponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 4: Send the RopGetPropertiesAll request and verify the success response.
            #region RopGetPropertiesAll success response

            RopGetPropertiesAllRequest getPropertiesAllRequest;
            RopGetPropertiesAllResponse getPropertiesAllResponse;

            getPropertiesAllRequest.RopId = (byte)RopId.RopGetPropertiesAll;

            getPropertiesAllRequest.LogonId = TestSuiteBase.LogonId;
            getPropertiesAllRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned.
            getPropertiesAllRequest.PropertySizeLimit = TestSuiteBase.PropertySizeLimit;

            getPropertiesAllRequest.WantUnicode = (ushort)Zero;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetPropertiesAll request.");

            // Send the RopGetPropertiesAll request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesAllRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesAllResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            int counter = 0;
            byte[] tempArray = BitConverter.GetBytes(attachmentId);
            foreach (TaggedPropertyValue item in getPropertiesAllResponse.PropertyValues)
            {
                // Refer to MS-OXPROPS, the Property ID of PidTagAttachNumber is 0x0E21.
                if (item.PropertyTag.PropertyId == 0x0E21)
                {
                    if (item.Value == null || tempArray == null)
                    {
                        break;
                    }
                    else if (item.Value.Length == 0 || tempArray.Length == 0)
                    {
                        break;
                    }
                    else if (item.Value.Length != tempArray.Length)
                    {
                        break;
                    }
                    else
                    {
                        for (int i = 0; i < tempArray.Length; i++)
                        {
                            if (item.Value[i] == tempArray[i])
                            {
                                counter++;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            bool isVerifyR2175 = counter == tempArray.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2175");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2175
            Site.CaptureRequirementIfIsTrue(
                isVerifyR2175,
                2175,
                @"[In RopCreateAttachment ROP Success Response Buffer,AttachmentID (4 bytes)] The value of this field is equivalent to the value of the PidTagAttachNumber property ([MS-OXCMSG] section 2.2.2.6).");
            #endregion

            // Step 5: Send the RopSaveChangesAttachment request and verify the success response.
            #region RopOpenAttachment success response

            RopOpenAttachmentRequest openAttachmentRequest;
            RopOpenAttachmentResponse openAttachmentResponse;

            openAttachmentRequest.RopId = (byte)RopId.RopOpenAttachment;
            openAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            openAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            openAttachmentRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            openAttachmentRequest.OpenAttachmentFlags = (byte)OpenAttachmentFlags.ReadOnly;

            // Set AttachmentID, which identifies the attachment to be opened.
            openAttachmentRequest.AttachmentID = attachmentId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopOpenAttachment request.");

            // Send the RopOpenAttachment request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openAttachmentResponse = (RopOpenAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openAttachmentResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 6: Send the RopSaveChangesAttachment request and verify the success response.
            #region  RopDeleteAttachment success response

            RopDeleteAttachmentRequest deleteAttachmentRequest;
            RopDeleteAttachmentResponse deleteAttachmentResponse;

            deleteAttachmentRequest.RopId = (byte)RopId.RopDeleteAttachment;
            deleteAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            deleteAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set AttachmentID, which e identifies the attachment to be deleted.
            deleteAttachmentRequest.AttachmentID = attachmentId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopDeleteAttachment request.");

            // Send the RopDeleteAttachment request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deleteAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            deleteAttachmentResponse = (RopDeleteAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                deleteAttachmentResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 7: Send the RopCreateAttachment request and verify the failure response.
            #region RopCreateAttachment failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            createAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopCreateAttachment request.");

            // Send the RopCreateAttachment request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            createAttachmentResponse = (RopCreateAttachmentResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createAttachmentResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetAttachmentTable.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC06_TestRopGetAttachmentTableAndGetValidAttachments()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create and save message, then create and save attachment.
            #region Create and save message, create and save attachment

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

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

            // Set FolderId to the 4th of logonResponse, which identifies the parent folder.
            createMessageRequest.FolderId = logonResponse.FolderIds[4];

            // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
            createMessageRequest.AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateMessage request.");

            // Send the RopCreateMessage request and verify the success response.
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

            // Save message.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Create an attachment to above message.
            #region Create an attachment

            RopCreateAttachmentRequest createAttachmentRequest;
            RopCreateAttachmentResponse createAttachmentResponse;

            createAttachmentRequest.RopId = (byte)RopId.RopCreateAttachment;
            createAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            createAttachmentRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateAttachment request.");

            // Send the RopCreateAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createAttachmentResponse = (RopCreateAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createAttachmentResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint attachmentHandle = responseSOHs[0][createAttachmentResponse.OutputHandleIndex];

            #endregion

            // Save the attachment.
            #region Save the attachment

            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest;
            RopSaveChangesAttachmentResponse saveChangesAttachmentReponse;

            saveChangesAttachmentRequest.RopId = (byte)RopId.RopSaveChangesAttachment;
            saveChangesAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesAttachmentRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            saveChangesAttachmentRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSaveChangesAttachment request.");

            // Send the RopSaveChangesAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesAttachmentRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesAttachmentReponse = (RopSaveChangesAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesAttachmentReponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Refer to MS-OXCMSG: Exchange 2010 and Exchange 2013 do not support the RopGetValidAttachments ROP.
            if (Common.IsRequirementEnabled(171501, this.Site))
            {
                // Step 2: Send the RopGetValidAttachments request to the server and verify the success response.
                #region RopGetValidAttachmentsRequest success response

                RopGetValidAttachmentsRequest getValidAttachmentsRequest;
                RopGetValidAttachmentsResponse getValidAttachmentsResponse;

                getValidAttachmentsRequest.RopId = (byte)RopId.RopGetValidAttachments;
                getValidAttachmentsRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                getValidAttachmentsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetValidAttachments request.");

                // Send the RopGetValidAttachments request and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getValidAttachmentsRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getValidAttachmentsResponse = (RopGetValidAttachmentsResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R171501");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R071501
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getValidAttachmentsResponse.ReturnValue,
                    171501,
                    @"[In Appendix A: Product Behavior] Implementation does support the RopgetValidAttachments ROP. (Exchange 2007 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getValidAttachmentsResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion

                // Step 3: Send the RopGetValidAttachments request to the server and verify the failure response.
                #region RopGetValidAttachments failure response

                getValidAttachmentsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetValidAttachments request.");

                // Send the RopGetValidAttachments request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getValidAttachmentsRequest,
                    targetMessageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getValidAttachmentsResponse = (RopGetValidAttachmentsResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getValidAttachmentsResponse.ReturnValue,
                    "if ROP failure, the ReturnValue of its response is not 0(failure)");

                #endregion
            }

            // Step 4: Send the RopGetAttachmentTable request to the server and verify the success response.
            #region RopGetAttachmentTable response

            RopGetAttachmentTableRequest getAttachmentTableRequest;
            RopGetAttachmentTableResponse getAttachmentTableResponse;

            getAttachmentTableRequest.RopId = (byte)RopId.RopGetAttachmentTable;
            getAttachmentTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            getAttachmentTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            getAttachmentTableRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            getAttachmentTableRequest.TableFlags = (byte)MsgTableFlags.Standard;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetAttachmentTable request.");

            // Send the RopGetAttachmentTable request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getAttachmentTableRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getAttachmentTableResponse = (RopGetAttachmentTableResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getAttachmentTableResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopOpenEmbeddedMessage.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC07_TestRopOpenEmbeddedMessage()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create, modify recipients and save message, then create and save attachment on the message
            #region Create, modify recipients and save message, then create and save attachment on the message.

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

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

            // Send the RopCreateMessage request and verify the success response.
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

            // Modify recipients.
            #region Modify recipients

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

            // Save message.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            // Set RopId to 0x0C, which specifies the type of ROP is RopSaveChangesMessage.
            saveChangesMessageRequest.RopId = 0x0C;

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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Create attachment.
            #region Create attachment

            RopCreateAttachmentRequest createAttachmentRequest;
            RopCreateAttachmentResponse createAttachmentResponse;

            createAttachmentRequest.RopId = (byte)RopId.RopCreateAttachment;
            createAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            createAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            createAttachmentRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCreateAttachment request.");

            // Send the RopCreateAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                createAttachmentRequest,
                targetMessageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            createAttachmentResponse = (RopCreateAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                createAttachmentResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint attachmentHandle = responseSOHs[0][createAttachmentResponse.OutputHandleIndex];

            #endregion

            // Save attachment.
            #region Save attachment

            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest;
            RopSaveChangesAttachmentResponse saveChangesAttachmentReponse;

            saveChangesAttachmentRequest.RopId = (byte)RopId.RopSaveChangesAttachment;
            saveChangesAttachmentRequest.LogonId = TestSuiteBase.LogonId;

            // Set ResponseHandleIndex to 0x01, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesAttachmentRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex1;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesAttachmentRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            saveChangesAttachmentRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopSaveChangesAttachment request.");

            // Send the RopSaveChangesAttachment request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                saveChangesAttachmentRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            saveChangesAttachmentReponse = (RopSaveChangesAttachmentResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                saveChangesAttachmentReponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            #endregion

            // Step 2: Send the RopOpenEmbeddedMessage request and verify the success response.
            #region RopOpenEmbeddedMessage success response

            RopOpenEmbeddedMessageRequest openEmbeddedMessageRequest;
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse;

            openEmbeddedMessageRequest.RopId = (byte)RopId.RopOpenEmbeddedMessage;
            openEmbeddedMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            openEmbeddedMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            openEmbeddedMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
            openEmbeddedMessageRequest.CodePageId = TestSuiteBase.CodePageId;

            openEmbeddedMessageRequest.OpenModeFlags = (byte)EmbeddedMessageOpenModeFlags.Create;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOpenEmbeddedMessage request.");

            // Send the RopOpenEmbeddedMessage request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openEmbeddedMessageRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openEmbeddedMessageResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Send the RopOpenEmbeddedMessage request and verify the failure response.
            #region RopOpenEmbeddedMessage failure response

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            openEmbeddedMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopOpenEmbeddedMessage request.");

            // Send the RopOpenEmbeddedMessage request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openEmbeddedMessageRequest,
                attachmentHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openEmbeddedMessageResponse.ReturnValue,
                "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

            #endregion
        }

        /// <summary>
        /// This method tests multiple ROPs in a ROP request buffer by using operation RopOpenFolder, RopCreateMessage and RopSaveMessage. 
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S04_TC08_TestMultipleROPs()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Construct RopOpenFolder request.
            #region Open folder

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopOpenFolderRequest openFolderRequest;
            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle
            // for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table where the handle
            // for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to the 5th folder of the logonResponse(INBOX), which specifies the folder to be opened.
            openFolderRequest.FolderId = logonResponse.FolderIds[4];

            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            #endregion

            // Step 2: Construct RopCreateMessage request.
            #region Create message

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = TestSuiteBase.LogonId,

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
                InputHandleIndex = TestSuiteBase.InputHandleIndex1,
                
                // Set OutputHandleIndex to 0x02, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
                OutputHandleIndex = TestSuiteBase.OutputHandleIndex2,

                // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used.
                CodePageId = TestSuiteBase.CodePageId,

                // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder.
                FolderId = logonResponse.FolderIds[4],

                // Set AssociatedFlag to 0x00, which specified this message is not a folder associated information (FAI) message.
                AssociatedFlag = Convert.ToByte(TestSuiteBase.Zero)
            };
            #endregion

            // Step 3: Construct RopSaveChangesMessage request.
            #region Save message

            RopSaveChangesMessageRequest saveChangesMessageRequest;

            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x02, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex2;

            // Set ResponseHandleIndex to 0x03, which specifies the location in the Server object handle table that is referenced
            // in the response.
            saveChangesMessageRequest.ResponseHandleIndex = TestSuiteBase.ResponseHandleIndex2;

            saveChangesMessageRequest.SaveFlags = (byte)SaveFlags.ForceSave;

            #endregion

            // Step 4: Send multiple ROPs request.
            #region Send multiple ROPs request

            List<ISerializable> ropRequests = new List<ISerializable>
            {
                openFolderRequest,
                createMessageRequest,
                saveChangesMessageRequest
            };

            List<uint> inputObjects = new List<uint>
            {
                // 0xFFFF indicates a default input handle.
                this.inputObjHandle, 0xFFFF, 0xFFFF, 0xFFFF
            };

            List<IDeserializable> ropResponses = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the multiple ROPs request.");

            // Send multiple ROPs request.
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(
                ropRequests,
                inputObjects,
                ref ropResponses,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion

            #region Verify R4559, R4550, R4565 and R4719

            RopSaveChangesMessageResponse ropSaveChangesMessageResponse = (RopSaveChangesMessageResponse)ropResponses[2];
            bool isVerifyR4559 = ropSaveChangesMessageResponse.ReturnValue == 0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4559,the actual value is {0}.", isVerifyR4559);

            // The RopSaveChangesMessageRequest uses the handle of created message, so success of RopSaveChangesMessage
            // means R4559 is verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR4559,
                4559,
                @"[In Processing a ROP Input Buffer,The handle assigned is then set in the Server object handle table at the location specified by the output index in the ROP request] And can be used by subsequent ROP requests in the same ROP input buffer.");

            // The request order is openFolderRequest, createMessageRequest, saveChangesMessageRequest,
            // isVerifyR4550 indicates whether the orders in request and response are same.
            bool isVerifyR4550 = (typeof(RopOpenFolderResponse) == ropResponses[0].GetType())
                                 && (typeof(RopCreateMessageResponse) == ropResponses[1].GetType())
                                 && (typeof(RopSaveChangesMessageResponse) == ropResponses[2].GetType());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4550,the actual value is {0}.", isVerifyR4550);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4550
            // True means the orders in request and response are same.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR4550,
                4550,
                @"[In Processing a ROP Input Buffer] The ROP responses in the  ROP output buffer MUST  be in the same order in which they were processed.");

            // The request order is openFolderRequest, createMessageRequest, saveChangesMessageRequest,
            // isVerifyR4719 indicates whether the orders in request and response are same.
            bool isVerifyR4719 = (typeof(RopOpenFolderResponse) == ropResponses[0].GetType())
                                 && (typeof(RopCreateMessageResponse) == ropResponses[1].GetType())
                                 && (typeof(RopSaveChangesMessageResponse) == ropResponses[2].GetType());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4719,the actual value is {0}.", isVerifyR4719);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4719
            // True means the orders in request and response are same.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR4719,
                4719,
                @"[In Creating a ROP Output Buffer] The ROP responses in the ROP output buffer MUST be in the same order in which they were processed.");

            // The request order is openFolderRequest, createMessageRequest, saveChangesMessageRequest,
            // isVerifyR4565 indicates whether the orders in request and response are same.
            bool isVerifyR4565 = (typeof(RopOpenFolderResponse) == ropResponses[0].GetType())
                                 && (typeof(RopCreateMessageResponse) == ropResponses[1].GetType())
                                 && (typeof(RopSaveChangesMessageResponse) == ropResponses[2].GetType());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4565,the actual value is {0}.", isVerifyR4565);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4565
            // True means the orders in request and response are same.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR4565,
                4565,
                @"[In Creating a ROP Output Buffer]  The server MUST preserve the order of entries in the Server object handle table between the ROP input buffer and the ROP output buffer.");

            #endregion
        }

        #endregion
    }
}